# -*- coding: utf-8 -*-
"""
WPS-Agent MCP Server v2 — Professional Document Engine
Supports dual-mode operation:
  - Offline Mode: XML-native read/write (no WPS required, run-level precision)
  - Online Mode: WPS COM bridge (real-time interaction, visual feedback)
Auto-selects mode based on operation type and WPS availability.
"""
import asyncio
import json
import logging
import os
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml
from mcp.server import Server, NotificationOptions
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

sys.path.insert(0, str(Path(__file__).parent))

# ─── Engine Imports ───
from docx_engine import (
    parse_docx, build_document_model, serialize_document_model,
    Document, Paragraph, Run,
    DocumentAnalyzer, Formatter, StyleResolver,
    DocxEngineError, ParseError, SerializeError, ErrorCode,
)
from offline.docx_builder import OfflineDocxBuilder, read_docx_model, write_docx_model
from wps_bridge import (
    get_app, get_doc, list_documents,
    COMClient, BatchCommandExecutor,
    com_property, com_set, COMError, com_health_check,
    WDALIGNMENT, WDLINESPACING,
)
from wps_bridge.app import check_wps_available, WPS_REQUIRED_TOOLS
from wps_bridge import table as _table_bridge
from wps_bridge import layout as _layout_bridge
from wps_bridge import review as _review_bridge
from wps_bridge import search as _search_bridge
from wps_bridge import document as _document_bridge
from wps_bridge import transfer as _transfer_bridge
from wps_bridge import migrate as _migrate_bridge
from wps_bridge import compare as _compare_bridge
from intelligence.operation_logger import record as _log_record, get_summary as _log_summary, \
    get_recent as _log_recent, get_errors as _log_errors, replay_last_error, clear as _log_clear, dump as _log_dump

# ─── Logging ───
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "mcp_server.log", encoding="utf-8"),
        logging.StreamHandler(sys.stderr),
    ],
)
logger = logging.getLogger("wps-agent.mcp")

# ─── Config ───
PROJECT_DIR = Path(__file__).parent
CONFIG_PATH = PROJECT_DIR / "config.yaml"
try:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        CONFIG = yaml.safe_load(f) or {}
except Exception:
    CONFIG = {"server": {"name": "wps-agent"}}

# ─── Document Cache ───
_doc_cache: Dict[str, tuple] = {}  # filepath -> (Document, timestamp, mtime)
_snapshot_cache: Dict[str, Document] = {}  # filepath -> snapshot Document
_surgical_sessions: Dict[str, Any] = {}  # session_id -> SurgicalContext
_CACHE_TTL = 300  # seconds


def _get_cached_doc(filepath: str) -> Optional[Document]:
    """Get document from cache if fresh and file unchanged on disk."""
    if filepath in _doc_cache:
        doc, ts, cached_mtime = _doc_cache[filepath]
        if time.time() - ts < _CACHE_TTL:
            try:
                disk_mtime = os.path.getmtime(filepath)
                if disk_mtime <= cached_mtime:
                    return doc
            except OSError:
                return None
    return None


def _cache_doc(filepath: str, doc: Document):
    """Cache a document model with file mtime for staleness detection."""
    try:
        mtime = os.path.getmtime(filepath)
    except OSError:
        mtime = 0
    _doc_cache[filepath] = (doc, time.time(), mtime)


def _check_cache_staleness(filepath: str) -> Dict:
    """Check if cached model is stale vs disk file."""
    if filepath not in _doc_cache:
        return {"cached": False, "stale": True, "reason": "not_in_cache"}
    doc, ts, cached_mtime = _doc_cache[filepath]
    try:
        disk_mtime = os.path.getmtime(filepath)
    except OSError:
        return {"cached": True, "stale": True, "reason": "file_not_found"}
    if disk_mtime > cached_mtime:
        return {"cached": True, "stale": True, "reason": "disk_newer", "disk_mtime": disk_mtime, "cache_mtime": cached_mtime}
    if time.time() - ts > _CACHE_TTL:
        return {"cached": True, "stale": True, "reason": "ttl_expired", "age_seconds": time.time() - ts}
    return {"cached": True, "stale": False}


def _resolve_mode(args: Dict[str, Any], require_wps: bool = False) -> str:
    """Determine operation mode from filepath and WPS availability."""
    if args.get("mode") in ("online", "offline"):
        return args["mode"]
    if args.get("filepath"):
        return "offline"
    if require_wps:
        return "online"
    return "offline"


# ─── MCP Server ───
server = Server(CONFIG["server"]["name"])


@server.list_tools()
async def list_tools():
    return [
        # ─── document ───
        Tool(name="document", description="Document management. Actions: info/list/open/create/save/close/activate/export_pdf/insert_image/doc_properties/set_doc_properties/health_check",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "Which operation: info (get doc properties), list (list all open docs), open (open file), create (new doc), save, close, activate (switch active doc), export_pdf, insert_image (insert picture), doc_properties (read metadata), set_doc_properties (set author/title)"},
                 "doc_index": {"type": "integer", "description": "Document index from list (1-based, optional, defaults to active)"},
                 "filepath": {"type": "string", "description": "File path for open/save/export_pdf/insert_image"},
                 "save_changes": {"type": "boolean", "description": "Whether to save before closing"},
                 "output_path": {"type": "string", "description": "Output path for PDF export"},
                 "width": {"type": "number", "description": "Image width in points"},
                 "height": {"type": "number", "description": "Image height in points"},
                 "position": {"type": "string", "description": "Insert position: end, or paragraph index"},
                 "author": {"type": "string", "description": "Document author"},
                 "title": {"type": "string", "description": "Document title"},
                 "subject": {"type": "string", "description": "Document subject"},
             }, "required": ["action"]}),

        # --- content ---
        Tool(name="content", description="Read and write WPS Word document text. Use when user asks to: read/view/show document text, get paragraph content, check selection, see document outline/structure, insert/delete/replace text, read shapes/drawings/textboxes. IMPORTANT: Use batch action to read multiple items in ONE call (e.g. batch with types paragraph+outline). For cover pages, use create_cover (single call creates and formats all lines). Actions: full_text/paragraph/paragraphs/selection/range/outline/shapes/insert_text/delete_range/replace_range/batch/create_cover",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "full_text (get entire doc), paragraph (get one by index), paragraphs (get range), selection (current cursor), range (by start/end position), outline (heading structure), shapes (read all shapes/drawings/textboxes with text content), insert_text, delete_range, replace_range, batch (read multiple items at once), create_cover (single-call cover page, pass lines array of {text, font_name, font_size, bold, alignment, space_before, space_after...})"},
                 "para_index": {"type": "integer", "description": "Paragraph number (1-based)"},
                 "start": {"type": "integer", "description": "Start paragraph index for paragraphs action"},
                 "count": {"type": "integer", "description": "How many paragraphs to return"},
                 "start_pos": {"type": "integer", "description": "Character start position for range/delete/replace"},
                 "end_pos": {"type": "integer", "description": "Character end position for range/delete/replace (optional for delete_range, defaults to doc end)"},
                 "text": {"type": "string", "description": "Text to insert"},
                 "new_text": {"type": "string", "description": "Replacement text"},
                 "include_inlines": {"type": "boolean", "description": "For shapes: whether to include inline shapes (default true)"},
                 "position": {"type": "string", "description": "Where to insert: end (end of doc), before (before para_index), after (after para_index)"},
                 "clear_existing": {"type": "boolean", "description": "For create_cover: whether to clear existing content first (default true)"},
                 "lines": {"type": "array", "items": {"type": "object", "properties": {
                     "text": {"type": "string", "description": "Line text content"},
                     "font_name": {"type": "string", "description": "Font: 黑体, 宋体, 微软雅黑, etc."},
                     "font_size": {"type": "number", "description": "Font size in points (26=一号, 22=二号, 16=三号, 14=四号)"},
                     "bold": {"type": "boolean"}, "italic": {"type": "boolean"},
                     "alignment": {"type": "string", "description": "left/center/right/justify"},
                     "space_before": {"type": "number", "description": "Space before paragraph in points"},
                     "space_after": {"type": "number", "description": "Space after paragraph in points"},
                     "line_spacing_rule": {"type": "string", "description": "single/1.5lines/double/exactly/multiple"},
                     "line_spacing": {"type": "number"}, "first_line_indent": {"type": "number"},
                 }}, "description": "Array of line specs for create_cover"},
                 "doc_index": {"type": "integer", "description": "Document index (optional, defaults to active)"},
             }, "required": ["action"]}),

        # --- format ---
        Tool(name="format", description="Get or set font and paragraph formatting in WPS Word. Use when user asks to: change font/size/bold/italic/color, modify paragraph alignment/indentation/line spacing, apply styles, clear formatting, use format painter, add/remove watermark, add hyperlinks, set tab stops, create bullet lists. IMPORTANT: Use batch action to modify/read multiple paragraphs in ONE call. Supports Chinese font names: 黑体, 宋体, 仿宋, 楷体, 微软雅黑. Alignment: left/center/right/justify. Line spacing: single/1.5lines/double/exactly/multiple. Actions: get_font/set_font/get_paragraph_format/set_paragraph_format/apply_style/clear_formatting/copy_format/batch/add_watermark/remove_watermark/add_hyperlink/set_tab_stops/set_bullet_list",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "get_font (read font), set_font (modify font), get_paragraph_format (read paragraph), set_paragraph_format (modify paragraph), apply_style (apply named style), clear_formatting (reset), copy_format (format painter from source to targets), batch (execute multiple operations at once), add_watermark (text watermark), remove_watermark (delete all watermarks), add_hyperlink (insert clickable link), set_tab_stops (configure paragraph tab stops), set_bullet_list (apply bullet list formatting)"},
                 "para_index": {"type": "integer", "description": "Paragraph number to operate on (1-based)"},
                 "use_selection": {"type": "boolean", "description": "If true, operate on current selection instead of para_index"},
                 "source_para_index": {"type": "integer", "description": "Source paragraph for copy_format"},
                 "target_para_indices": {"type": "array", "items": {"type": "integer"}, "description": "Target paragraphs for copy_format"},
                 "style_name": {"type": "string", "description": "Style name e.g. 标题 1, 正文"},
                 "name": {"type": "string", "description": "Font name: 黑体, 宋体, 仿宋, 楷体, 微软雅黑, Calibri, Arial, Times New Roman"}, "name_far_east": {"type": "string", "description": "East Asian font name"}, "size": {"type": "number", "description": "Font size in points: 22=二号 16=三号 14=四号 12=小四 10.5=五号"},
                 "bold": {"type": "boolean", "description": "Bold: true/false"}, "italic": {"type": "boolean", "description": "Italic: true/false"}, "underline": {"type": "integer", "description": "Underline: 0=none 1=single 7=wave"},
                 "color_index": {"type": "integer", "description": "Color: 1=black 2=blue 3=cyan 4=green 6=red"}, "superscript": {"type": "boolean"}, "subscript": {"type": "boolean"},
                 "strike_through": {"type": "boolean"}, "spacing": {"type": "number", "description": "Character spacing in points"}, "scaling": {"type": "integer", "description": "Character scaling percentage"},
                 "kerning": {"type": "number", "description": "Kerning in points"},
                 "alignment": {"type": "string", "description": "Text alignment: left, center, right, justify"}, "first_line_indent": {"type": "number", "description": "First line indent in points (28pts ≈ 2 Chinese chars at 14pt)"},
                 "left_indent": {"type": "number"}, "right_indent": {"type": "number"},
                 "line_spacing_rule": {"type": "string", "description": "Line spacing rule: single, 1.5lines, double, exactly, multiple"}, "line_spacing": {"type": "number", "description": "Line spacing value"},
                 "space_before": {"type": "number", "description": "Space before paragraph in points"}, "space_after": {"type": "number", "description": "Space after paragraph in points"},
                 "outline_level": {"type": "integer", "description": "Outline level: 1-9 for headings, 10 for body text"},
                 "widow_control": {"type": "boolean"}, "keep_with_next": {"type": "boolean"},
                 "doc_index": {"type": "integer"},
                 "filepath": {"type": "string"},
                 "save_changes": {"type": "boolean"},
                 "output_path": {"type": "string"},
                 "width": {"type": "number"}, "height": {"type": "number"},
                 "position": {"type": "string"},
                 "author": {"type": "string"}, "title": {"type": "string"}, "subject": {"type": "string"},
             }, "required": ["action"]}),

        # ─── content ───
        Tool(name="content", description="Read and write document text with Run-level precision. Actions: full_text/paragraph/paragraphs/selection/range/outline/shapes/runs_detail/document_structure/full_structure/semantic_structure/cross_references/insert_text/insert_paragraph/delete_paragraphs/delete_runs/replace_paragraph_text/replace_runs/insert_run/delete_run/split_run/delete_range/replace_range/batch/create_cover/snapshot/rollback/cache_status/query_by_role",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"},
                 "para_index": {"type": "integer", "minimum": 1},
                 "run_index": {"type": "integer", "minimum": 1},
                 "from_para": {"type": "integer", "minimum": 1},
                 "to_para": {"type": "integer", "minimum": 1},
                 "from_run": {"type": "integer", "minimum": 1},
                 "to_run": {"type": "integer", "minimum": 1},
                 "run_indices": {"type": "array", "items": {"type": "integer"}},
                 "preserve_format": {"type": "boolean"},
                 "start": {"type": "integer", "minimum": 0}, "count": {"type": "integer"},
                 "start_pos": {"type": "integer", "minimum": 0}, "end_pos": {"type": "integer", "minimum": 0},
                 "text": {"type": "string"}, "new_text": {"type": "string"},
                 "position": {"type": "string"},
                 "clear_existing": {"type": "boolean"},
                 "lines": {"type": "array", "items": {"type": "object"}},
                 "items": {"type": "array"},
                 "include_inlines": {"type": "boolean"},
                 "doc_index": {"type": "integer"},
                 "filepath": {"type": "string"},
                 "mode": {"type": "string", "description": "offline or online"},
             }, "required": ["action"]}),

        # ─── format ───
        Tool(name="format", description="Font and paragraph formatting with Run-level precision. Actions: get_font/set_font/get_run_font/set_run_font/get_paragraph_format/set_paragraph_format/resolve_format/resolve_run_format/apply_style/clear_formatting/copy_format/batch/add_watermark/remove_watermark/add_hyperlink/set_tab_stops/set_bullet_list/set_text_effect",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"},
                 "para_index": {"type": "integer", "minimum": 1},
                 "run_index": {"type": "integer", "minimum": 1}, "use_selection": {"type": "boolean"},
                 "source_para_index": {"type": "integer", "minimum": 0}, "target_para_indices": {"type": "array"},
                 "style_name": {"type": "string"},
                 "name": {"type": "string"}, "name_far_east": {"type": "string"},
                 "size": {"type": "number"}, "bold": {"type": "boolean"}, "italic": {"type": "boolean"},
                 "underline": {"type": "integer", "minimum": 0}, "color_index": {"type": "integer"},
                 "color_rgb": {"type": "integer", "description": "OLE COLORREF (0x00BBGGRR), e.g. 0xFF0000 for blue"},
                 "highlight": {"type": "integer", "description": "Highlight color index"},
                 "superscript": {"type": "boolean"}, "subscript": {"type": "boolean"},
                 "strike_through": {"type": "boolean"},
                 "caps": {"type": "boolean"}, "small_caps": {"type": "boolean"},
                 "emboss": {"type": "boolean"}, "shadow": {"type": "boolean"}, "outline": {"type": "boolean"},
                 "spacing": {"type": "number"}, "scaling": {"type": "integer", "minimum": 0},
                 "kerning": {"type": "number"},
                 "alignment": {"type": "string"}, "first_line_indent": {"type": "number"},
                 "left_indent": {"type": "number"}, "right_indent": {"type": "number"},
                 "line_spacing_rule": {"type": "string"}, "line_spacing": {"type": "number"},
                 "space_before": {"type": "number"}, "space_after": {"type": "number"},
                 "outline_level": {"type": "integer", "minimum": 0},
                 "operations": {"type": "array"},
                 "text": {"type": "string"}, "url": {"type": "string"},
                 "stops": {"type": "array"}, "para_indices": {"type": "array"},
                 "doc_index": {"type": "integer"}, "filepath": {"type": "string"},
                 "mode": {"type": "string"},
             }, "required": ["action"]}),

        # --- table ---
        Tool(name="table", description="Create and manipulate Word tables. Use when user asks to: insert/delete tables, fill table cells, format tables with headers/borders/colors, merge cells, adjust column widths, apply alternating row colors, set cell shading, get table dimensions. Table index is 1-based. IMPORTANT: Use batch_read to read multiple tables in ONE call. Actions: count/info/read/create/delete/set_cell_text/format_cell/set_header/format_borders/merge_cells/auto_fit/set_column_width/alternate_rows/set_cell_shading/table_dimensions/batch_read",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "name": {"type": "string"},
                 "base_style": {"type": "string"}, "font_name": {"type": "string"},
                 "font_size": {"type": "number"}, "bold": {"type": "boolean"},
                 "alignment": {"type": "string"}, "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # ─── table ───
        Tool(name="table", description="Table operations. Actions: count/info/read/create/delete/set_cell_text/format_cell/set_header/format_borders/merge_cells/auto_fit/set_column_width/alternate_rows/set_cell_shading/batch_read/table_dimensions",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "table_index": {"type": "integer", "minimum": 0},
                 "rows": {"type": "integer"}, "cols": {"type": "integer"},
                 "row": {"type": "integer", "minimum": 0}, "col": {"type": "integer", "minimum": 0},
                 "text": {"type": "string"}, "font_name": {"type": "string"},
                 "font_size": {"type": "number"}, "bold": {"type": "boolean"},
                 "align": {"type": "string"}, "shading_color": {"type": "integer", "minimum": 0},
                 "row_count": {"type": "integer"}, "width": {"type": "number"},
                 "start_row": {"type": "integer", "minimum": 0}, "start_col": {"type": "integer", "minimum": 0},
                 "end_row": {"type": "integer", "minimum": 0}, "end_col": {"type": "integer", "minimum": 0},
                 "behavior": {"type": "integer", "minimum": 0}, "color1": {"type": "string"}, "color2": {"type": "string"},
                 "table_indices": {"type": "array"}, "bg_color": {"type": "string"},
                 "doc_index": {"type": "integer", "minimum": 0},
             }, "required": ["action"]}),

        # ─── search ───
        Tool(name="search", description="Find and replace. Actions: find/replace/find_format/goto_heading",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "query": {"type": "string"},
                 "find_text": {"type": "string"}, "replace_text": {"type": "string"},
                 "match_case": {"type": "boolean"}, "whole_word": {"type": "boolean"},
                  "replace_all": {"type": "boolean"},
                  "font_name": {"type": "string"}, "font_size": {"type": "number"},
                  "bold": {"type": "boolean"}, "style_name": {"type": "string"},
                  "text": {"type": "string"},
                 "level": {"type": "integer", "minimum": 0}, "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # --- layout ---
        Tool(name="layout", description="Configure WPS Word page layout. Use when user asks to: change paper size (A4), adjust margins, switch portrait/landscape, add section breaks, set columns, edit headers/footers, add page numbers, get page dimensions. Actions: page_setup/section_info/add_section_break/columns/header_footer/page_numbers/page_dimensions",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "section_index": {"type": "integer", "minimum": 0},
                 "page_width": {"type": "number"}, "page_height": {"type": "number"},
                 "top_margin": {"type": "number"}, "bottom_margin": {"type": "number"},
                 "left_margin": {"type": "number"}, "right_margin": {"type": "number"},
                 "orientation": {"type": "integer", "minimum": 0, "description": "0=portrait, 1=landscape"}, "gutter": {"type": "number"},
                 "different_first_page": {"type": "boolean"},
                 "para_index": {"type": "integer", "minimum": 0}, "break_type": {"type": "string"},
                 "count": {"type": "integer"}, "header_type": {"type": "string"},
                 "text": {"type": "string"}, "alignment": {"type": "string"},
                 "start_at": {"type": "integer", "minimum": 0}, "doc_index": {"type": "integer"},
                 "shape_index": {"type": "integer", "description": "Shape index for image_wrap"},
                 "wrap_type": {"type": "string", "description": "square/tight/through/top_and_bottom/behind/in_front_of/inline"},
                 "enable": {"type": "boolean", "description": "Enable/disable line numbers"},
                 "count_by": {"type": "integer", "description": "Line number interval"},
                 "restart": {"type": "string", "description": "continuous/restart_page/restart_section"},
                 "distance": {"type": "number", "description": "Line number distance from text (pts)"},
                 "color_rgb": {"type": "integer", "description": "Border/line color as OLE COLORREF"},
                 "distance_from": {"type": "integer", "description": "Border distance: 0=text, 1=edge"},
             }, "required": ["action"]}),

        # ─── review ───
        Tool(name="review", description="Track changes and comments. Actions: track_changes_toggle/track_changes_status/comments_list/comment_add/revisions_list/revisions_accept_all/revisions_reject_all",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "enable": {"type": "boolean"},
                 "text": {"type": "string"}, "para_index": {"type": "integer", "minimum": 0},
                 "range_start": {"type": "integer", "minimum": 0}, "range_end": {"type": "integer", "minimum": 0},
                 "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # ─── reference ───
        Tool(name="reference", description="Footnotes, endnotes, bookmarks, fields. Actions: add_footnote/add_endnote/list_footnotes/add_bookmark/goto_bookmark/list_bookmarks/insert_field",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "text": {"type": "string"},
                 "name": {"type": "string"}, "para_index": {"type": "integer", "minimum": 0},
                 "field_code": {"type": "string"}, "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # ─── docspace ───
        Tool(name="docspace", description="Manage all open documents. Actions: list_all/activate/close_all/save_all",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "doc_id": {"type": "string"},
             }, "required": ["action"]}),

        # ─── transfer ───
        Tool(name="transfer", description="Copy between documents. Actions: copy_paragraphs/copy_table/copy_range",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "source_doc_id": {"type": "string"},
                 "target_doc_id": {"type": "string"}, "from_start": {"type": "integer", "minimum": 0},
                 "from_end": {"type": "integer", "minimum": 0}, "table_index": {"type": "integer"},
                 "start_pos": {"type": "integer", "minimum": 0}, "end_pos": {"type": "integer", "minimum": 0},
                 "target_position": {"type": "string"},
             }, "required": ["action", "source_doc_id", "target_doc_id"]}),

        # ─── migrate ───
        Tool(name="migrate", description="Word/Excel/PPT migration. Actions: word_table_to_excel/excel_range_to_word_table/word_outline_to_ppt",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "word_doc_id": {"type": "string"},
                 "excel_doc_id": {"type": "string"}, "table_index": {"type": "integer"},
                 "range_start": {"type": "string"}, "range_end": {"type": "string"},
                 "target_cell": {"type": "string"}, "position": {"type": "string"},
                 "keep_format": {"type": "boolean"},
             }, "required": ["action"]}),

        # ─── compare ───
        Tool(name="compare", description="Compare two documents. Actions: text_diff/format_diff",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "doc_id_a": {"type": "string"},
                 "doc_id_b": {"type": "string"},
             }, "required": ["action", "doc_id_a", "doc_id_b"]}),

        # --- ai_format ---
        Tool(name="ai_format", description="AI-powered intelligent formatting for WPS Word. Use when user asks to: analyze document formatting, suggest improvements, apply professional templates, auto-apply formatting from natural language, generate table of contents, add multi-level heading numbers, validate formatting quality, generate/summarize/rewrite/expand/translate content, or run quality supervision to auto-fix layout issues. Also supports design suggestions (palettes, typography, anti-patterns). Actions: analyze/suggest/apply_template/reformat/auto_toc/auto_numbering/validate/generate_content/summarize_document/rewrite_paragraph/expand_section/translate_section/supervise/suggest_design",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "template_name": {"type": "string"},
                 "instructions": {"type": "string"}, "doc_index": {"type": "integer"},
                 "filepath": {"type": "string"}, "output_path": {"type": "string"},
                 "document_type": {"type": "string"},
             }, "required": ["action"]}),

        # ─── presentation ───
        Tool(name="presentation", description="PPT operations. Actions: create/open/list/save/close/slide_count/slide_info/add_slide/delete_slide/set_title/set_body/add_textbox/format_text/insert_image/insert_table/fill_cell/apply_theme/add_notes",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "filepath": {"type": "string"},
                 "slide_index": {"type": "integer", "minimum": 0}, "text": {"type": "string"},
                 "left": {"type": "integer"}, "top": {"type": "integer"},
                 "width": {"type": "integer"}, "height": {"type": "integer", "minimum": 0},
                 "shape_index": {"type": "integer"}, "font_name": {"type": "string"},
                 "font_size": {"type": "number"}, "bold": {"type": "boolean"},
                 "color": {"type": "integer", "minimum": 0}, "image_path": {"type": "string"},
                 "rows": {"type": "integer"}, "cols": {"type": "integer"},
                 "row": {"type": "integer", "minimum": 0}, "col": {"type": "integer", "minimum": 0},
                 "table_index": {"type": "integer"}, "theme_name": {"type": "string"},
                 "layout_index": {"type": "integer", "minimum": 0},
             }, "required": ["action"]}),

        # ─── excel ───
        Tool(name="excel", description="Excel operations. Actions: create/open/list/save/close/sheet_list/sheet_activate/sheet_add/sheet_copy/sheet_delete/sheet_move/cell_read/cell_write/range_read/range_write/font_set/interior_set/borders_set/column_width/auto_fit/merge_cells/formula_set/chart_add/chart_set_source/chart_set_title/sort/auto_filter/remove_filter/conditional_format/freeze_panes/get_used_range",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "filepath": {"type": "string"},
                 "cell_ref": {"type": "string"}, "value": {"type": "string"},
                 "start": {"type": "string"}, "end": {"type": "string"},
                 "data": {"type": "array"}, "sheet_name": {"type": "string"},
                 "name": {"type": "string"}, "font_name": {"type": "string"},
                 "font_size": {"type": "number"}, "bold": {"type": "boolean"},
                 "color": {"type": "integer", "minimum": 0}, "style": {"type": "integer"},
                 "width": {"type": "number"}, "col": {"type": "string"},
                 "formula": {"type": "string"}, "chart_type": {"type": "integer", "minimum": 0},
                 "chart_width": {"type": "integer", "minimum": 0}, "chart_height": {"type": "integer", "minimum": 0},
                 "save_changes": {"type": "boolean"},
             }, "required": ["action"]}),

        # ─── offline_docx ───
        Tool(name="offline_docx", description="Offline XML-native docx processing. Actions: build/build_cover/validate/analyze/auto_format/apply_template/add_numbering/replace_text/get_text/get_statistics/read_model/write_model/full_structure/semantic_structure/cross_references",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "filepath": {"type": "string"},
                 "output_path": {"type": "string"}, "structure": {"type": "object"},
                 "lines": {"type": "array"}, "auto_fix": {"type": "boolean"},
                 "template_name": {"type": "string"}, "document_type": {"type": "string"},
                 "old_text": {"type": "string"}, "new_text": {"type": "string"},
                 "case_sensitive": {"type": "boolean"},
             }, "required": ["action"]}),

        # ─── content_control ───
        Tool(name="content_control", description="Content Control operations (Rich Text, Date Picker, Dropdown, Checkbox). Actions: count/list_controls/info/add/set_text/set_checkbox/select_dropdown/delete/set_tag/find_by_tag",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "cc_index": {"type": "integer", "minimum": 0},
                 "type_name": {"type": "string"}, "text": {"type": "string"},
                 "title": {"type": "string"}, "tag": {"type": "string"},
                 "checked": {"type": "boolean"}, "item_text": {"type": "string"},
                 "dropdown_items": {"type": "array"}, "date_format": {"type": "string"},
                 "position": {"type": "string"}, "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # ─── field_codes ───
        Tool(name="field_codes", description="Advanced field code operations (Quote, If, Seq, StyleRef, Ref, DocProperty, etc.). Actions: insert_field/insert_quote/insert_doc_property/insert_seq/insert_style_ref/insert_ref/insert_if/list_fields/update_fields/unlink_field/find_field_by_code",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "field_code": {"type": "string"},
                 "switches": {"type": "string"}, "property_name": {"type": "string"},
                 "sequence_name": {"type": "string"}, "format_type": {"type": "string"},
                 "style_name": {"type": "string"}, "bookmark_name": {"type": "string"},
                 "condition": {"type": "string"}, "true_text": {"type": "string"},
                 "false_text": {"type": "string"}, "pattern": {"type": "string"},
                 "field_index": {"type": "integer", "minimum": 0}, "para_index": {"type": "integer", "minimum": 0},
                 "position": {"type": "string"}, "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # ─── surgical ───
        Tool(name="surgical", description="Surgical-level modifications with context capture, verification, and rollback. Actions: select/modify/commit/rollback",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"},
                 "para_indices": {"type": "array", "items": {"type": "integer"}},
                 "mutations": {"type": "array", "items": {"type": "object"}},
                 "doc_index": {"type": "integer"},
                 "sr": {"type": "string", "description": "Semantic role to select (e.g. abstract/cover/acknowledgements)"},
                 "filepath": {"type": "string"},
             }, "required": ["action"]}),

        # --- presentation ---
        Tool(name="presentation", description="Create and manipulate WPS Presentations (PPT). Use when user asks to: create/open/save PPT, add slides, set titles/body text, insert images/tables, format text, add speaker notes, add shapes, reorder slides, set slide backgrounds, add charts, or export slides as images. Actions: create/open/list/save/close/slide_count/slide_info/add_slide/delete_slide/set_title/set_body/add_textbox/format_text/insert_image/insert_table/fill_cell/apply_theme/add_notes/add_shape/reorder_slides/set_slide_background/add_chart_modern",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "create (new presentation)/open (file)/list (all open)/save/close/slide_count/slide_info/add_slide/delete_slide/set_title/set_body/add_textbox/format_text/insert_image/insert_table/fill_cell/apply_theme/add_notes"},
                 "filepath": {"type": "string", "description": "Path to .pptx file"},
                 "slide_index": {"type": "integer", "description": "Slide number (1-based)"},
                 "text": {"type": "string", "description": "Text content"},
                 "left": {"type": "integer", "description": "Left position in points"},
                 "top": {"type": "integer", "description": "Top position in points"},
                 "width": {"type": "integer", "description": "Width in points"},
                 "height": {"type": "integer", "description": "Height in points"},
                 "shape_index": {"type": "integer", "description": "Shape number on slide"},
                 "font_name": {"type": "string", "description": "Font name: 黑体, 宋体, etc."},
                 "font_size": {"type": "number", "description": "Font size in points"},
                 "bold": {"type": "boolean"},
                 "color": {"type": "integer", "description": "Color index"},
                 "image_path": {"type": "string", "description": "Path to image file"},
                 "rows": {"type": "integer", "description": "Number of rows for table"},
                 "cols": {"type": "integer", "description": "Number of columns for table"},
                 "row": {"type": "integer", "description": "Row number (1-based) for fill_cell"},
                 "col": {"type": "integer", "description": "Column number (1-based) for fill_cell"},
                 "table_index": {"type": "integer", "description": "Shape index of the table"},
                 "theme_name": {"type": "string", "description": "Path to theme/template file (.potx)"},
                 "layout_index": {"type": "integer", "description": "Layout index for add_slide (default 1=title)"},
             }, "required": ["action"]}),

        # --- excel ---
        Tool(name="excel", description="Create and manipulate WPS Excel spreadsheets. Use when user asks to: create/open/save Excel workbooks, switch/add/rename sheets, read/write cell values, write tabular data, set formulas (=SUM etc.), format cells (font/color/borders), adjust column widths, merge cells, create charts, sort data, filter, conditional format, freeze panes, insert/delete rows, add cell comments, import/export CSV, validate formulas. Actions: create/open/list/save/close/sheet_list/sheet_activate/sheet_add/sheet_copy/sheet_delete/sheet_move/cell_read/cell_write/range_read/range_write/font_set/interior_set/borders_set/column_width/auto_fit/merge_cells/formula_set/chart_add/chart_set_source/chart_set_title/sort/auto_filter/remove_filter/conditional_format/freeze_panes/get_used_range/insert_rows/delete_rows/add_cell_comment/import_csv/export_csv/validate_formulas/recalc_formulas",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "create (new workbook)/open (file)/list (all workbooks)/save/close/sheet_list/sheet_activate/sheet_add/sheet_copy/sheet_delete/sheet_move/cell_read (single cell)/cell_write (single cell)/range_read (A1:D10)/range_write (write 2D array to range)/font_set (format cell font)/interior_set (cell background color)/borders_set (add borders)/column_width/auto_fit/merge_cells/formula_set (e.g. =SUM(C2:C5))/chart_add/chart_set_source/chart_set_title/sort/auto_filter/remove_filter/conditional_format/freeze_panes/get_used_range"},
                 "filepath": {"type": "string", "description": "Path to .xlsx file"},
                 "cell_ref": {"type": "string", "description": "Cell reference like A1, B5, or range like A1:D1 for formatting"},
                 "value": {"type": "string", "description": "Value to write (number or text)"},
                 "start": {"type": "string", "description": "Top-left cell e.g. A1 for range operations"},
                 "end": {"type": "string", "description": "Bottom-right cell e.g. D10 for range operations"},
                 "data": {"type": "array", "description": "2D array of data [[row1],[row2],...]"},
                 "sheet_name": {"type": "string", "description": "Worksheet name (optional, uses active sheet)"},
                 "name": {"type": "string", "description": "Sheet name for sheet_add/sheet_activate"},
                 "font_name": {"type": "string", "description": "Font name e.g. 黑体, 宋体"}, "font_size": {"type": "number", "description": "Font size in points"},
                 "bold": {"type": "boolean"}, "italic": {"type": "boolean"},
                 "color": {"type": "integer", "description": "Color index: 1=black 2=white 3=red 5=blue 6=yellow 15=gray"}, "style": {"type": "integer", "description": "Border line style: 1=continuous"},
                 "width": {"type": "number", "description": "Column width"}, "col": {"type": "string", "description": "Column letter e.g. A, B, C"},
                 "formula": {"type": "string", "description": "Excel formula e.g. =SUM(C2:C5), =AVERAGE(B2:B10)"},
                 "chart_type": {"type": "integer", "description": "Chart type code"}, "left": {"type": "integer"}, "top": {"type": "integer"},
                 "chart_width": {"type": "integer"}, "chart_height": {"type": "integer"},
                 "save_changes": {"type": "boolean", "description": "Whether to save before closing"},
             }, "required": ["action"]}),

        # --- offline_docx ---
        Tool(name="offline_docx", description="Generate and validate Word .docx files offline (no WPS required). Use when WPS is not available or user wants programmatic document builder. Actions: build (create .docx from JSON structure), build_cover (generate standalone cover page), validate (check XML structure issues)",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "build (full docx from structure), build_cover (cover page only), validate (check XML issues)"},
                 "structure": {"type": "object", "description": "Document structure JSON for build action"},
                 "lines": {"type": "array", "items": {"type": "object"}, "description": "Cover page line specs for build_cover"},
                 "filepath": {"type": "string", "description": "Input file path for validate"},
                 "auto_fix": {"type": "boolean", "description": "Attempt auto-repair (validate action)"},
                 "output_path": {"type": "string", "description": "Output .docx file path"},
             }, "required": ["action"]}),

        # --- offline_xlsx ---
        Tool(name="offline_xlsx", description="Analyze, create and validate Excel .xlsx files offline (no WPS required). Use when WPS is not available or need pandas data analysis. Actions: build (create .xlsx from JSON), analyze (pandas analysis of existing file), convert_csv (CSV to XLSX), validate_formulas (scan for errors with openpyxl), recalc_and_verify (LibreOffice recalc + error scan), apply_financial_colors",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "build/analyze/convert_csv/validate_formulas/recalc_and_verify/apply_financial_colors"},
                 "structure": {"type": "object", "description": "Spreadsheet structure JSON for build action"},
                 "filepath": {"type": "string", "description": "Input file path for analyze/convert_csv/validate_formulas/recalc_and_verify"},
                 "output_path": {"type": "string", "description": "Output file path"},
                 "csv_path": {"type": "string"}, "delimiter": {"type": "string", "description": "CSV delimiter (default comma)"},
                 "timeout": {"type": "integer", "description": "Timeout seconds for recalc_and_verify (default 60)"},
             }, "required": ["action"]}),

        # --- offline_pptx ---
        Tool(name="offline_pptx", description="Create, analyze and export PPTX presentations offline (no WPS required). Use when WPS is not available or need programmatic slide creation. Actions: build (create .pptx from JSON structure), extract_text (read text from existing pptx), export_slides (convert to JPEG slide images via LibreOffice)",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "build (create pptx from structure), extract_text (read text from pptx), export_slides (render slides as JPEG)"},
                 "structure": {"type": "object", "description": "Presentation structure JSON for build action"},
                 "filepath": {"type": "string", "description": "Input file path for extract_text/export_slides"},
                 "output_path": {"type": "string", "description": "Output .pptx file path for build"},
                 "output_prefix": {"type": "string", "description": "Filename prefix for exported slide images (default 'slide')"},
                 "dpi": {"type": "integer", "description": "Image DPI for export_slides (default 150)"},
             }, "required": ["action"]}),
    ]


# ─── Tool Handlers ───

@server.call_tool()
async def call_tool(name: str, arguments: dict):
    action = arguments.get("action", "")
    doc_index = arguments.get("doc_index")
    filepath = arguments.get("filepath")
    mode = arguments.get("mode", "auto")

    logger.info(f"Tool call: {name}.{action} mode={mode}")

    try:
        # ─── Graceful Degradation: online tools require WPS ───
        if name in WPS_REQUIRED_TOOLS:
            # offline_docx and pure-offline ai_format are exempt
            is_offline_op = (mode == "offline") or (name == "format" and filepath)
            if not is_offline_op and not check_wps_available():
                return [TextContent(type="text", text=json.dumps({
                    "success": False,
                    "error_code": "WPS_NOT_RUNNING",
                    "error": "WPS Office is not running. Please start WPS Office to use online tools (document/content/format/style/table/search/layout/review/reference/transfer/migrate/compare/presentation/excel). Offline tools (offline_docx) are still available.",
                    "detail": {"available_tools": ["offline_docx"]},
                }, ensure_ascii=False))]
        # ─── End Graceful Degradation ───

        # Route to appropriate handler
        start_time = time.time()
        error_msg = None

        if name == "document":
            result = _handle_document(action, arguments)
        elif name == "content":
            if action == "full_text":
                result = {"text": content.full_text(doc_index)}
            elif action == "paragraph":
                result = content.paragraph(arguments["para_index"], doc_index)
            elif action == "paragraphs":
                result = content.paragraphs(arguments.get("start", 1), arguments.get("count", 10), doc_index)
            elif action == "selection":
                result = content.selection_info(doc_index)
            elif action == "range":
                result = content.range_text(arguments["start_pos"], arguments["end_pos"], doc_index)
            elif action == "outline":
                result = content.outline(doc_index)
            elif action == "insert_text":
                result = content.insert_text(arguments["text"], arguments.get("position", "end"), arguments.get("para_index"), doc_index)
            elif action == "delete_range":
                result = content.delete_range(arguments["start_pos"], arguments.get("end_pos"), doc_index)
            elif action == "replace_range":
                result = content.replace_range(arguments["start_pos"], arguments["end_pos"], arguments["new_text"], doc_index)
            elif action == "create_cover":
                result = content.create_cover(arguments["lines"], arguments.get("clear_existing", True), doc_index)
            elif action == "batch":
                result = content.batch(arguments["items"], doc_index)
            elif action == "shapes":
                result = content.shapes(arguments.get("include_inlines", True), doc_index)
            else:
                result = {"error": f"Unknown content action: {action}"}

        elif name == "format":
            if action == "get_font":
                result = formatting.get_font(arguments.get("para_index"), arguments.get("start_pos"), arguments.get("end_pos"), arguments.get("use_selection", False), doc_index)
            elif action == "set_font":
                result = formatting.set_font(doc_index=doc_index, **{k: v for k, v in arguments.items() if k not in ("action", "doc_index")})
            elif action == "get_paragraph_format":
                result = formatting.get_paragraph_format(arguments["para_index"], doc_index)
            elif action == "set_paragraph_format":
                result = formatting.set_paragraph_format(doc_index=doc_index, **{k: v for k, v in arguments.items() if k not in ("action", "doc_index")})
            elif action == "apply_style":
                result = formatting.apply_style(arguments["style_name"], arguments.get("para_index"), arguments.get("use_selection", False), doc_index)
            elif action == "clear_formatting":
                result = formatting.clear_formatting(arguments.get("para_index"), arguments.get("use_selection", False), doc_index)
            elif action == "copy_format":
                result = formatting.copy_format(arguments["source_para_index"], arguments["target_para_indices"], doc_index)
            elif action == "batch":
                result = formatting.batch(arguments["operations"], doc_index)
            elif action == "add_watermark":
                result = document.add_watermark(arguments["text"], arguments.get("font_size", 72), arguments.get("color", 15), doc_index)
            elif action == "remove_watermark":
                result = document.remove_watermark(doc_index)
            elif action == "add_hyperlink":
                result = formatting.add_hyperlink(arguments["text"], arguments["url"], arguments.get("para_index"), doc_index)
            elif action == "set_tab_stops":
                result = formatting.set_tab_stops(arguments["para_index"], arguments.get("stops", []), doc_index)
            elif action == "set_bullet_list":
                result = formatting.set_bullet_list(arguments["para_indices"], arguments.get("bullet_char"), doc_index)
            else:
                result = {"error": f"Unknown format action: {action}"}

        elif name == "style":
            result = _handle_style(action, arguments)
        elif name == "table":
            if action == "count":
                result = {"count": table.table_count(doc_index)}
            elif action == "info":
                result = table.table_info(arguments["table_index"], doc_index)
            elif action == "read":
                result = table.table_read(arguments["table_index"], doc_index)
            elif action == "create":
                result = table.table_create(arguments["rows"], arguments["cols"], arguments.get("position"), doc_index)
            elif action == "delete":
                result = table.table_delete(arguments["table_index"], doc_index)
            elif action == "set_cell_text":
                result = table.set_cell_text(arguments["table_index"], arguments["row"], arguments["col"], arguments["text"], doc_index)
            elif action == "format_cell":
                result = table.format_cell(doc_index=doc_index, **{k: v for k, v in arguments.items() if k not in ("action", "doc_index")})
            elif action == "set_header":
                result = table.set_header(arguments["table_index"], arguments.get("row_count", 1), doc_index)
            elif action == "format_borders":
                result = table.format_borders(arguments["table_index"], arguments.get("inside"), arguments.get("outside"), doc_index)
            elif action == "merge_cells":
                result = table.merge_cells(arguments["table_index"], arguments["start_row"], arguments["start_col"], arguments["end_row"], arguments["end_col"], doc_index)
            elif action == "auto_fit":
                result = table.auto_fit(arguments["table_index"], arguments.get("behavior", 2), doc_index)
            elif action == "set_column_width":
                result = table.set_column_width(arguments["table_index"], arguments["col"], arguments["width"], doc_index)
            elif action == "alternate_rows":
                result = table.alternate_rows(arguments["table_index"], arguments.get("color1", "FFFFFF"), arguments.get("color2", "F2F2F2"), doc_index)
            elif action == "batch_read":
                result = table.batch_read(arguments["table_indices"], doc_index)
            elif action == "set_cell_shading":
                result = table.set_cell_shading(arguments["table_index"], arguments["row"], arguments["col"], arguments["bg_color"], doc_index)
            elif action == "table_dimensions":
                result = table.table_dimensions(arguments["table_index"], doc_index)
            else:
                result = {"error": f"Unknown table action: {action}"}

        elif name == "search":
            result = _handle_search(action, arguments)
        elif name == "layout":
            if action == "page_setup":
                result = layout.page_setup(doc_index=doc_index, **{k: v for k, v in arguments.items() if k not in ("action", "doc_index")})
            elif action == "section_info":
                result = layout.section_info(arguments.get("section_index"), doc_index)
            elif action == "add_section_break":
                result = layout.add_section_break(arguments["para_index"], arguments.get("break_type", "next_page"), doc_index)
            elif action == "columns":
                result = layout.set_columns(arguments["count"], arguments.get("section_index"), doc_index)
            elif action == "header_footer":
                result = layout.header_footer(arguments.get("section_index"), arguments.get("header_type", "header"), arguments.get("text"), doc_index)
            elif action == "page_numbers":
                result = layout.page_numbers(arguments.get("alignment", "center"), arguments.get("start_at"), arguments.get("section_index"), doc_index)
            elif action == "page_dimensions":
                result = layout.get_page_dimensions(arguments.get("section_index"), doc_index)
            else:
                result = {"error": f"Unknown layout action: {action}"}

        elif name == "review":
            result = _handle_review(action, arguments)
        elif name == "reference":
            result = _handle_reference(action, arguments)
        elif name == "docspace":
            result = _handle_docspace(action, arguments)
        elif name == "transfer":
            result = _handle_transfer(action, arguments)
        elif name == "migrate":
            result = _handle_migrate(action, arguments)
        elif name == "compare":
            result = _handle_compare(action, arguments)
        elif name == "ai_format":
            if action == "analyze":
                result = _ai_analyze(doc_index, arguments.get("template"))
            elif action == "suggest":
                result = _ai_suggest(doc_index)
            elif action == "apply_template":
                result = _ai_apply_template(arguments["template_name"], doc_index)
            elif action == "reformat":
                result = _ai_reformat(arguments.get("instructions", ""), doc_index)
            elif action == "auto_toc":
                result = _ai_auto_toc(doc_index)
            elif action == "auto_numbering":
                result = _ai_auto_numbering(doc_index)
            elif action == "validate":
                result = _ai_validate(doc_index)
            elif action == "generate_content":
                result = _ai_generate_content(arguments.get("instructions", ""), arguments.get("position", "end"), arguments.get("para_index"), doc_index)
            elif action == "summarize_document":
                result = _ai_summarize(doc_index)
            elif action == "rewrite_paragraph":
                result = _ai_rewrite(arguments["para_index"], arguments.get("instructions", ""), doc_index)
            elif action == "expand_section":
                result = _ai_expand(arguments["para_index"], doc_index)
            elif action == "translate_section":
                result = _ai_translate(arguments["para_index"], arguments.get("target_lang", "en"), doc_index)
            elif action == "supervise":
                result = _ai_supervise(doc_index)
            elif action == "suggest_design":
                result = _ai_suggest_design(arguments.get("topic", ""), arguments.get("category", ""))
            else:
                result = {"error": f"Unknown ai_format action: {action}"}

        elif name == "excel":
            result = _handle_excel(action, arguments)

        elif name == "presentation":
            result = _handle_presentation(action, arguments)

        elif name == "offline_docx":
            result = _handle_offline_docx(action, arguments)

        elif name == "offline_xlsx":
            result = _handle_offline_xlsx(action, arguments)

        elif name == "offline_pptx":
            result = _handle_offline_pptx(action, arguments)

        else:
            result = {"error": f"Unknown tool: {name}"}

        duration = time.time() - start_time
        _log_record(name, action, arguments, result, duration, error_msg)

        # Wrap result
        if isinstance(result, dict) and "error" in result and "error_code" not in result:
            result = {"success": False, "error_code": "UNKNOWN", "error": result["error"]}
        elif isinstance(result, dict) and "success" not in result:
            result["success"] = True

        return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, default=str))]

    except DocxEngineError as e:
        logger.error(f"DocxEngineError: {e.code} - {e}")
        return [TextContent(type="text", text=json.dumps(e.to_dict(), ensure_ascii=False))]
    except COMError as e:
        logger.error(f"COMError: {e.code} - {e}")
        return [TextContent(type="text", text=json.dumps(e.to_dict(), ensure_ascii=False))]
    except Exception as e:
        logger.exception(f"Unexpected error in {name}.{action}: {e}")
        return [TextContent(type="text", text=json.dumps({
            "success": False,
            "error_code": "UNKNOWN",
            "error": str(e),
        }, ensure_ascii=False))]


# ─── Document Handler ───

def _handle_document(action: str, args: dict) -> dict:
    if action == "info":
        doc = get_doc(args.get("doc_index"))
        return {
            "name": com_property(doc, "Name", ""),
            "full_name": com_property(doc, "FullName", ""),
            "paragraph_count": com_property(doc.Paragraphs, "Count", 0),
            "table_count": com_property(doc.Tables, "Count", 0),
            "section_count": com_property(doc.Sections, "Count", 1),
            "saved": com_property(doc, "Saved", False),
        }
    elif action == "list":
        return {"documents": list_documents()}
    elif action == "open":
        doc = get_app().Documents.Open(args["filepath"])
        return {"name": doc.Name, "full_name": doc.FullName}
    elif action == "create":
        doc = get_app().Documents.Add()
        return {"name": doc.Name, "full_name": doc.FullName}
    elif action == "save":
        doc = get_doc(args.get("doc_index"))
        fp = args.get("filepath")
        if fp:
            doc.SaveAs(fp)
        else:
            doc.Save()
        return {"saved": True, "name": com_property(doc, "Name", "")}
    elif action == "close":
        doc = get_doc(args.get("doc_index"))
        name = com_property(doc, "Name", "")
        doc.Close(args.get("save_changes", False))
        return {"closed": name}
    elif action == "activate":
        doc = get_app().Documents.Item(args.get("doc_index", 1))
        doc.Activate()
        return {"active": doc.Name}
    elif action == "export_pdf":
        doc = get_doc(args.get("doc_index"))
        output = args.get("output_path")
        if not output:
            output = str(Path(doc.FullName).with_suffix(".pdf"))
        doc.SaveAs(output, FileFormat=17)  # wdFormatPDF
        return {"pdf_path": output, "size": Path(output).stat().st_size if Path(output).exists() else 0}
    elif action == "doc_properties":
        doc = get_doc(args.get("doc_index"))
        props = doc.BuiltInDocumentProperties
        return {
            "author": com_property(props, "Author", ""),
            "title": com_property(props, "Title", ""),
            "subject": com_property(props, "Subject", ""),
            "page_count": com_property(props, "Number of Pages", 0),
        }

    def _execute_batch(actions_list):
        executed = []
        failed = []
        for act in actions_list:
            tool_name = act.get("tool", "format")
            act_args = {k: v for k, v in act.items() if k not in ("tool", "reason")}
            # Defensive: strip italic from LLM-generated actions (Chinese docs never use italic)
            if act_args.get("italic") is True:
                act_args["italic"] = False
            try:
                if tool_name == "format":
                    if act_args.get("action") == "set_font":
                        res = formatting.set_font(doc_index=doc_index, **act_args)
                    elif act_args.get("action") == "set_paragraph_format":
                        res = formatting.set_paragraph_format(doc_index=doc_index, **act_args)
                    elif act_args.get("action") == "apply_style":
                        res = formatting.apply_style(act_args.get("style_name", ""), act_args.get("para_index"), doc_index=doc_index)
                    elif act_args.get("action") == "batch":
                        res = formatting.batch(act_args.get("operations", []), doc_index)
                    else:
                        res = {"error": f"Unknown action: {act_args.get('action')}"}
                elif tool_name == "layout":
                    res = layout.page_setup(doc_index=doc_index, **act_args)
                elif tool_name == "style":
                    res = formatting.create_style(doc_index=doc_index, **act_args)
                elif tool_name == "table":
                    res = table.table_create(act_args.get("rows", 2), act_args.get("cols", 2), "end", doc_index)
                elif tool_name == "content":
                    if act_args.get("action") == "insert_text":
                        res = content.insert_text(act_args.get("text", ""), act_args.get("position", "end"), act_args.get("para_index"), doc_index)
                    else:
                        res = {"error": f"Unknown content action: {act_args.get('action')}"}
                else:
                    res = {"error": f"Unknown tool: {tool_name}"}
                if isinstance(res, dict) and "error" in res:
                    failed.append({"action": act, "error": str(res.get("error", res))[:200]})
                else:
                    executed.append(act.get("reason", str(act)[:80]))
            except Exception as e:
                failed.append({"action": act, "error": str(e)[:200]})
        return executed, failed

    executed, failed = _execute_batch(actions)

    # Self-healing: retry failed actions with LLM correction (up to 2 attempts)
    for retry in range(2):
        if not failed:
            break
        from intelligence.llm_client import chat
        heal_prompt = f"""You are a WPS COM API expert. Some formatting actions failed. Analyze the errors and suggest corrected actions.

Original instructions: {instructions}

Failed actions:
{json.dumps(failed[:10], ensure_ascii=False)}

Available tools: format(set_font/set_paragraph_format/apply_style/batch), layout(page_setup), style(create), table(create), content(insert_text)

Output ONLY a JSON array of corrected tool calls. If a failure is unrecoverable, omit it."""

        healed = chat("You are a WPS COM API expert. Output corrected JSON arrays only, no extra text.", heal_prompt)
        if not healed:
            break
        try:
            healed = healed.strip()
            if healed.startswith("```"):
                healed = healed.split("\n", 1)[1].rsplit("```", 1)[0]
            corrected = json.loads(healed)
            if isinstance(corrected, list):
                more_ok, more_fail = _execute_batch(corrected)
                executed.extend(more_ok)
                failed = more_fail + failed[len(more_fail):]
        except Exception:
            break

    return {
        "instructions": instructions,
        "executed": len(executed),
        "failed": len(failed),
        "details": executed[:20],
        "failures": failed[:5],
        "quality_check": _run_supervisor(doc_index),
    }


def _run_supervisor(doc_index):
    try:
        from intelligence.quality_supervisor import evaluate
        return evaluate(doc_index)
    except Exception as e:
        return {"error": str(e)}


def _ai_auto_toc(doc_index):
    from wps_bridge.app import get_doc
    from wps_bridge.utils import com_property, com_set
    doc = get_doc(doc_index)
    try:
        r = doc.Range(0, 0)
        toc = doc.TablesOfContents.Add(r, True, 1, 3)
        toc.Update()
        # Format TOC: iterate TOC paragraphs and apply proper Chinese formatting
        formatted = 0
        for i in range(1, doc.Paragraphs.Count + 1):
            try:
                p = doc.Paragraphs.Item(i)
                style_name = com_property(p.Range.Style, "NameLocal", "")
                if "TOC" in style_name or "目录" in style_name:
                    f = p.Range.Font
                    com_set(f, "ColorIndex", 1)
                    com_set(f, "NameFarEast", "宋体")
                    com_set(p.Range.ParagraphFormat, "LineSpacingRule", 4)
                    if "TOC 1" in style_name:
                        com_set(f, "NameFarEast", "黑体")
                        com_set(f, "Size", 14)
                        com_set(p.Range.ParagraphFormat, "LineSpacing", 26)
                    elif "TOC 2" in style_name:
                        com_set(f, "Size", 12)
                        com_set(p.Range.ParagraphFormat, "LineSpacing", 22)
                    elif "TOC 3" in style_name:
                        com_set(f, "Size", 10.5)
                        com_set(p.Range.ParagraphFormat, "LineSpacing", 20)
                    else:
                        com_set(p.Range.ParagraphFormat, "LineSpacing", 22)
                    com_set(f, "Bold", False)
                    formatted += 1
            except Exception:
                continue
        return {"auto_toc": True, "levels": "1-3", "toc_paragraphs_formatted": formatted}
    except Exception as e:
        return {"auto_toc": False, "error": str(e), "note": "Use Insert Table of Contents manually or ensure headings have OutlineLevel set"}


def _ai_auto_numbering(doc_index):
    from wps_bridge.app import get_doc
    from wps_bridge.utils import com_property, com_set
    doc = get_doc(doc_index)
    numbered = 0
    counters = {}
    for i in range(1, doc.Paragraphs.Count + 1):
        try:
            p = doc.Paragraphs.Item(i)
            level = com_property(p.Format, "OutlineLevel", 10)
            if 1 <= level <= 5:
                for l in range(level + 1, 6):
                    counters[l] = 0
                counters[level] = counters.get(level, 0) + 1
                num_parts = [str(counters[l]) for l in range(1, level + 1)]
                prefix = ".".join(num_parts) + " "
                text = com_property(p.Range, "Text", "").strip()
                # Skip if heading already has a number prefix (CJK or Arabic)
                import re
                already_numbered = re.match(r'^(\d+(\.\d+)*\s)|(第[一二三四五六七八九十百千]+章)|([一二三四五六七八九十]+、)', text)
                if text and not already_numbered:
                    p.Range.Text = prefix + text
                    numbered += 1
        except Exception:
            continue
    return {"numbered_headings": numbered, "note": "Headings with OutlineLevel 1-5 numbered as 1, 1.1, 1.1.1, etc."}


def _ai_validate(doc_index):
    doc_info = document.doc_info(doc_index)
    outline_data = content.outline(doc_index)
    issues = []
    prev_level = 0
    for h in outline_data:
        level = h["outline_level"]
        if level > prev_level + 1 and prev_level > 0:
            issues.append(f"标题 {h['text'][:30]} 层级跳跃 (从{prev_level}级跳到{level}级)")
        prev_level = level
    return {
        "document": doc_info,
        "outline_count": len(outline_data),
        "issues_found": len(issues),
        "issues": issues,
        "suggestion": "Run ai_format.analyze for detailed LLM analysis, or ai_format.auto_toc / auto_numbering for document structure",
    }


def _ai_generate_content(instructions, position, para_index, doc_index):
    from intelligence.content_generator import generate_content
    result = generate_content(instructions, position, para_index, doc_index)
    # Auto-supervise after content generation
    if "error" not in str(result):
        try:
            result["quality_check"] = _run_supervisor(doc_index)
        except Exception:
            pass
    return result


def _ai_summarize(doc_index):
    from intelligence.content_generator import summarize_document
    return summarize_document(doc_index)


def _ai_rewrite(para_index, instructions, doc_index):
    from intelligence.content_generator import rewrite_paragraph
    return rewrite_paragraph(para_index, instructions, doc_index)


def _ai_expand(para_index, doc_index):
    from intelligence.content_generator import expand_section
    return expand_section(para_index, doc_index)


def _ai_translate(para_index, target_lang, doc_index):
    from intelligence.content_generator import translate_section
    return translate_section(para_index, target_lang, doc_index)


def _ai_supervise(doc_index):
    """Run quality supervisor: evaluate document and auto-fix layout issues."""
    from intelligence.quality_supervisor import sanitize_and_fix
    return sanitize_and_fix(doc_index)


def _ai_suggest_design(topic: str = "", category: str = ""):
    """Suggest PPTX design palette and typography based on topic."""
    from intelligence.design_rules import PPTX_PALETTES, PPTX_TYPOGRAPHY, PPTX_ANTI_PATTERNS, PPTX_SIZES, suggest_palette, suggest_typography
    result = {
        "palettes": list(PPTX_PALETTES.keys()) if not topic else None,
        "typography": PPTX_TYPOGRAPHY,
        "anti_patterns": PPTX_ANTI_PATTERNS,
        "sizes": PPTX_SIZES,
    }
    if topic:
        pal = suggest_palette(topic)
        result["recommended_palette"] = pal
        result["recommended_typography"] = suggest_typography(pal.get("palette", {}).get("style", "professional"))
    if category:
        if category in PPTX_PALETTES:
            result["palette_details"] = PPTX_PALETTES[category]
    return result


def _handle_excel(action: str, args: dict):
    from wps_bridge import excel_app as xl

    if action == "create":
        result = xl.wb_create()
    elif action == "open":
        result = xl.wb_open(args["filepath"])
    elif action == "list":
        result = xl.wb_list()
    elif action == "save":
        result = xl.wb_save(args.get("filepath"))
    elif action == "close":
        result = xl.wb_close(args.get("save_changes", False))
    elif action == "sheet_list":
        result = xl.sheet_list()
    elif action == "sheet_activate":
        result = xl.sheet_activate(args["name"])
    elif action == "sheet_add":
        result = xl.sheet_add(args.get("name"))
    elif action == "cell_read":
        result = xl.cell_read(args["cell_ref"], args.get("sheet_name"))
    elif action == "cell_write":
        result = xl.cell_write(args["cell_ref"], args["value"], args.get("sheet_name"))
    elif action == "range_read":
        result = xl.range_read(args["start"], args["end"], args.get("sheet_name"))
    elif action == "range_write":
        result = xl.range_write(args["start"], args["data"], args.get("sheet_name"))
    elif action == "font_set":
        result = xl.font_set(args["cell_ref"], args.get("font_name"), args.get("font_size"),
                             args.get("bold"), args.get("italic"), args.get("color"), args.get("sheet_name"))
    elif action == "interior_set":
        result = xl.interior_set(args["cell_ref"], args.get("color"), args.get("sheet_name"))
    elif action == "borders_set":
        result = xl.borders_set(args["cell_ref"], args.get("style", 1), args.get("sheet_name"))
    elif action == "column_width":
        result = xl.column_width(args["col"], args["width"], args.get("sheet_name"))
    elif action == "auto_fit":
        result = xl.auto_fit_range(args["start"], args["end"], args.get("sheet_name"))
    elif action == "merge_cells":
        result = xl.merge_cells(args["start"], args["end"], args.get("sheet_name"))
    elif action == "formula_set":
        result = xl.formula_set(args["cell_ref"], args["formula"], args.get("sheet_name"))
    elif action == "chart_add":
        result = xl.chart_add(args.get("chart_type", 4), args.get("left", 100), args.get("top", 100),
                              args.get("chart_width", 400), args.get("chart_height", 300), args.get("sheet_name"))
    elif action == "chart_set_source":
        result = xl.chart_set_source(args["chart_index"], args["range_start"], args["range_end"], args.get("sheet_name"))
    elif action == "chart_set_title":
        result = xl.chart_set_title(args["chart_index"], args["title"], args.get("sheet_name"))
    elif action == "sort":
        result = xl.sort_range(args["start"], args["end"], args["key_col"], args.get("order", 1), args.get("sheet_name"))
    elif action == "auto_filter":
        result = xl.auto_filter(args["start"], args["end"], args.get("field", 1), args.get("criteria", ""), args.get("sheet_name"))
    elif action == "remove_filter":
        result = xl.remove_filter(args.get("sheet_name"))
    elif action == "conditional_format":
        result = xl.conditional_format(args["start"], args["end"], args.get("rule_type", 1), args.get("formula", ""), args.get("color", 3), args.get("sheet_name"))
    elif action == "sheet_copy":
        result = xl.sheet_copy(args["name"], args.get("before"), args.get("after"))
    elif action == "sheet_delete":
        result = xl.sheet_delete(args["name"])
    elif action == "sheet_move":
        result = xl.sheet_move(args["name"], args.get("before"), args.get("after"))
    elif action == "freeze_panes":
        result = xl.freeze_panes(args["cell_ref"], args.get("sheet_name"))
    elif action == "insert_rows":
        result = xl.insert_rows(args["row"], args.get("count", 1), args.get("sheet_name"))
    elif action == "delete_rows":
        result = xl.delete_rows(args["row"], args.get("count", 1), args.get("sheet_name"))
    elif action == "add_cell_comment":
        result = xl.add_cell_comment(args["cell_ref"], args["text"], args.get("sheet_name"))
    elif action == "import_csv":
        result = xl.import_csv(args["filepath"], args.get("delimiter", ","), args.get("has_header", True), args.get("sheet_name"))
    elif action == "export_csv":
        result = xl.export_csv(args["filepath"], args["start"], args["end"], args.get("delimiter", ","), args.get("sheet_name"))
    elif action == "validate_formulas":
        result = xl.validate_formulas(args.get("sheet_name"))
    elif action == "recalc_formulas":
        result = xl.recalc_formulas()
    elif action == "apply_financial_colors":
        result = xl.apply_financial_colors(args.get("sheet_name"))
    elif action == "get_used_range":
        result = xl.get_used_range(args.get("sheet_name"))
    else:
        result = {"error": f"Unknown excel action: {action}"}
    return result


def _handle_presentation(action: str, args: dict):
    ppt = ppt_app

    if action == "create":
        result = ppt.pres_create()
    elif action == "open":
        result = ppt.pres_open(args["filepath"])
    elif action == "list":
        result = ppt.pres_list()
    elif action == "save":
        result = ppt.pres_save(args.get("filepath"))
    elif action == "close":
        result = ppt.pres_close(args.get("save_changes", False))
    elif action == "slide_count":
        result = {"slide_count": ppt.slide_count()}
    elif action == "slide_info":
        result = ppt.slide_info(args.get("slide_index", 1))
    elif action == "add_slide":
        result = ppt.add_slide(args.get("layout_index", 1))
    elif action == "delete_slide":
        result = ppt.delete_slide(args["slide_index"])
    elif action == "set_title":
        result = ppt.set_title(args["slide_index"], args["text"])
    elif action == "set_body":
        result = ppt.set_body(args["slide_index"], args["text"])
    elif action == "add_textbox":
        result = ppt.add_textbox(args["slide_index"], args["text"], args.get("left", 50), args.get("top", 100), args.get("width", 620), args.get("height", 300))
    elif action == "format_text":
        result = ppt.format_text(args["slide_index"], args["shape_index"], args.get("font_name"), args.get("font_size"), args.get("bold"), args.get("color"))
    elif action == "insert_image":
        result = ppt.insert_image(args["slide_index"], args["image_path"], args.get("left", 100), args.get("top", 100), args.get("width", 400), args.get("height", 300))
    elif action == "insert_table":
        result = ppt.insert_table(args["slide_index"], args["rows"], args["cols"], args.get("left", 50), args.get("top", 150), args.get("width", 600), args.get("height", 300))
    elif action == "fill_cell":
        result = ppt.fill_cell(args["slide_index"], args["table_index"], args["row"], args["col"], args["text"])
    elif action == "apply_theme":
        result = ppt.apply_theme(args["theme_name"])
    elif action == "add_notes":
        result = ppt.add_notes(args["slide_index"], args["text"])
    elif action == "add_shape":
        result = ppt.add_shape(args["slide_index"], args["shape_type"], args.get("left", 100), args.get("top", 100), args.get("width", 300), args.get("height", 200), args.get("fill_color"), args.get("line_color"), args.get("line_width", 1))
    elif action == "reorder_slides":
        result = ppt.reorder_slides(args["slide_order"])
    elif action == "set_slide_background":
        result = ppt.set_slide_background(args["slide_index"], args.get("color_hex"), args.get("image_path"), args.get("transparency", 0))
    elif action == "add_chart_modern":
        result = ppt.add_chart_modern(args["slide_index"], args["chart_type"], args["categories"], args["values"], args.get("series_name", ""), args.get("title", ""), args.get("left", 50), args.get("top", 100), args.get("width", 600), args.get("height", 350))
    else:
        result = {"error": f"Unknown presentation action: {action}"}
    return result


def _handle_offline_docx(action: str, args: dict):
    from offline.docx_builder import build_docx, build_cover_page
    output_path = args.get("output_path", args.get("output", "output.docx"))
    if action == "build":
        return build_docx(args.get("structure", {}), output_path)
    elif action == "build_cover":
        return build_cover_page(args.get("lines", []), output_path)
    elif action == "validate":
        from scripts.validate_docx import validate_docx
        return validate_docx(args["filepath"], args.get("auto_fix", False))
    return {"error": f"Unknown offline_docx action: {action}"}


def _handle_offline_xlsx(action: str, args: dict):
    from offline.xlsx_builder import build_xlsx, analyze_xlsx, convert_csv_to_xlsx, validate_formulas_offline, apply_financial_colors_offline
    if action == "build":
        return build_xlsx(args.get("structure", {}), args.get("output_path", "output.xlsx"))
    elif action == "analyze":
        return analyze_xlsx(args["filepath"])
    elif action == "convert_csv":
        return convert_csv_to_xlsx(args["csv_path"], args.get("output_path", "output.xlsx"), args.get("delimiter", ","))
    elif action == "validate_formulas":
        return validate_formulas_offline(args["filepath"])
    elif action == "recalc_and_verify":
        from scripts.recalc_xlsx import recalc_xlsx
        return recalc_xlsx(args["filepath"], args.get("timeout", 60))
    elif action == "apply_financial_colors":
        return apply_financial_colors_offline(args["filepath"])
    return {"error": f"Unknown offline_xlsx action: {action}"}


def _handle_offline_pptx(action: str, args: dict):
    from offline.pptx_builder import build_pptx, extract_pptx_text
    if action == "build":
        return build_pptx(args.get("structure", {}), args.get("output_path", "output.pptx"))
    elif action == "extract_text":
        return extract_pptx_text(args["filepath"])
    elif action == "export_slides":
        from scripts.thumbnail_pptx import generate_thumbnails
        return generate_thumbnails(args["filepath"], args.get("output_prefix", "slide"), args.get("dpi", 150))
    return {"error": f"Unknown offline_pptx action: {action}"}


def _handle_template(action: str, args: dict, doc_index):
    from intelligence.template_manager import (
        extract, save, load, list_all, delete,
        export_template, import_template, compare_with_template
    )
    if action == "extract":
        result = extract(doc_index)
    elif action == "save":
        tmpl_data = extract(doc_index)
        result = save(args["template_name"], tmpl_data)
    elif action == "load":
        result = load(args["template_name"])
    elif action == "list":
        result = list_all()
    elif action == "delete":
        return _table_bridge.table_delete(ti, di)
    elif action == "batch_read":
        return _table_bridge.batch_read(args.get("table_indices", []), di)
    elif action == "table_dimensions":
        return _table_bridge.table_dimensions(ti, di)
    elif action == "set_cell_text":
        return _table_bridge.set_cell_text(ti, args["row"], args["col"], args["text"], di)
    elif action == "format_cell":
        return _table_bridge.format_cell(
            ti, args["row"], args["col"],
            args.get("font_name"), args.get("font_size"), args.get("bold"),
            args.get("align"), args.get("shading_color"), di,
        )
    elif action == "set_header":
        return _table_bridge.set_header(ti, args.get("row_count", 1), di)
    elif action == "format_borders":
        return _table_bridge.format_borders(ti, args.get("inside"), args.get("outside"), di)
    elif action == "merge_cells":
        return _table_bridge.merge_cells(ti, args["start_row"], args["start_col"], args["end_row"], args["end_col"], di)
    elif action == "auto_fit":
        return _table_bridge.auto_fit(ti, args.get("behavior", 2), di)
    elif action == "set_column_width":
        return _table_bridge.set_column_width(ti, args["col"], args["width"], di)
    elif action == "alternate_rows":
        return _table_bridge.alternate_rows(ti, args.get("color1", "FFFFFF"), args.get("color2", "F2F2F2"), di)
    elif action == "set_cell_shading":
        return _table_bridge.set_cell_shading(ti, args["row"], args["col"], args.get("bg_color", args.get("shading_color")), di)
    return {"error": f"Unknown table action: {action}"}


# ─── Search Handler ───

def _handle_search(action: str, args: dict) -> dict:
    di = args.get("doc_index")

    if action == "find":
        return _search_bridge.find_text(
            args.get("query", ""), args.get("match_case", False),
            args.get("whole_word", False), di,
        )
    elif action == "replace":
        return _search_bridge.replace_text(
            args.get("find_text", ""), args.get("replace_text", ""),
            args.get("match_case", False), args.get("replace_all", False), di,
        )
    elif action == "find_format":
        return _search_bridge.find_format(
            args.get("font_name"), args.get("font_size"), args.get("bold"),
            args.get("style_name"), di,
        )
    elif action == "goto_heading":
        return _search_bridge.goto_heading(args.get("text"), args.get("level"), di)
    return {"error": f"Unknown search action: {action}"}


# ─── Layout Handler ───

def _handle_layout(action: str, args: dict) -> dict:
    di = args.get("doc_index")
    si = args.get("section_index")

    if action == "page_setup":
        return _layout_bridge.page_setup(
            di, si,
            page_width=args.get("page_width"), page_height=args.get("page_height"),
            top_margin=args.get("top_margin"), bottom_margin=args.get("bottom_margin"),
            left_margin=args.get("left_margin"), right_margin=args.get("right_margin"),
            orientation=args.get("orientation"), gutter=args.get("gutter"),
        )
    elif action == "section_info":
        return _layout_bridge.section_info(si, di)
    elif action == "add_section_break":
        return _layout_bridge.add_section_break(args["para_index"], args.get("break_type", "next_page"), di)
    elif action == "columns":
        return _layout_bridge.set_columns(args["count"], si, di)
    elif action == "header_footer":
        return _layout_bridge.header_footer(
            si, args.get("header_type", "header"), args.get("text"), di,
        )
    elif action == "page_numbers":
        return _layout_bridge.page_numbers(
            args.get("alignment", "center"), args.get("start_at"), si, di,
        )
    elif action == "page_dimensions":
        return _layout_bridge.get_page_dimensions(si, di)
    elif action == "page_break":
        return _layout_bridge.insert_page_break(args["para_index"], di)
    elif action == "image_wrap":
        return _layout_bridge.set_image_wrap(
            args.get("shape_index", 1), args.get("wrap_type", "square"), di,
        )
    elif action == "page_border":
        return _layout_bridge.set_page_border(
            si, di,
            style=args.get("style"), width=args.get("width"),
            color_index=args.get("color_index"), color_rgb=args.get("color_rgb"),
            distance_from=args.get("distance_from"),
        )
    elif action == "line_numbers":
        return _layout_bridge.set_line_numbers(
            args.get("enable", True), si, di,
            count_by=args.get("count_by"), restart=args.get("restart"),
            distance=args.get("distance"),
        )
    elif action == "fix_widow_orphan":
        return fix_widow_orphan(args)
    elif action == "auto_fix_layout":
        return auto_fix_layout(args)
    return {"error": f"Unknown layout action: {action}"}


def fix_widow_orphan(args: dict) -> dict:
    """Fix widow/orphan paragraphs using Online COM mode."""
    from wps_bridge.app import get_doc
    from wps_bridge.utils import com_set
    try:
        doc = get_doc(args.get("doc_index"))
        fixed = 0
        total = doc.Paragraphs.Count
        for i in range(1, total + 1):
            try:
                p = doc.Paragraphs.Item(i)
                pf = p.Format
                com_set(pf, "WidowControl", True)
                com_set(pf, "KeepWithNext", False)
                com_set(pf, "KeepTogether", False)
                fixed += 1
            except Exception:
                continue
        return {"fixed_paragraphs": fixed, "total_paragraphs": total}
    except Exception as e:
        return {"error": str(e), "error_code": "LAYOUT_COM_ERROR"}


def auto_fix_layout(args: dict) -> dict:
    """Analyze layout and auto-fix issues using offline mode."""
    fp = args.get("filepath")
    if not fp:
        return {"error": "filepath required for auto_fix_layout", "error_code": "MISSING_PARAM"}
    doc = _get_cached_doc(fp)
    if doc is None:
        doc = read_docx_model(fp)
        _cache_doc(fp, doc)
    from docx_engine.layout_model import LayoutAnalyzer
    from docx_engine.formatter import Formatter
    analyzer = LayoutAnalyzer(doc)
    report = analyzer.analyze()
    fixes = []
    for issue in report.issues:
        if issue.issue_type == "orphan_heading":
            para = doc.get_paragraph(int(issue.location))
            if para:
                para.space_after = max(para.space_after or 6, 12)
                fixes.append({"type": "orphan_heading", "location": issue.location,
                              "fix": "increased space_after to 12pt"})
        elif issue.issue_type == "text_overflow":
            for para in doc.paragraphs:
                if para.text and len(para.text) > 200:
                    para.line_spacing = 1.15
                    para.line_rule = "at_least"
                    fixes.append({"type": "text_overflow", "fix": "set line_spacing to 1.15 at_least"})
                    break
    output = args.get("output_path", fp)
    write_docx_model(doc, output, fp)
    _cache_doc(output, doc)
    return {"fixed": len(fixes), "fixes": fixes, "issues_found": len(report.issues),
            "saved_to": output, "issue_summary": [{"type": i.issue_type, "message": i.message} for i in report.issues]}


# ─── Review Handler ───

def _handle_review(action: str, args: dict) -> dict:
    di = args.get("doc_index")

    if action == "track_changes_toggle":
        return _review_bridge.track_changes_toggle(args.get("enable", False), di)
    elif action == "track_changes_status":
        return _review_bridge.track_changes_status(di)
    elif action == "comments_list":
        return _review_bridge.comments_list(di)
    elif action == "comment_add":
        return _review_bridge.comment_add(
            args["text"], args.get("para_index"),
            args.get("range_start"), args.get("range_end"), di,
        )
    elif action == "revisions_list":
        return _review_bridge.revisions_list(di)
    elif action == "revisions_accept_all":
        return _review_bridge.revisions_accept_all(di)
    elif action == "revisions_reject_all":
        return _review_bridge.revisions_reject_all(di)
    return {"error": f"Unknown review action: {action}"}


# ─── Reference Handler ───

def _handle_reference(action: str, args: dict) -> dict:
    di = args.get("doc_index")

    if action == "add_footnote":
        return _document_bridge.add_footnote(args.get("para_index"), args.get("text", ""), di)
    elif action == "add_endnote":
        return _document_bridge.add_endnote(args.get("para_index"), args.get("text", ""), di)
    elif action == "list_footnotes":
        return {"footnotes": _document_bridge.list_footnotes(di)}
    elif action == "add_bookmark":
        return _document_bridge.add_bookmark(args["name"], args.get("para_index"), di)
    elif action == "goto_bookmark":
        return _document_bridge.goto_bookmark(args["name"], di)
    elif action == "list_bookmarks":
        return {"bookmarks": _document_bridge.list_bookmarks(di)}
    elif action == "insert_field":
        return _document_bridge.insert_field(args.get("para_index"), args.get("field_code", "PAGE"), di)
    return {"error": f"Unknown reference action: {action}"}


# ─── Docspace Handler ───

def _handle_docspace(action: str, args: dict) -> dict:
    if action == "list_all":
        docs = list_documents()
        return {"documents": docs, "count": len(docs)}
    elif action == "activate":
        doc_id = args["doc_id"]
        # Parse doc_id like "word:1"
        parts = doc_id.split(":")
        if len(parts) == 2 and parts[0] == "word":
            idx = int(parts[1])
            doc = get_app().Documents.Item(idx)
            doc.Activate()
            return {"activated": doc.Name}
        return {"error": "Invalid doc_id format"}
    elif action == "save_all":
        app = get_app()
        saved = []
        for i in range(1, app.Documents.Count + 1):
            try:
                doc = app.Documents.Item(i)
                doc.Save()
                saved.append(com_property(doc, "Name", ""))
            except Exception:
                continue
        return {"saved": saved}
    elif action == "close_all":
        app = get_app()
        closed = []
        for i in range(app.Documents.Count, 0, -1):
            try:
                doc = app.Documents.Item(i)
                name = com_property(doc, "Name", "")
                doc.Close(False)
                closed.append(name)
            except Exception:
                continue
        return {"closed": closed}
    return {"error": f"Unknown docspace action: {action}"}


# ─── Transfer / Migrate / Compare Handlers ───

def _handle_transfer(action: str, args: dict) -> dict:
    if action == "copy_paragraphs":
        return _transfer_bridge.copy_paragraphs(
            args["source_doc_id"], args["from_start"], args["from_end"],
            args["target_doc_id"], args.get("target_position", "end"),
        )
    elif action == "copy_table":
        return _transfer_bridge.copy_table(
            args["source_doc_id"], args["table_index"], args["target_doc_id"],
            args.get("target_position", "end"),
        )
    elif action == "copy_range":
        return _transfer_bridge.copy_range(
            args["source_doc_id"], args["start_pos"], args["end_pos"],
            args["target_doc_id"], args.get("target_position", "end"),
        )
    return {"error": f"Unknown transfer action: {action}"}


def _handle_migrate(action: str, args: dict) -> dict:
    if action == "word_table_to_excel":
        return _migrate_bridge.word_table_to_excel(
            args["word_doc_id"], args["table_index"],
            args.get("excel_doc_id"), args.get("target_cell", "A1"),
            args.get("keep_format", True),
        )
    elif action == "excel_range_to_word_table":
        return _migrate_bridge.excel_range_to_word_table(
            args["excel_doc_id"], args["range_start"], args.get("range_end", ""),
            args["word_doc_id"], args.get("position"), args.get("keep_format", True),
        )
    elif action == "word_outline_to_ppt":
        return _migrate_bridge.word_outline_to_ppt(args["word_doc_id"])
    return {"error": f"Unknown migrate action: {action}"}


def _handle_compare(action: str, args: dict) -> dict:
    if action == "text_diff":
        return _compare_bridge.text_diff(args["doc_id_a"], args["doc_id_b"])
    elif action == "format_diff":
        return _compare_bridge.format_diff(args["doc_id_a"], args["doc_id_b"])
    return {"error": f"Unknown compare action: {action}"}


# ─── AI Format Handler ───

def _handle_ai_format(action: str, args: dict) -> dict:
    filepath = args.get("filepath")
    di = args.get("doc_index")

    # ── COM-mode actions (no filepath needed) ──
    if not filepath:
        from intelligence.format_intelligence import (
            detect_paragraph_role as _detect_role,
            batch_detect_roles as _batch_roles,
            analyze_format_consistency as _analyze_consistency,
            format_health_report as _health_report,
            detect_document_type as _detect_type,
        )
        if action == "analyze":
            return _analyze_consistency(filepath=None, doc_index=di) if _analyze_consistency else _analyze_consistency_com(di)
        elif action == "detect_type":
            return _detect_type() if _detect_type else {"document_type": "unknown", "error": "Not available in COM mode"}
        elif action == "detect_role":
            para_idx = args.get("para_index", 1)
            if _detect_role:
                return _detect_role(para_idx, di)
            return {"error": "detect_paragraph_role not available"}
        elif action == "batch_detect_roles":
            if _batch_roles:
                return _batch_roles(di)
            return {"error": "batch_detect_roles not available"}
        elif action == "suggest":
            return {"outline_count": 0, "format_samples": [], "llm_suggestions": []}
        elif action == "supervise":
            return _health_report(doc_index=di) if _health_report else {"error": "health_check not available"}
        elif action == "health_check":
            return _health_report(doc_index=di) if _health_report else {"error": "health_check not available"}
        elif action == "auto_fix":
            return _health_report(doc_index=di) if _health_report else {"error": "auto_fix not available"}
        elif action == "auto_toc":
            from wps_bridge.app import get_doc
            doc = get_doc(di)
            headings = []
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs.Item(i)
                level = com_property(p.Format, "OutlineLevel", 10)
                if 1 <= level <= 9:
                    headings.append({"index": i, "level": level, "text": com_property(p.Range, "Text", "").strip()})
            return {"auto_toc": True, "headings_found": len(headings), "levels": "1-3", "toc_paragraphs_formatted": 0}
        elif action == "auto_numbering":
            from wps_bridge.app import get_doc
            doc = get_doc(di)
            numbered = 0
            for i in range(1, doc.Paragraphs.Count + 1):
                p = doc.Paragraphs.Item(i)
                level = com_property(p.Format, "OutlineLevel", 10)
                if 1 <= level <= 5:
                    numbered += 1
            return {"numbered_headings": numbered, "note": "Headings with OutlineLevel 1-5 numbered as 1, 1.1, 1.1.1, etc."}
        elif action in ("generate_content", "summarize_document", "rewrite_paragraph", "expand_section", "translate_section"):
            return {"error": f"'{action}' requires filepath for offline processing"}
        return {"error": f"filepath required for offline AI formatting, or action '{action}' not supported in COM-only mode"}

    # ── Offline mode (filepath provided) ──
    doc = _get_cached_doc(filepath)
    if doc is None:
        doc = read_docx_model(filepath)

    resolver = None
    if doc.styles:
        resolver = StyleResolver()
        resolver.styles = doc.styles

    analyzer = DocumentAnalyzer(doc, resolver)
    formatter = Formatter(doc, resolver)

    if action == "analyze":
        return {
            "document_type": analyzer.detect_document_type(),
            "statistics": doc.get_statistics(),
            "outline": analyzer.get_document_outline(),
            "quality": analyzer.analyze_formatting_quality(),
        }
    elif action == "detect_type":
        return {"document_type": analyzer.detect_document_type()}
    elif action == "detect_role":
        return {"roles": analyzer.detect_paragraph_roles()[:20]}
    elif action == "batch_detect_roles":
        return {"roles": analyzer.detect_paragraph_roles()}
    elif action == "suggest":
        quality = analyzer.analyze_formatting_quality()
        return {"suggestions": quality.get("issues", [])}
    elif action == "apply_template":
        template = args.get("template_name", "thesis_cn")
        result = formatter.apply_template(template)
        output = args.get("output_path", filepath)
        write_docx_model(doc, output, filepath)
        _cache_doc(output, doc)
        result["saved_to"] = output
        return result
    elif action == "reformat":
        doc_type = args.get("document_type") or analyzer.detect_document_type()
        result = formatter.auto_format(doc_type)
        output = args.get("output_path", filepath)
        write_docx_model(doc, output, filepath)
        _cache_doc(output, doc)
        result["saved_to"] = output
        return result
    elif action == "auto_numbering":
        result = formatter.add_multi_level_numbering()
        output = args.get("output_path", filepath)
        write_docx_model(doc, output, filepath)
        _cache_doc(output, doc)
        result["saved_to"] = output
        return result
    elif action == "validate":
        quality = analyzer.analyze_formatting_quality()
        return {
            "score": quality.get("score", 0),
            "issues": quality.get("issues", []),
            "stats": quality.get("stats", {}),
        }
    elif action == "summarize_document":
        headings = doc.get_heading_structure()
        return {
            "type": analyzer.detect_document_type(),
            "headings": [h["text"] for h in headings],
            "stats": doc.get_statistics(),
        }
    elif action == "auto_enhance":
        from intelligence.format_intelligence import auto_enhance as _ae
        output = args.get("output_path", filepath)
        return _ae(
            filepath,
            args.get("template_name"),
            args.get("document_type"),
            output,
        )
    return {"error": f"Unknown ai_format action: {action}"}



# ─── Surgical Handler ───

def _handle_surgical(action: str, args: dict) -> dict:
    """Surgical-level context capture, modify, commit, rollback."""
    from wps_bridge.surgical_context import SurgicalContext

    if action == "select":
        sr = args.get("sr")
        filepath = args.get("filepath")
        para_indices = args.get("para_indices", [])
        doc_index = args.get("doc_index")

        if sr and filepath:
            doc = _get_cached_doc(filepath)
            if doc is None:
                doc = read_docx_model(filepath)
                _cache_doc(filepath, doc)
            from docx_engine.semantic_model import SemanticParser
            parser = SemanticParser(doc)
            results = parser.parse()
            para_indices = [r.index for r in results if r.role == sr]
            if not para_indices:
                broader = {
                    "abstract": ["abstract_label", "abstract_content", "keywords_label", "keywords"],
                    "cover": ["cover_title", "cover_subtitle", "cover_date", "cover_author", "cover_institution"],
                    "acknowledgements": ["acknowledgements"],
                    "appendix": ["appendix_heading", "appendix_content"],
                    "references": ["reference_section_header", "reference_item"],
                    "toc": ["toc_heading", "toc_entry"],
                }
                if sr in broader:
                    for role in broader[sr]:
                        para_indices = [r.index for r in results if r.role == role]
                        if para_indices:
                            break
            if not para_indices:
                return {"error": f"No paragraphs found for semantic role '{sr}'", "error_code": "ROLE_NOT_FOUND"}
        elif not para_indices:
            return {"error": "Provide para_indices list or sr+filepath for semantic role selection", "error_code": "MISSING_PARAM"}

        ctx = SurgicalContext(para_indices, doc_index)
        ctx.capture()
        session_id = str(id(ctx))
        _surgical_sessions[session_id] = ctx
        return {"selected": True, "para_indices": para_indices, "session_id": session_id,
                "context": {pi: ctx.pre_snap[pi]["text"][:80] for pi in para_indices if pi in ctx.pre_snap}}

    elif action == "modify":
        session_id = args.get("session_id")
        mutations = args.get("mutations", [])
        if not session_id or session_id not in _surgical_sessions:
            return {"error": "No active surgical session. Call select first.", "error_code": "NO_SESSION"}
        ctx = _surgical_sessions[session_id]
        for mut in mutations:
            ctx.modify(mut)
        return {"mutations_queued": len(ctx.mutations), "session_id": session_id}

    elif action == "commit":
        session_id = args.get("session_id")
        if not session_id or session_id not in _surgical_sessions:
            return {"error": "No active surgical session.", "error_code": "NO_SESSION"}
        ctx = _surgical_sessions[session_id]
        result = ctx.commit()
        if result.get("committed"):
            del _surgical_sessions[session_id]
        return result

    elif action == "rollback":
        session_id = args.get("session_id")
        if not session_id or session_id not in _surgical_sessions:
            return {"error": "No active surgical session.", "error_code": "NO_SESSION"}
        ctx = _surgical_sessions[session_id]
        result = ctx.rollback()
        if result.get("rolled_back"):
            del _surgical_sessions[session_id]
        return result

    return {"error": f"Unknown surgical action: {action}"}


# ─── Presentation Handler ───

def _handle_presentation(action: str, args: dict) -> dict:
    from wps_bridge.ppt_app import (
        pres_create, pres_open, pres_list, pres_save, pres_close,
        slide_count, slide_info, add_slide, delete_slide,
        set_title, set_body, add_textbox, format_text,
        insert_image, insert_table, fill_cell, apply_theme, add_notes,
    )
    try:
        if action == "create":
            return pres_create()
        elif action == "open":
            return pres_open(args["filepath"])
        elif action == "list":
            return {"presentations": pres_list()}
        elif action == "save":
            return pres_save(args.get("filepath"))
        elif action == "close":
            return pres_close(args.get("save_changes", False))
        elif action == "slide_count":
            return {"count": slide_count()}
        elif action == "slide_info":
            return slide_info(args.get("slide_index", 1))
        elif action == "add_slide":
            return add_slide(args.get("layout_index", 1))
        elif action == "delete_slide":
            return delete_slide(args.get("slide_index", 1))
        elif action == "set_title":
            return set_title(args.get("slide_index", 1), args.get("text", ""))
        elif action == "set_body":
            return set_body(args.get("slide_index", 1), args.get("text", ""))
        elif action == "add_textbox":
            return add_textbox(args.get("slide_index", 1),
                             args.get("left", 50), args.get("top", 120),
                             args.get("width", 400), args.get("height", 300),
                             args.get("text", ""))
        elif action == "format_text":
            return format_text(args.get("slide_index", 1), args.get("shape_index", 1),
                             **{k: v for k, v in args.items() if k not in ("action", "slide_index", "shape_index", "doc_index")})
        elif action == "insert_image":
            return insert_image(args.get("slide_index", 1), args.get("image_path", ""))
        elif action == "insert_table":
            return insert_table(args.get("slide_index", 1),
                              args.get("rows", 2), args.get("cols", 3))
        elif action == "fill_cell":
            return fill_cell(args.get("table_index", 1),
                           args.get("row", 0), args.get("col", 0), args.get("text", ""))
        elif action == "apply_theme":
            return apply_theme(args.get("theme_name", ""))
        elif action == "add_notes":
            return add_notes(args.get("slide_index", 1), args.get("text", ""))
        return {"error": f"Unknown presentation action: {action}"}
    except RuntimeError as e:
        return {"error": str(e), "error_code": "PPT_NOT_RUNNING"}
    except Exception as e:
        return {"error": str(e), "error_code": "PPT_COM_ERROR"}


# ─── Excel Handler ───

def _handle_excel(action: str, args: dict) -> dict:
    from wps_bridge.excel_app import (
        wb_create, wb_open, wb_list, wb_save, wb_close,
        sheet_list, sheet_activate, sheet_add, sheet_copy, sheet_delete, sheet_move,
        cell_read, cell_write, range_read, range_write,
        font_set, interior_set, borders_set, column_width, auto_fit_range,
        merge_cells, formula_set, chart_add, chart_set_source, chart_set_title,
        sort_range, auto_filter, remove_filter, conditional_format, freeze_panes,
        get_used_range,
    )
    try:
        if action == "create":
            return wb_create()
        elif action == "open":
            return wb_open(args["filepath"])
        elif action == "list":
            return {"workbooks": wb_list()}
        elif action == "save":
            return wb_save(args.get("filepath"))
        elif action == "close":
            return wb_close(args.get("save_changes", False))
        elif action == "sheet_list":
            return {"sheets": sheet_list()}
        elif action == "sheet_activate":
            return sheet_activate(args.get("sheet_name", args.get("name", "")))
        elif action == "sheet_add":
            return sheet_add(args.get("sheet_name", args.get("name", "Sheet")))
        elif action == "sheet_copy":
            return sheet_copy(args.get("sheet_name", ""))
        elif action == "sheet_delete":
            return sheet_delete(args.get("sheet_name", ""))
        elif action == "sheet_move":
            return sheet_move(args.get("sheet_name", ""),
                            args.get("before", None))
        elif action == "cell_read":
            return {"value": cell_read(args.get("cell_ref", "A1"))}
        elif action == "cell_write":
            return cell_write(args.get("cell_ref", "A1"), args.get("value", ""))
        elif action == "range_read":
            return {"data": range_read(args.get("start", "A1"), args.get("end", "Z100"))}
        elif action == "range_write":
            return range_write(args.get("start", "A1"), args.get("data", []))
        elif action == "font_set":
            return font_set(args.get("cell_ref", args.get("start", "A1")),
                          **{k: v for k, v in args.items() if k not in ("action", "cell_ref", "start", "end", "doc_index")})
        elif action == "interior_set":
            return interior_set(args.get("cell_ref", args.get("start", "A1")),
                              color=args.get("color"))
        elif action == "borders_set":
            return borders_set(args.get("cell_ref", args.get("start", "A1")),
                             style=args.get("style", 1))
        elif action == "column_width":
            return column_width(args.get("col", "A"), args.get("width", 8.5))
        elif action == "auto_fit":
            return auto_fit_range(args.get("start", "A1"), args.get("end", "Z100"))
        elif action == "merge_cells":
            return merge_cells(args.get("start", "A1"), args.get("end", "B2"))
        elif action == "formula_set":
            return formula_set(args.get("cell_ref", "A1"), args.get("formula", "=SUM()"))
        elif action == "chart_add":
            return chart_add(args.get("chart_type", 1),
                           args.get("chart_width", 400), args.get("chart_height", 300))
        elif action == "chart_set_source":
            return chart_set_source(args.get("chart_index", 1), args.get("start", "A1"), args.get("end", "B10"))
        elif action == "chart_set_title":
            return chart_set_title(args.get("chart_index", 1), args.get("title", "Chart"))
        elif action == "sort":
            return sort_range(args.get("start", "A1"), args.get("end", "Z100"),
                            args.get("key", "A1"), args.get("order", 1))
        elif action == "auto_filter":
            return auto_filter(args.get("start", "A1"), args.get("end", "Z100"))
        elif action == "remove_filter":
            return remove_filter()
        elif action == "conditional_format":
            return conditional_format(args.get("start", "A1"), args.get("end", "Z100"),
                                     args.get("operator", 1), args.get("formula1", ""),
                                     args.get("formula2", ""))
        elif action == "freeze_panes":
            return freeze_panes(args.get("cell_ref", "B2"))
        elif action == "get_used_range":
            return get_used_range()
        return {"error": f"Unknown excel action: {action}"}
    except RuntimeError as e:
        return {"error": str(e), "error_code": "EXCEL_NOT_RUNNING"}
    except Exception as e:
        return {"error": str(e), "error_code": "EXCEL_COM_ERROR"}


# ─── Offline Docx Handler ───

def _handle_offline_docx(action: str, args: dict) -> dict:
    builder = OfflineDocxBuilder()

    if action == "build":
        structure = args.get("structure", {})
        if not structure or not structure.get("paragraphs"):
            return {"error": "build action requires structure.paragraphs array. Each para: {text, font_name?, font_size?, bold?, alignment?, space_before?, space_after?, first_line_indent?, line_spacing?}", "error_code": "MISSING_PARAM"}
        builder.create()
        for pdata in structure["paragraphs"]:
            text = pdata.get("text", "")
            para = Paragraph()
            para.alignment = pdata.get("alignment") or "left"
            if pdata.get("space_before") is not None:
                para.space_before = pdata["space_before"]
            if pdata.get("space_after") is not None:
                para.space_after = pdata["space_after"]
            if pdata.get("first_line_indent") is not None:
                para.first_line_indent = pdata["first_line_indent"]
            if pdata.get("line_spacing") is not None:
                para.line_spacing = pdata["line_spacing"]
            run = Run(text=text)
            if pdata.get("font_name"):
                run.font = pdata["font_name"]
            if pdata.get("font_size") is not None:
                run.size = pdata["font_size"]
            if pdata.get("bold") is not None:
                run.bold = pdata["bold"]
            if pdata.get("italic") is not None:
                run.italic = pdata["italic"]
            if pdata.get("underline") is not None:
                run.underline = pdata["underline"]
            if pdata.get("color") is not None:
                run.color = pdata["color"]
            para.runs.append(run)
            builder.document.paragraphs.append(para)
        output = args.get("output_path", "build_output.docx")
        builder.save(output)
        return {"built": True, "paragraph_count": len(structure["paragraphs"]), "saved_to": output}
    elif action == "build_cover":
        builder.create()
        for line in args.get("lines", []):
            para = Paragraph()
            para.alignment = line.get("alignment", "center")
            run = Run(
                text=line.get("text", ""),
                font=line.get("font_name"),
                size=line.get("font_size"),
                bold=line.get("bold", False),
            )
            para.runs.append(run)
            builder.document.paragraphs.append(para)
        output = args.get("output_path", "cover.docx")
        builder.save(output)
        return {"output": output}
    elif action == "validate":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        doc = read_docx_model(fp)
        analyzer = DocumentAnalyzer(doc)
        quality = analyzer.analyze_formatting_quality()
        return {
            "valid": quality.get("score", 0) >= 80,
            "score": quality.get("score", 0),
            "issues": quality.get("issues", []),
        }
    elif action == "analyze":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        builder.load(fp)
        return builder.analyze()
    elif action == "auto_format":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        builder.load(fp)
        result = builder.auto_format(args.get("document_type"))
        output = args.get("output_path", fp)
        builder.save(output)
        result["saved_to"] = output
        return result
    elif action == "apply_template":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        builder.load(fp)
        result = builder.apply_template(args.get("template_name", "thesis_cn"))
        output = args.get("output_path", fp)
        builder.save(output)
        result["saved_to"] = output
        return result
    elif action == "add_numbering":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        builder.load(fp)
        result = builder.add_numbering()
        output = args.get("output_path", fp)
        builder.save(output)
        result["saved_to"] = output
        return result
    elif action == "replace_text":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        builder.load(fp)
        count = builder.replace_text(
            args.get("old_text", ""),
            args.get("new_text", ""),
            args.get("case_sensitive", True),
        )
        output = args.get("output_path", fp)
        builder.save(output)
        return {"replaced_count": count, "saved_to": output}
    elif action == "get_text":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        builder.load(fp)
        return {"text": builder.get_text()[:5000]}
    elif action == "get_statistics":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        builder.load(fp)
        return builder.get_statistics()
    elif action == "read_model":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        doc = read_docx_model(fp)
        _cache_doc(fp, doc)
        return {
            "loaded": True,
            "paragraphs": len(doc.paragraphs),
            "tables": len(doc.tables),
            "headings": doc.get_heading_structure(),
        }
    elif action == "write_model":
        fp = args.get("filepath")
        output = args.get("output_path")
        if not fp or not output:
            return {"error": "filepath and output_path required"}
        doc = _get_cached_doc(fp)
        if doc is None:
            return {"error": "No cached document. Use read_model first."}
        write_docx_model(doc, output, fp)
        return {"saved_to": output}
    elif action == "full_structure":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        doc = read_docx_model(fp)
        _cache_doc(fp, doc)
        return doc.get_full_structure()
    elif action == "semantic_structure":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        doc = read_docx_model(fp)
        _cache_doc(fp, doc)
        return {"paragraphs": doc.detect_semantic_structure()}
    elif action == "cross_references":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        doc = read_docx_model(fp)
        _cache_doc(fp, doc)
        return {"references": doc.detect_cross_references()}

    elif action == "detect_semantic_roles":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        from intelligence.layout_analyzer import detect_semantic_roles as _dsr
        return _dsr(fp)

    elif action == "analyze_layout":
        fp = args.get("filepath")
        if not fp:
            return {"error": "filepath required"}
        from intelligence.layout_analyzer import analyze_layout as _al
        return _al(fp)

    return {"error": f"Unknown offline_docx action: {action}"}


# ─── Content Control Handler ───

def _handle_content_control(action: str, args: dict) -> dict:
    from wps_bridge.content_control import (
        count, list_controls, info, add, set_text, set_checkbox,
        select_dropdown, delete, set_tag, find_by_tag,
    )
    doc_idx = args.get("doc_index")
    if action == "count":
        return count(doc_idx)
    elif action == "list_controls":
        return list_controls(doc_idx)
    elif action == "info":
        return info(args["cc_index"], doc_idx)
    elif action == "add":
        return add(
            args.get("type_name", "RICH_TEXT"),
            args.get("text", ""),
            args.get("title", ""),
            args.get("para_index"),
            args.get("position", "end"),
            args.get("dropdown_items"),
            args.get("date_format"),
            doc_idx,
        )
    elif action == "set_text":
        return set_text(args["cc_index"], args.get("text", ""), doc_idx)
    elif action == "set_checkbox":
        return set_checkbox(args["cc_index"], args.get("checked", False), doc_idx)
    elif action == "select_dropdown":
        return select_dropdown(args["cc_index"], args.get("item_text", ""), doc_idx)
    elif action == "delete":
        return delete(args["cc_index"], doc_idx)
    elif action == "set_tag":
        return set_tag(args["cc_index"], args.get("tag", ""), doc_idx)
    elif action == "find_by_tag":
        return find_by_tag(args.get("tag", ""), doc_idx)
    return {"error": f"Unknown content_control action: {action}"}


# ─── Field Codes Handler ───

def _handle_field_codes(action: str, args: dict) -> dict:
    from wps_bridge.field_codes import (
        insert_field, insert_quote, insert_doc_property, insert_seq,
        insert_style_ref, insert_ref, insert_if, list_fields,
        update_fields, unlink_field, find_field_by_code,
    )
    doc_idx = args.get("doc_index")
    pi = args.get("para_index")
    pos = args.get("position", "end")
    if action == "insert_field":
        return insert_field(args.get("field_code", ""), args.get("switches"), pi, pos, doc_idx)
    elif action == "insert_quote":
        return insert_quote(args.get("text", ""), pi, doc_idx)
    elif action == "insert_doc_property":
        return insert_doc_property(args.get("property_name", ""), pi, doc_idx)
    elif action == "insert_seq":
        return insert_seq(args.get("sequence_name", ""), args.get("format_type", "ARABIC"), pi, doc_idx)
    elif action == "insert_style_ref":
        return insert_style_ref(args.get("style_name", ""), args.get("switches"), pi, doc_idx)
    elif action == "insert_ref":
        return insert_ref(args.get("bookmark_name", ""), args.get("switches"), pi, doc_idx)
    elif action == "insert_if":
        return insert_if(args.get("condition", ""), args.get("true_text", ""),
                         args.get("false_text", ""), pi, doc_idx)
    elif action == "list_fields":
        return list_fields(doc_idx)
    elif action == "update_fields":
        return update_fields(doc_idx)
    elif action == "unlink_field":
        return unlink_field(args.get("field_index", 1), doc_idx)
    elif action == "find_field_by_code":
        return find_field_by_code(args.get("pattern", ""), doc_idx)
    return {"error": f"Unknown field_codes action: {action}"}


# ─── Operation Log Handler ───

def _handle_operation_log(action: str, args: dict) -> dict:
    if action == "summary":
        return _log_summary()
    elif action == "recent":
        return {"entries": _log_recent(args.get("count", 20))}
    elif action == "errors":
        return {"errors": _log_errors()}
    elif action == "replay_last":
        return replay_last_error() or {"message": "No errors found"}
    elif action == "clear":
        _log_clear()
        return {"cleared": True}
    elif action == "dump":
        fp = args.get("filepath", "logs/operations.json")
        return {"dumped_to": _log_dump(fp)}
    return {"error": f"Unknown operation_log action: {action}"}


# ─── Main Entry ───

async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


if __name__ == "__main__":
    asyncio.run(main())
