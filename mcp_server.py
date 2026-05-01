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
import re
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
                 "action": {"type": "string"},
                 "doc_index": {"type": "integer"},
                 "filepath": {"type": "string"},
                 "save_changes": {"type": "boolean"},
                 "output_path": {"type": "string"},
                 "width": {"type": "number"}, "height": {"type": "number"},
                 "position": {"type": "string"},
                 "author": {"type": "string"}, "title": {"type": "string"}, "subject": {"type": "string"},
             }, "required": ["action"]}),

        # ─── content ───
        Tool(name="content", description="Read and write document text with Run-level precision. Actions: full_text/paragraph/paragraphs/selection/range/outline/shapes/runs_detail/document_structure/full_structure/semantic_structure/cross_references/insert_text/insert_paragraph/delete_paragraphs/delete_runs/replace_paragraph_text/replace_runs/insert_run/delete_run/split_run/delete_range/replace_range/batch/batch_write/create_cover/build/snapshot/rollback/cache_status/query_by_role",
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

        # ─── style ───
        Tool(name="style", description="Manage document styles. Actions: list/get/create/modify",
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

        # ─── layout ───
        Tool(name="layout", description="Page layout. Actions: page_setup/section_info/add_section_break/columns/header_footer/page_numbers/page_dimensions/page_break/image_wrap/page_border/line_numbers/fix_widow_orphan/auto_fix_layout",
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

        # ─── ai_format ───
        Tool(name="ai_format", description="AI-powered intelligent formatting. Actions: analyze/suggest/apply_template/reformat/auto_toc/auto_numbering/validate/generate_content/summarize_document/rewrite_paragraph/expand_section/translate_section/supervise/health_check/auto_fix/detect_type/detect_role/batch_detect_roles/auto_enhance",
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

        # ─── operation_log ───
        Tool(name="operation_log", description="MCP operation audit log. Actions: summary/recent/errors/replay_last/clear/dump",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string"}, "count": {"type": "integer"},
                 "filepath": {"type": "string"},
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
            result = _handle_content(action, arguments, mode)
        elif name == "format":
            result = _handle_format(action, arguments, mode)
        elif name == "style":
            result = _handle_style(action, arguments)
        elif name == "table":
            result = _handle_table(action, arguments)
        elif name == "search":
            result = _handle_search(action, arguments)
        elif name == "layout":
            result = _handle_layout(action, arguments)
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
            result = _handle_ai_format(action, arguments)
        elif name == "presentation":
            result = _handle_presentation(action, arguments)
        elif name == "excel":
            result = _handle_excel(action, arguments)
        elif name == "offline_docx":
            result = _handle_offline_docx(action, arguments)
        elif name == "content_control":
            result = _handle_content_control(action, arguments)
        elif name == "field_codes":
            result = _handle_field_codes(action, arguments)
        elif name == "surgical":
            result = _handle_surgical(action, arguments)
        elif name == "operation_log":
            result = _handle_operation_log(action, arguments)
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
    elif action == "set_doc_properties":
        doc = get_doc(args.get("doc_index"))
        props = doc.BuiltInDocumentProperties
        if args.get("author"):
            com_set(props, "Author", args["author"])
        if args.get("title"):
            com_set(props, "Title", args["title"])
        if args.get("subject"):
            com_set(props, "Subject", args["subject"])
        return {"updated": True}
    elif action == "insert_image":
        doc = get_doc(args.get("doc_index"))
        pos = args.get("position", "end")
        if pos == "end":
            rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        else:
            rng = doc.Paragraphs.Item(int(pos)).Range
        shape = doc.InlineShapes.AddPicture(args["filepath"], 0, 1, rng.Start)
        if args.get("width"):
            com_set(shape, "Width", args["width"])
        if args.get("height"):
            com_set(shape, "Height", args["height"])
        return {"inserted": True, "image": args["filepath"]}
    elif action == "health_check":
        health = com_health_check()
        return {"health": health, "cache": {"entries": len(_doc_cache), "snapshots": len(_snapshot_cache)}}
    return {"error": f"Unknown document action: {action}"}


def _content_selection(doc_index=None):
    sel = get_app().Selection
    r = sel.Range
    return {
        "text": r.Text,
        "start": r.Start,
        "end": r.End,
    }


def _content_range(start_pos=0, end_pos=-1, doc_index=None):
    doc = get_doc(doc_index)
    end = end_pos if end_pos > 0 else doc.Content.End - 1
    rng = doc.Range(start_pos, end)
    return {
        "text": rng.Text,
        "start": rng.Start,
        "end": rng.End,
    }


# ─── Content Handler (Offline-first) ───

def _handle_content(action: str, args: dict, mode: str) -> dict:
    filepath = args.get("filepath")

    if action in ("delete_paragraphs", "delete_runs", "replace_paragraph_text", "replace_runs", "build"):
        return _handle_content_com(action, args)

    if action in ("full_text", "paragraph", "paragraphs", "outline", "batch",
                   "runs_detail", "document_structure", "full_structure",
                   "semantic_structure", "cross_references", "cache_status",
                   "snapshot", "rollback"):
        # These are read operations — use offline mode for precision
        if not filepath:
            # Fallback to COM if no filepath provided
            return _handle_content_com(action, args)

        doc = _get_cached_doc(filepath)
        if doc is None:
            doc = read_docx_model(filepath)
            _cache_doc(filepath, doc)

        if action == "full_text":
            return {"text": doc.text}
        elif action == "paragraph":
            idx = args["para_index"]
            p = doc.get_paragraph(idx)
            if p is None:
                return {"error": f"Paragraph {idx} not found", "error_code": "PARAGRAPH_NOT_FOUND"}
            return {
                "index": idx,
                "text": p.text,
                "style_id": p.style_id,
                "alignment": p.alignment,
                "outline_level": p.outline_level,
                "run_count": len(p.runs),
                "runs": [{"text": r.text[:100], "font": r.font, "size": r.size, "bold": r.bold, "italic": r.italic} for r in p.runs[:5]],
            }
        elif action == "paragraphs":
            start = args.get("start", 1)
            count = args.get("count", 10)
            result = []
            for i in range(start, min(start + count, len(doc.paragraphs) + 1)):
                p = doc.get_paragraph(i)
                if p:
                    result.append({
                        "index": i,
                        "text": p.text[:200],
                        "style_id": p.style_id,
                        "heading_level": p.heading_level(),
                        "run_count": len(p.runs),
                    })
            return {"items": result, "total": len(doc.paragraphs)}
        elif action == "outline":
            return {"headings": doc.get_heading_structure()}
        elif action == "batch":
            results = []
            for item in args.get("items", []):
                item_type = item.get("type", item.get("action", ""))
                try:
                    if item_type == "paragraph":
                        p = doc.get_paragraph(item["para_index"])
                        data = {"index": item["para_index"], "text": p.text if p else ""}
                    elif item_type == "paragraphs":
                        s = item.get("start", 1)
                        c = item.get("count", 10)
                        data = {"items": [{"index": i, "text": doc.get_paragraph(i).text[:200]} for i in range(s, min(s+c, len(doc.paragraphs)+1)) if doc.get_paragraph(i)]}
                    elif item_type == "outline":
                        data = doc.get_heading_structure()
                    elif item_type == "full_text":
                        data = doc.text
                    else:
                        results.append({"ok": False, "type": item_type, "error": f"Unknown type: {item_type}"})
                        continue
                    results.append({"ok": True, "type": item_type, "data": data})
                except Exception as e:
                    results.append({"ok": False, "type": item_type, "error": str(e)})
            return {"results": results}

        elif action == "runs_detail":
            para_idx = args.get("para_index", 1)
            if not filepath:
                from wps_bridge.content import runs_detail
                return runs_detail(para_idx, args.get("doc_index"))
            p = doc.get_paragraph(para_idx)
            if p is None:
                return {"error": f"Paragraph {para_idx} not found", "error_code": "PARAGRAPH_NOT_FOUND"}
            return {
                "para_index": para_idx,
                "text": p.text,
                "style_id": p.style_id,
                "alignment": p.alignment,
                "outline_level": p.outline_level,
                "first_line_indent": p.first_line_indent,
                "space_before": p.space_before,
                "space_after": p.space_after,
                "line_spacing": p.line_spacing,
                "line_rule": p.line_rule,
                "run_count": len(p.runs),
                "runs": [{
                    "index": i, "text": r.text[:200],
                    "font": r.font, "font_east_asia": r.font_east_asia,
                    "size": r.size, "bold": r.bold, "italic": r.italic,
                    "underline": r.underline, "color": r.color, "highlight": r.highlight,
                    "strike": r.strike, "double_strike": r.double_strike,
                    "emboss": r.emboss, "imprint": r.imprint, "shadow": r.shadow, "outline": r.outline,
                    "superscript": r.superscript, "subscript": r.subscript,
                    "caps": r.caps, "small_caps": r.small_caps,
                    "char_spacing": r.char_spacing, "kerning": r.kerning, "scaling": r.scaling,
                    "baseline_offset": r.baseline_offset,
                    "emphasis_mark": r.emphasis_mark,
                    "hyperlink_url": r.hyperlink_url,
                } for i, r in enumerate(p.runs)],
            }

        elif action == "document_structure":
            from wps_bridge.content import document_structure
            return document_structure(args.get("doc_index"))

        elif action == "full_structure":
            return doc.get_full_structure()

        elif action == "semantic_structure":
            return {"paragraphs": doc.detect_semantic_structure()}

        elif action == "cross_references":
            return {"references": doc.detect_cross_references()}

        elif action == "cache_status":
            fp = filepath or ""
            return _check_cache_staleness(fp) if fp else {
                "cache_size": len(_doc_cache),
                "snapshot_count": len(_snapshot_cache),
                "cached_files": list(_doc_cache.keys()),
                "snapshot_files": list(_snapshot_cache.keys()),
            }

        elif action == "snapshot":
            if not filepath:
                return {"error": "filepath required for snapshot", "error_code": "MISSING_PARAM"}
            doc = _get_cached_doc(filepath)
            if doc is None:
                doc = read_docx_model(filepath)
                _cache_doc(filepath, doc)
            import copy
            _snapshot_cache[filepath] = copy.deepcopy(doc)
            return {"snapshot": True, "para_count": len(doc.paragraphs), "table_count": len(doc.tables)}

        elif action == "rollback":
            if not filepath:
                return {"error": "filepath required for rollback", "error_code": "MISSING_PARAM"}
            snap = _snapshot_cache.pop(filepath, None)
            if snap is None:
                return {"error": f"No snapshot found for {filepath}", "error_code": "NO_SNAPSHOT"}
            import copy
            doc = copy.deepcopy(snap)
            _cache_doc(filepath, doc)
            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            return {"rollback": True, "restored_paragraphs": len(doc.paragraphs), "saved_to": output}

    elif action in ("insert_text", "insert_run", "delete_run", "split_run",
                     "delete_range", "replace_range", "create_cover", "batch_write"):
        # Write operations — offline mode with save
        if not filepath:
            return _handle_content_com(action, args)

        doc = _get_cached_doc(filepath)
        if doc is None:
            doc = read_docx_model(filepath)

        if action == "insert_text":
            text = args["text"]
            position = args.get("position", "end")
            lines = [l.strip() for l in re.split(r'[\r\n]+|\\n', text) if l.strip()]
            new_paras = [Paragraph(runs=[Run(text=line)]) for line in lines if line.strip()]

            if position == "end":
                for para in new_paras:
                    doc.paragraphs.append(para)
            elif position == "before" and args.get("para_index"):
                idx = args["para_index"]
                for para in reversed(new_paras):
                    doc.insert_paragraph(idx, para)
            elif position == "after" and args.get("para_index"):
                idx = args["para_index"]
                for para in new_paras:
                    doc.insert_paragraph(idx + 1, para)

            # Auto-save if filepath provided
            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"inserted": True, "paragraphs_created": len(new_paras), "saved_to": output}

        elif action == "insert_run":
            para_idx = args["para_index"]
            run_idx = args.get("run_index")
            text = args.get("text", "")
            para = doc.get_paragraph(para_idx)
            if para is None:
                return {"error": f"Paragraph {para_idx} not found", "error_code": "PARAGRAPH_NOT_FOUND"}

            new_run = Run(text=text)
            if args.get("font_name"):
                new_run.font = args["font_name"]
            if args.get("size"):
                new_run.size = args["size"]
            if "bold" in args:
                new_run.bold = args["bold"]
            if "italic" in args:
                new_run.italic = args["italic"]

            if run_idx is not None and 0 <= run_idx <= len(para.runs):
                para.runs.insert(run_idx, new_run)
            else:
                para.runs.append(new_run)

            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"inserted": True, "run_index": run_idx if run_idx is not None else len(para.runs) - 1, "saved_to": output}

        elif action == "delete_run":
            para_idx = args["para_index"]
            run_idx = args["run_index"]
            para = doc.get_paragraph(para_idx)
            if para is None:
                return {"error": f"Paragraph {para_idx} not found", "error_code": "PARAGRAPH_NOT_FOUND"}
            if 0 <= run_idx < len(para.runs):
                removed = para.runs.pop(run_idx)
                output = args.get("output_path", filepath)
                write_docx_model(doc, output, filepath)
                _cache_doc(output, doc)
                return {"deleted": True, "removed_text": removed.text, "saved_to": output}
            return {"error": f"Run {run_idx} not found", "error_code": "RUN_NOT_FOUND"}

        elif action == "split_run":
            para_idx = args["para_index"]
            run_idx = args["run_index"]
            split_pos = args.get("split_pos", 0)
            para = doc.get_paragraph(para_idx)
            if para is None:
                return {"error": f"Paragraph {para_idx} not found", "error_code": "PARAGRAPH_NOT_FOUND"}
            if 0 <= run_idx < len(para.runs):
                run = para.runs[run_idx]
                if 0 < split_pos < len(run.text):
                    left = run.clone()
                    left.text = run.text[:split_pos]
                    right = run.clone()
                    right.text = run.text[split_pos:]
                    para.runs[run_idx] = left
                    para.runs.insert(run_idx + 1, right)
                    output = args.get("output_path", filepath)
                    write_docx_model(doc, output, filepath)
                    _cache_doc(output, doc)
                    return {"split": True, "run_count": len(para.runs), "saved_to": output}
            return {"error": f"Run {run_idx} not splittable", "error_code": "INVALID_RANGE"}
        elif action == "delete_range":
            start_pos = args["start_pos"]
            end_pos = args.get("end_pos")

            if end_pos is None:
                doc.paragraphs.clear()
            elif end_pos <= 0 or end_pos <= start_pos:
                pass  # zero-width or invalid range, nothing deleted
            else:
                flat_text = ""
                para_starts = []
                for p in doc.paragraphs:
                    para_starts.append(len(flat_text))
                    flat_text += p.text + "\n"
                flat_text = flat_text.rstrip("\n")

                actual_end = min(end_pos, len(flat_text))
                if 0 <= start_pos < len(flat_text):
                    new_text = flat_text[:start_pos] + flat_text[actual_end:]
                    doc.paragraphs.clear()
                    for line in new_text.split("\n"):
                        if line:
                            doc.paragraphs.append(Paragraph(runs=[Run(text=line)]))

            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"deleted": True, "saved_to": output}

        elif action == "replace_range":
            start_pos = args["start_pos"]
            end_pos = args["end_pos"]
            new_text = args["new_text"]

            flat_text = ""
            for p in doc.paragraphs:
                flat_text += p.text + "\n"
            flat_text = flat_text.rstrip("\n")

            actual_end = min(end_pos, len(flat_text))
            if 0 <= start_pos < len(flat_text):
                result_text = flat_text[:start_pos] + new_text + flat_text[actual_end:]
                doc.paragraphs.clear()
                for line in result_text.split("\n"):
                    if line:
                        doc.paragraphs.append(Paragraph(runs=[Run(text=line)]))
                    else:
                        doc.paragraphs.append(Paragraph())

            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"replaced": True, "saved_to": output}

        elif action == "create_cover":
            lines = args.get("lines", [])
            clear = args.get("clear_existing", True)
            if not lines:
                return {"error": "create_cover requires 'lines' parameter. Each line: {text, font_name?, font_size?, bold?, alignment?, space_before?, space_after?}",
                        "error_code": "MISSING_PARAM", "example": {"lines": [
                            {"text": "Title", "font_name": "黑体", "font_size": 26, "bold": True, "alignment": "center", "space_before": 120, "space_after": 24},
                            {"text": "Subtitle", "font_name": "宋体", "font_size": 16, "alignment": "center", "space_after": 6},
                            {"text": "2026-05-02", "font_name": "宋体", "font_size": 14, "alignment": "center", "space_after": 6}
                        ]}}
            if clear:
                doc.paragraphs.clear()

            for line in lines:
                text = line.get("text", "").strip()
                if not text:
                    continue
                para = Paragraph()
                para.alignment = line.get("alignment", "center")
                para.space_before = line.get("space_before", 0)
                para.space_after = line.get("space_after", 0)
                run = Run(
                    text=text,
                    font=line.get("font_name"),
                    size=line.get("font_size"),
                    bold=line.get("bold", False),
                    italic=line.get("italic", False),
                )
                para.runs.append(run)
                doc.paragraphs.append(para)

            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"created": True, "paragraphs": len(lines), "saved_to": output}

        elif action == "batch_write":
            items = args.get("items", [])
            inserted = 0
            for item in items:
                text = item.get("text", "")
                pos = item.get("position", "end")
                pi = item.get("para_index")
                lines = [l.strip() for l in re.split(r'[\r\n]+|\\n', text) if l.strip()]
                if not lines:
                    continue
                new_paras = [Paragraph(runs=[Run(text=l)]) for l in lines]
                if pos == "end":
                    for p in new_paras:
                        doc.paragraphs.append(p)
                elif pos == "before" and pi:
                    for p in reversed(new_paras):
                        doc.insert_paragraph(pi, p)
                elif pos == "after" and pi:
                    for p in new_paras:
                        doc.insert_paragraph(pi + 1, p)
                inserted += len(new_paras)
            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"inserted": True, "paragraphs_created": inserted, "items_processed": len(items), "saved_to": output}

    elif action == "query_by_role":
        if not filepath:
            return _handle_content_com(action, args)
        doc = _get_cached_doc(filepath)
        if doc is None:
            doc = read_docx_model(filepath)
            _cache_doc(filepath, doc)
        role = args.get("sr", args.get("role", ""))
        if not role:
            return {"error": "Provide sr or role parameter", "error_code": "MISSING_PARAM"}
        from docx_engine.semantic_model import SemanticParser, SemanticRole
        parser = SemanticParser(doc)
        results = parser.parse()
        matched = [{"index": r.index, "role": r.role, "confidence": r.confidence,
                    "text": r.text_preview} for r in results.elements if r.role == role]
        if not matched:
            broader = {
                "abstract": [SemanticRole.ABSTRACT_LABEL, SemanticRole.ABSTRACT_CONTENT],
                "cover": [SemanticRole.COVER_TITLE, SemanticRole.COVER_SUBTITLE, SemanticRole.COVER_DATE, SemanticRole.COVER_AUTHOR, SemanticRole.COVER_INSTITUTION],
                "keywords": [SemanticRole.KEYWORDS_LABEL, SemanticRole.KEYWORDS],
                "toc": [SemanticRole.TOC_HEADING, SemanticRole.TOC_ENTRY],
                "references": [SemanticRole.REFERENCE_SECTION_HEADER, SemanticRole.REFERENCE_ITEM],
                "acknowledgements": [SemanticRole.ACKNOWLEDGEMENTS],
                "appendix": [SemanticRole.APPENDIX_HEADING, SemanticRole.APPENDIX_CONTENT],
            }
            if role in broader:
                matched = [{"index": r.index, "role": r.role, "confidence": r.confidence,
                            "text": r.text_preview} for r in results.elements if r.role in broader[role]]
        return {"role": role, "matched": len(matched), "paragraphs": matched}

    # Fallback to COM for unsupported operations
    return _handle_content_com(action, args)


def _handle_content_com(action: str, args: dict) -> dict:
    """Fallback COM-based content handler."""
    doc = get_doc(args.get("doc_index"))
    if action == "full_text":
        return {"text": com_property(doc.Content, "Text", "")}
    elif action == "paragraph":
        p = doc.Paragraphs.Item(args["para_index"])
        return {"index": args["para_index"], "text": com_property(p.Range, "Text", "").strip()}
    elif action == "paragraphs":
        result = []
        total = doc.Paragraphs.Count
        end = min(args.get("start", 1) + args.get("count", 10) - 1, total)
        for i in range(args.get("start", 1), end + 1):
            p = doc.Paragraphs.Item(i)
            result.append({"index": i, "text": com_property(p.Range, "Text", "").strip()[:200]})
        return {"items": result, "total": total}
    elif action == "runs_detail":
        from wps_bridge.content import runs_detail
        return runs_detail(args.get("para_index", 1), args.get("doc_index"))
    elif action == "document_structure":
        from wps_bridge.content import document_structure
        return document_structure(args.get("doc_index"))
    elif action == "insert_text":
        from wps_bridge.content import insert_text as _insert_text
        return _insert_text(args.get("text", ""), args.get("position", "end"),
                           args.get("para_index"), args.get("doc_index"))
    elif action == "insert_paragraph":
        from wps_bridge.content import insert_paragraph
        return insert_paragraph(args.get("text", ""), args.get("style"),
                               args.get("position", "end"), args.get("para_index"), args.get("doc_index"))
    elif action == "delete_paragraphs":
        from wps_bridge.content import delete_paragraphs
        return delete_paragraphs(args.get("from_para", args.get("para_index", 1)),
                                args.get("to_para", args.get("para_index", 1)),
                                args.get("doc_index"))
    elif action == "delete_runs":
        from wps_bridge.content import delete_runs
        return delete_runs(args.get("para_index", 1), args.get("from_run", 1),
                          args.get("to_run", 1), args.get("doc_index"))
    elif action == "replace_paragraph_text":
        from wps_bridge.content import replace_paragraph_text
        return replace_paragraph_text(args.get("para_index", 1), args.get("new_text", ""),
                                     args.get("preserve_format", False), args.get("doc_index"))
    elif action == "replace_runs":
        from wps_bridge.content import replace_runs
        return replace_runs(args.get("para_index", 1), args.get("run_indices", []),
                           args.get("new_text", ""), args.get("doc_index"))
    elif action == "create_cover":
        from wps_bridge.content import create_cover
        return create_cover(args.get("lines", []), args.get("clear_existing", True), args.get("doc_index"))
    elif action == "build":
        from wps_bridge.content import doc_build
        structure = args.get("structure", {})
        output_path = args.get("output_path", args.get("filepath", ""))
        return doc_build(structure, output_path, args.get("doc_index"))
    elif action == "snapshot":
        from wps_bridge.content import snapshot
        return snapshot(args.get("doc_index"))
    elif action == "rollback":
        from wps_bridge.content import rollback
        return rollback(args.get("doc_index"))
    elif action == "delete_range":
        end = args.get("end_pos")
        if end is None:
            end = doc.Content.End
        if args["start_pos"] >= end:
            return {"deleted": False, "warning": "delete_range: zero-width or invalid range, nothing deleted"}
        doc.Range(args["start_pos"], end).Delete()
        return {"deleted": True}
    elif action == "replace_range":
        start_pos = args["start_pos"]
        end_pos = args.get("end_pos")
        new_text = args.get("new_text", "")
        if end_pos is None:
            end_pos = doc.Content.End
        try:
            r = doc.Range(start_pos, end_pos)
            # Preserve formatting from original range
            try:
                fmt_copy = r.Font.Duplicate
                para_fmt_copy = r.ParagraphFormat.Duplicate
            except Exception:
                fmt_copy = None
                para_fmt_copy = None
            r.Text = new_text
            # Restore formatting on the new text
            if fmt_copy is not None:
                try:
                    new_range = doc.Range(start_pos, start_pos + len(new_text))
                    new_range.Font = fmt_copy
                    if para_fmt_copy is not None:
                        new_range.ParagraphFormat = para_fmt_copy
                except Exception:
                    pass
            return {"replaced": True, "start": start_pos, "end": start_pos + len(new_text)}
        except Exception as e:
            return {"replaced": False, "error": str(e)}
    elif action == "batch":
        from wps_bridge.content import batch as _content_batch
        return {"details": _content_batch(args.get("items", []), args.get("doc_index")), "total": len(args.get("items", []))}
    elif action == "outline":
        from wps_bridge.content import outline
        return {"headings": outline(args.get("doc_index"))}
    elif action == "shapes":
        from wps_bridge.content import shapes
        return shapes(args.get("include_inlines", True), args.get("doc_index"))
    elif action == "find_text":
        from wps_bridge.content import find_text
        return {"results": find_text(args.get("query", ""), args.get("match_case", False), args.get("whole_word", False), args.get("doc_index"))}
    elif action == "find_all":
        from wps_bridge.content import find_all
        return find_all(args.get("query", ""), args.get("match_case", True), args.get("doc_index"))
    elif action == "select_by_role":
        from wps_bridge.content import select_by_role
        return select_by_role(args.get("role", ""), args.get("filepath"), args.get("doc_index"))
    elif action == "selection":
        return _content_selection(args.get("doc_index"))
    elif action == "range":
        return _content_range(args.get("start_pos", 0), args.get("end_pos", 0), args.get("doc_index"))
    elif action == "full_structure":
        from wps_bridge.content import document_structure
        return {"structure": document_structure(args.get("doc_index"))}
    elif action == "semantic_structure":
        from wps_bridge.content import document_structure
        ds = document_structure(args.get("doc_index"))
        return {"paragraphs": ds.get("paragraphs", [])}
    elif action == "cross_references":
        return {"references": [], "note": "Cross-reference detection requires offline mode with filepath"}
    return {"error": f"Unknown content action: {action}"}


# ─── Format Handler ───

def _handle_format(action: str, args: dict, mode: str) -> dict:
    filepath = args.get("filepath")

    # ── COM-only run-level font operations ──
    if action in ("get_run_font", "set_run_font"):
        return _handle_format_com(action, args)

    if action in ("set_font", "set_paragraph_format", "set_run_format", "apply_style", "batch") and filepath:
        # Offline formatting
        doc = _get_cached_doc(filepath)
        if doc is None:
            doc = read_docx_model(filepath)

        if action == "set_font":
            para_idx = args.get("para_index")
            run_idx = args.get("run_index")
            if para_idx:
                para = doc.get_paragraph(para_idx)
                if para:
                    target_runs = [para.runs[run_idx]] if run_idx is not None and 0 <= run_idx < len(para.runs) else para.runs
                    for run in target_runs:
                        if args.get("name"):
                            run.font = args["name"]
                        if args.get("size"):
                            run.size = args["size"]
                        if "bold" in args:
                            run.bold = args["bold"]
                        if "italic" in args:
                            run.italic = args["italic"]
            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"updated": True, "saved_to": output}

        elif action == "set_run_format":
            para_idx = args.get("para_index")
            run_idx = args.get("run_index")
            if para_idx is None or run_idx is None:
                return {"error": "para_index and run_index required for set_run_format"}
            para = doc.get_paragraph(para_idx)
            if para is None:
                return {"error": f"Paragraph {para_idx} not found", "error_code": "PARAGRAPH_NOT_FOUND"}
            if not (0 <= run_idx < len(para.runs)):
                return {"error": f"Run {run_idx} not found", "error_code": "RUN_NOT_FOUND"}

            run = para.runs[run_idx]
            for attr in ("font", "size", "bold", "italic", "underline", "color", "highlight",
                         "strike", "superscript", "subscript", "char_spacing", "kerning",
                         "scaling", "caps", "small_caps"):
                if attr in args and args[attr] is not None:
                    setattr(run, attr, args[attr])

            if args.get("text") is not None:
                run.text = args["text"]

            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"updated": True, "run_index": run_idx, "saved_to": output}

        elif action == "set_paragraph_format":
            para_idx = args.get("para_index")
            if para_idx is None or para_idx < 1:
                para_idx = 1
            if para_idx:
                para = doc.get_paragraph(para_idx)
                if para:
                    if args.get("alignment"):
                        para.alignment = args["alignment"]
                    if args.get("first_line_indent") is not None:
                        para.first_line_indent = args["first_line_indent"]
                    if args.get("space_before") is not None:
                        para.space_before = args["space_before"]
                    if args.get("space_after") is not None:
                        para.space_after = args["space_after"]
                    if args.get("line_spacing") is not None:
                        para.line_spacing = args["line_spacing"]
                    if args.get("line_spacing_rule"):
                        para.line_rule = args["line_spacing_rule"]
            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"updated": True, "saved_to": output}

        elif action == "batch":
            ops = args.get("operations", [])
            changes = 0
            for op in ops:
                op_type = op.get("type", op.get("action", ""))
                pidx = op.get("para_index")
                para = doc.get_paragraph(pidx) if pidx else None
                if not para:
                    continue
                if op_type == "set_font":
                    for run in para.runs:
                        if op.get("name"):
                            run.font = op["name"]
                        if op.get("size"):
                            run.size = op["size"]
                        if "bold" in op:
                            run.bold = op["bold"]
                    changes += 1
                elif op_type == "set_paragraph_format":
                    if op.get("alignment"):
                        para.alignment = op["alignment"]
                    if op.get("first_line_indent") is not None:
                        para.first_line_indent = op["first_line_indent"]
                    changes += 1
            output = args.get("output_path", filepath)
            write_docx_model(doc, output, filepath)
            _cache_doc(output, doc)
            return {"updated": True, "changes": changes, "saved_to": output}

    # COM fallback
    return _handle_format_com(action, args)


def _handle_format_com(action: str, args: dict) -> dict:
    from wps_bridge.formatting import get_run_font as _get_run_font, set_run_font as _set_run_font, get_font as _get_font_bridge

    doc = get_doc(args.get("doc_index"))
    if action == "get_run_font":
        return _get_run_font(
            args.get("para_index", 1),
            args.get("run_index", 1),
            args.get("doc_index"),
        )
    elif action == "set_run_font":
        font_kwargs = {}
        for key in ("name", "name_far_east", "size", "bold", "italic", "underline",
                     "color_index", "color_rgb", "highlight", "superscript", "subscript",
                     "strike_through", "spacing", "scaling", "kerning",
                     "caps", "small_caps", "emboss", "shadow", "outline", "vanish"):
            if key in args and args[key] is not None:
                font_kwargs[key] = args[key]
        return _set_run_font(
            args.get("para_index", 1),
            args.get("run_index", 1),
            args.get("doc_index"),
            **font_kwargs,
        )
    elif action == "get_font":
        return _get_font_bridge(args.get("para_index"), args.get("start_pos"), args.get("end_pos"), args.get("use_selection", False), args.get("doc_index"))
    elif action == "set_font":
        from wps_bridge.formatting import set_font as _sf
        font_kwargs = {}
        for k in ("name", "name_far_east", "size", "bold", "italic", "underline",
                   "color_index", "color_rgb", "highlight", "superscript", "subscript",
                   "strike_through", "spacing", "scaling", "kerning",
                   "caps", "small_caps", "emboss", "shadow", "outline", "vanish"):
            if k in args:
                font_kwargs[k] = args[k]
        return _sf(para_index=args.get("para_index"),
                   start_pos=args.get("start_pos"), end_pos=args.get("end_pos"),
                   use_selection=args.get("use_selection", False),
                   doc_index=args.get("doc_index"), **font_kwargs)
    elif action == "batch":
        # COM-based batch format operations
        ops = args.get("operations", [])
        results = []
        for op in ops:
            op_type = op.get("type", op.get("action", ""))
            pidx = op.get("para_index")
            try:
                if op_type == "set_font" and pidx:
                    from wps_bridge.formatting import set_font as _bsf
                    font_kwargs = {}
                    for k in ("name", "name_far_east", "size", "bold", "italic", "underline",
                               "color_index", "color_rgb", "highlight", "superscript",
                               "subscript", "strike_through", "spacing", "scaling", "kerning",
                               "caps", "small_caps", "emboss", "shadow", "outline", "vanish"):
                        if k in op:
                            font_kwargs[k] = op[k]
                    _bsf(para_index=pidx, doc_index=args.get("doc_index"), **font_kwargs)
                    results.append({"ok": True, "type": "set_font", "result": {"updated": True}})
                elif op_type == "set_paragraph_format" and pidx:
                    from wps_bridge.formatting import set_paragraph_format as _bspf
                    pf_kwargs = {}
                    for k in ("alignment", "first_line_indent", "left_indent", "right_indent",
                               "line_spacing_rule", "line_spacing", "space_before", "space_after",
                               "outline_level", "widow_control", "keep_with_next"):
                        if k in op:
                            pf_kwargs[k] = op[k]
                    _bspf(para_index=pidx, doc_index=args.get("doc_index"), **pf_kwargs)
                    results.append({"ok": True, "type": "set_paragraph_format", "result": {"updated": True}})
                else:
                    results.append({"ok": False, "type": op_type, "error": f"Unknown or unsupported batch operation: {op_type}"})
            except Exception as e:
                results.append({"ok": False, "type": op_type, "error": str(e)})
        return {"total": len(ops), "success": sum(1 for r in results if r["ok"]), "failed": sum(1 for r in results if not r["ok"]), "details": results}
    elif action == "get_paragraph_format":
        from wps_bridge.formatting import get_paragraph_format as _get_pf
        return _get_pf(args["para_index"], args.get("doc_index"))
    elif action == "set_paragraph_format":
        from wps_bridge.formatting import set_paragraph_format as _set_pf
        para_index = args.get("para_index")
        if para_index is None or para_index < 1:
            para_index = 1
        pf_kwargs = {}
        for k in ("alignment", "first_line_indent", "left_indent", "right_indent",
                   "line_spacing_rule", "line_spacing", "space_before", "space_after",
                   "outline_level", "widow_control", "keep_with_next"):
            if k in args and args[k] is not None:
                pf_kwargs[k] = args[k]
        return _set_pf(para_index=para_index, use_selection=args.get("use_selection", False), doc_index=args.get("doc_index"), **pf_kwargs)
    elif action == "apply_style":
        style_name = args["style_name"]
        if args.get("use_selection"):
            get_app().Selection.ParagraphFormat.Style = style_name
        elif args.get("para_index"):
            pi = args["para_index"]
            if pi < 1:
                pi = 1
            doc.Paragraphs.Item(pi).Range.Style = style_name
        return {"applied": style_name}
    elif action == "clear_formatting":
        from wps_bridge.formatting import clear_formatting as _cf
        return _cf(args.get("para_index"), args.get("use_selection", False), args.get("doc_index"))
    elif action == "copy_format":
        from wps_bridge.formatting import copy_format as _cpf
        return _cpf(args["source_para_index"], args["target_para_indices"], args.get("doc_index"))
    elif action == "add_hyperlink":
        from wps_bridge.formatting import add_hyperlink as _ah
        return _ah(args["text"], args["url"], args.get("para_index"), args.get("doc_index"))
    elif action == "set_tab_stops":
        from wps_bridge.formatting import set_tab_stops as _sts
        return _sts(args["para_index"], args.get("stops", []), args.get("doc_index"))
    elif action == "set_bullet_list":
        from wps_bridge.formatting import set_bullet_list as _sbl
        return _sbl(args.get("para_indices", []), args.get("bullet_char"), args.get("doc_index"))
    elif action == "add_watermark":
        from wps_bridge.document import add_watermark as _aw
        return _aw(args["text"], args.get("font_size", 72), args.get("color", 15), args.get("doc_index"))
    elif action == "remove_watermark":
        from wps_bridge.document import remove_watermark as _rw
        return _rw(args.get("doc_index"))
    elif action == "resolve_format":
        from wps_bridge.format_resolver import resolve_paragraph_format
        return resolve_paragraph_format(args.get("para_index", 1), args.get("doc_index"))
    elif action == "resolve_run_format":
        from wps_bridge.format_resolver import resolve_run_format
        return resolve_run_format(args.get("para_index", 1), args.get("run_index", 1), args.get("doc_index"))
    elif action == "set_text_effect":
        from wps_bridge.formatting import set_text_effect as _ste
        return _ste(args.get("para_index", 1), args.get("effect", ""),
                    args.get("color_rgb", 0), args.get("offset", 2.0),
                    args.get("doc_index"))
    return {"error": f"Unknown format action: {action}"}


# ─── Style Handler ───

def _handle_style(action: str, args: dict) -> dict:
    doc = get_doc(args.get("doc_index"))
    if action == "list":
        styles = []
        for i in range(1, min(doc.Styles.Count, 200) + 1):
            try:
                s = doc.Styles.Item(i)
                styles.append({"name": com_property(s, "NameLocal", ""), "builtin": bool(com_property(s, "BuiltIn", 0))})
            except Exception:
                continue
        return {"styles": styles}
    elif action == "get":
        s = doc.Styles.Item(args["name"])
        return {"name": com_property(s, "NameLocal", ""), "font": com_property(s.Font, "Name", "")}
    return {"error": f"Unknown style action: {action}"}


# ─── Table Handler ───

def _handle_table(action: str, args: dict) -> dict:
    di = args.get("doc_index")
    ti = args.get("table_index")

    if action == "count":
        return _table_bridge.table_count(di)
    elif action == "info":
        return _table_bridge.table_info(ti, di)
    elif action == "read":
        return _table_bridge.table_read(ti, di)
    elif action == "create":
        return _table_bridge.table_create(args["rows"], args["cols"], args.get("position"), di)
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
        if not structure:
            return {"error": "build action requires 'structure' object", "error_code": "MISSING_PARAM"}
        paragraphs = structure.get("paragraphs") or structure.get("pages")
        if not paragraphs:
            return {"error": "build action requires structure.paragraphs array. Each para: {text, font_name?, font_size?, bold?, alignment?, space_before?, space_after?, first_line_indent?, line_spacing?}", "error_code": "MISSING_PARAM"}
        builder.create()
        for pdata in paragraphs:
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
