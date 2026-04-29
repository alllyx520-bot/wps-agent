# -*- coding: utf-8 -*-
import asyncio
import json
import logging
import os
import sys
from pathlib import Path

import pythoncom
import yaml
from mcp.server import Server, NotificationOptions
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

sys.path.insert(0, str(Path(__file__).parent))

from wps_bridge.app import get_app, get_doc as _bridge_get_doc
from wps_bridge import document, content, formatting, table, layout, search, review, docspace, transfer, migrate, compare, ppt_app

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("wps-agent")

PROJECT_DIR = Path(__file__).parent
CONFIG_PATH = PROJECT_DIR / "config.yaml"
try:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        CONFIG = yaml.safe_load(f) or {}
except Exception:
    CONFIG = {}

server = Server(CONFIG["server"]["name"])


@server.list_tools()
async def list_tools():
    return [
        # --- document ---
        Tool(name="document", description="WPS Word document management. Use when user asks to: create/open/save/close/list Word documents, export to PDF, switch between open docs, insert images, set document properties. Actions: info/list/open/create/save/close/activate/export_pdf/insert_image/doc_properties/set_doc_properties",
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
        Tool(name="content", description="Read and write WPS Word document text. Use when user asks to: read/view/show document text, get paragraph content, check selection, see document outline/structure, insert/delete/replace text. IMPORTANT: Use batch action to read multiple items in ONE call (e.g. batch with types paragraph+outline). For cover pages, use create_cover (single call creates and formats all lines). Actions: full_text/paragraph/paragraphs/selection/range/outline/insert_text/delete_range/replace_range/batch/create_cover",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "full_text (get entire doc), paragraph (get one by index), paragraphs (get range), selection (current cursor), range (by start/end position), outline (heading structure), insert_text, delete_range, replace_range, batch (read multiple items at once), create_cover (single-call cover page, pass lines array of {text, font_name, font_size, bold, alignment, space_before, space_after...})"},
                 "para_index": {"type": "integer", "description": "Paragraph number (1-based)"},
                 "start": {"type": "integer", "description": "Start paragraph index for paragraphs action"},
                 "count": {"type": "integer", "description": "How many paragraphs to return"},
                 "start_pos": {"type": "integer", "description": "Character start position for range/delete/replace"},
                 "end_pos": {"type": "integer", "description": "Character end position for range/delete/replace (optional for delete_range, defaults to doc end)"},
                 "text": {"type": "string", "description": "Text to insert"},
                 "new_text": {"type": "string", "description": "Replacement text"},
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
        Tool(name="format", description="Get or set font and paragraph formatting in WPS Word. Use when user asks to: change font/size/bold/italic/color, modify paragraph alignment/indentation/line spacing, apply styles, clear formatting, use format painter, add/remove watermark. IMPORTANT: Use batch action to modify/read multiple paragraphs in ONE call. Supports Chinese font names: 黑体, 宋体, 仿宋, 楷体, 微软雅黑. Alignment: left/center/right/justify. Line spacing: single/1.5lines/double/exactly/multiple. Actions: get_font/set_font/get_paragraph_format/set_paragraph_format/apply_style/clear_formatting/copy_format/batch/add_watermark/remove_watermark",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "get_font (read font), set_font (modify font), get_paragraph_format (read paragraph), set_paragraph_format (modify paragraph), apply_style (apply named style), clear_formatting (reset), copy_format (format painter from source to targets), batch (execute multiple operations at once), add_watermark (text watermark), remove_watermark (delete all watermarks)"},
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
             }, "required": ["action"]}),

        # --- style ---
        Tool(name="style", description="Manage WPS Word styles (样式). Use when user asks to: list available styles, view style details, create new styles, or modify existing styles. Style names use Chinese in WPS: 标题 1, 标题 2, 正文, etc. Actions: list/get/create/modify",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "list (all styles), get (details of one style), create (new style), modify (change existing style)"},
                 "name": {"type": "string", "description": "Style name, e.g. 标题 1, 正文, or custom name"}, "base_style": {"type": "string", "description": "Base style name to inherit from"},
                 "font_name": {"type": "string", "description": "Font name for the style"}, "font_size": {"type": "number", "description": "Font size in points"},
                 "bold": {"type": "boolean"}, "italic": {"type": "boolean"},
                 "alignment": {"type": "string", "description": "left, center, right, justify"},
                 "first_line_indent": {"type": "number"},
                 "line_spacing_rule": {"type": "string"}, "line_spacing": {"type": "number"},
                 "space_before": {"type": "number"}, "space_after": {"type": "number"},
                 "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # --- table ---
        Tool(name="table", description="Create and manipulate Word tables. Use when user asks to: insert/delete tables, fill table cells, format tables with headers/borders/colors, merge cells, adjust column widths, apply alternating row colors. Table index is 1-based. IMPORTANT: Use batch_read to read multiple tables in ONE call. Actions: count/info/read/create/delete/set_cell_text/format_cell/set_header/format_borders/merge_cells/auto_fit/set_column_width/alternate_rows/batch_read",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "count/info/read/create/delete/set_cell_text (fill cell)/format_cell/set_header (bold header row)/format_borders/merge_cells/auto_fit/set_column_width/alternate_rows/batch_read (read multiple tables at once)"},
                 "table_index": {"type": "integer", "description": "Table number (1-based)"}, "rows": {"type": "integer", "description": "Number of rows for create"}, "cols": {"type": "integer", "description": "Number of columns for create"},
                 "row": {"type": "integer", "description": "Row number (1-based)"}, "col": {"type": "integer", "description": "Column number (1-based)"},
                 "text": {"type": "string", "description": "Text to set in cell"}, "font_name": {"type": "string"}, "font_size": {"type": "number"},
                 "bold": {"type": "boolean"}, "align": {"type": "string", "description": "Cell text alignment"}, "shading_color": {"type": "integer"},
                 "row_count": {"type": "integer", "description": "Number of header rows"}, "width": {"type": "number", "description": "Column width in points"},
                 "start_row": {"type": "integer"}, "start_col": {"type": "integer"}, "end_row": {"type": "integer"}, "end_col": {"type": "integer"},
                 "behavior": {"type": "integer", "description": "Auto fit behavior: 1=content 2=window"}, "color1": {"type": "string", "description": "Even row hex color"}, "color2": {"type": "string", "description": "Odd row hex color"},
                 "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # --- search ---
        Tool(name="search", description="Find and replace text in WPS Word. Use when user asks to: search for words/phrases, find and replace text, search by formatting, or jump to headings. Actions: find/replace/find_format/goto_heading",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "find (search text), replace (replace text), find_format (search by font/style), goto_heading (navigate to heading)"},
                 "query": {"type": "string", "description": "Text to search for"}, "find_text": {"type": "string", "description": "Text to find in replace"}, "replace_text": {"type": "string", "description": "Replacement text"},
                 "match_case": {"type": "boolean"}, "whole_word": {"type": "boolean"},
                 "replace_all": {"type": "boolean", "description": "If true, replace all occurrences"},
                 "font_name": {"type": "string", "description": "Font name to search for"}, "font_size": {"type": "number", "description": "Font size to search for"},
                 "style_name": {"type": "string", "description": "Style name to search for"}, "text": {"type": "string", "description": "Heading text for goto_heading"}, "level": {"type": "integer", "description": "Heading level for goto_heading"},
                 "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # --- layout ---
        Tool(name="layout", description="Configure WPS Word page layout. Use when user asks to: change paper size (A4), adjust margins, switch portrait/landscape, add section breaks, set columns, edit headers/footers, or add page numbers. Actions: page_setup/section_info/add_section_break/columns/header_footer/page_numbers",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "page_setup (margins/paper/orientation), section_info (current section details), add_section_break (insert break), columns (set column count), header_footer (set text), page_numbers (add/configure)"},
                 "section_index": {"type": "integer", "description": "Section number (1-based, optional)"},
                 "page_width": {"type": "number", "description": "Page width in points (A4=595.3)"}, "page_height": {"type": "number", "description": "Page height in points (A4=841.9)"},
                 "top_margin": {"type": "number", "description": "Top margin in points (72pt=2.54cm)"}, "bottom_margin": {"type": "number"},
                 "left_margin": {"type": "number"}, "right_margin": {"type": "number"},
                 "orientation": {"type": "integer", "description": "0=portrait 1=landscape"}, "gutter": {"type": "number", "description": "Gutter margin"},
                 "different_first_page": {"type": "boolean", "description": "Different header/footer for first page"},
                 "para_index": {"type": "integer", "description": "Paragraph to insert break after"}, "break_type": {"type": "string", "description": "next_page/continuous/even_page/odd_page"},
                 "count": {"type": "integer", "description": "Number of columns"},
                 "header_type": {"type": "string", "description": "header or footer"},
                 "text": {"type": "string", "description": "Text for header or footer"},
                 "alignment": {"type": "string", "description": "Page number alignment: left/center/right"}, "start_at": {"type": "integer", "description": "Starting page number"},
                 "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # --- review ---
        Tool(name="review", description="Track changes and comments in WPS Word. Use when user asks to: turn on/off Track Changes, add/view/delete comments (批注), accept/reject revisions (修订). Actions: track_changes_toggle/track_changes_status/comments_list/comment_add/revisions_list/revisions_accept_all/revisions_reject_all",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "track_changes_toggle (turn on/off), track_changes_status (check), comments_list (view all comments), comment_add (add comment to paragraph), revisions_list, revisions_accept_all, revisions_reject_all"},
                 "enable": {"type": "boolean", "description": "True to enable track changes, False to disable"},
                 "text": {"type": "string", "description": "Comment text"}, "para_index": {"type": "integer", "description": "Paragraph to attach comment to"},
                 "range_start": {"type": "integer"}, "range_end": {"type": "integer"},
                 "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # --- reference ---
        Tool(name="reference", description="Manage footnotes, endnotes, bookmarks, and field codes in WPS Word. Use when user asks to: add footnotes/endnotes, list footnotes, add/goto/list bookmarks, insert PAGE/DATE/FILENAME fields. Actions: add_footnote/add_endnote/list_footnotes/add_bookmark/goto_bookmark/list_bookmarks/insert_field",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "add_footnote (insert footnote), add_endnote (insert endnote), list_footnotes (view all), add_bookmark (create bookmark), goto_bookmark (jump to), list_bookmarks (view all), insert_field (PAGE/DATE etc.)"},
                 "text": {"type": "string", "description": "Text for footnote/endnote"},
                 "name": {"type": "string", "description": "Bookmark name"},
                 "para_index": {"type": "integer", "description": "Paragraph index for reference anchor"},
                 "field_code": {"type": "string", "description": "Field code: PAGE, DATE, FILENAME, NUMPAGES, etc."},
                 "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # --- docspace ---
        Tool(name="docspace", description="Unified document space: list and manage all open Word and Excel documents. Use when user asks to: see all open documents, switch between docs, save all, or close all. Actions: list_all/activate/close_all/save_all",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "list_all (all open docs), activate (switch to doc_id), close_all (close everything without saving), save_all (save everything)"},
                 "doc_id": {"type": "string", "description": "Document ID like word:1, excel:1"},
             }, "required": ["action"]}),

        # --- template ---
        Tool(name="template", description="Manage document formatting templates. Use when user asks to: extract formatting from a document, save/load templates, list available templates, import/export templates, or compare document against a template. Actions: extract/save/load/list/delete/export/import/compare",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "extract (get rules from doc), save (store template), load (get template), list (all templates), delete, export (to file), import (from file), compare (doc vs template)"},
                 "template_name": {"type": "string", "description": "Template name"},
                 "filepath": {"type": "string", "description": "Path for export/import"},
                 "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # --- transfer ---
        Tool(name="transfer", description="Copy content between Word documents. Use when user asks to: copy paragraphs, tables, or text ranges from one doc to another. Actions: copy_paragraphs/copy_table/copy_range",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "copy_paragraphs (copy paragraph range), copy_table (copy a table), copy_range (copy text range)"},
                 "source_doc_id": {"type": "string", "description": "Source doc ID e.g. word:1"},
                 "target_doc_id": {"type": "string", "description": "Target doc ID e.g. word:2"},
                 "from_start": {"type": "integer", "description": "Start paragraph index for copy_paragraphs"},
                 "from_end": {"type": "integer", "description": "End paragraph index for copy_paragraphs"},
                 "table_index": {"type": "integer", "description": "Table index for copy_table"},
                 "start_pos": {"type": "integer", "description": "Character start position for copy_range"},
                 "end_pos": {"type": "integer", "description": "Character end position for copy_range"},
                 "target_position": {"type": "string", "description": "Where to insert: end, or paragraph index number"},
             }, "required": ["action", "source_doc_id", "target_doc_id"]}),

        # --- migrate ---
        Tool(name="migrate", description="Migrate data between Word and Excel. Use when user asks to: export Word table to Excel, import Excel range into Word as table. Actions: word_table_to_excel/excel_range_to_word_table",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "word_table_to_excel (Word table -> Excel), excel_range_to_word_table (Excel range -> Word table), word_outline_to_ppt (Word outline -> PPT)"},
                 "word_doc_id": {"type": "string", "description": "Word doc ID e.g. word:1"},
                 "excel_doc_id": {"type": "string", "description": "Excel doc ID e.g. excel:1"},
                 "table_index": {"type": "integer", "description": "Word table index"},
                 "range_start": {"type": "string", "description": "Excel range start e.g. A1"},
                 "range_end": {"type": "string", "description": "Excel range end e.g. D10"},
                 "target_cell": {"type": "string", "description": "Target cell for Excel e.g. A1"},
                 "position": {"type": "string", "description": "Target position in Word: end, or paragraph index"},
                 "keep_format": {"type": "boolean", "description": "Preserve formatting"},
             }, "required": ["action"]}),

        # --- compare ---
        Tool(name="compare", description="Compare two Word documents. Use when user asks to: find text differences, find formatting differences, or compare document structure. Actions: text_diff/format_diff",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "text_diff (line-level text diff), format_diff (font/paragraph format comparison)"},
                 "doc_id_a": {"type": "string", "description": "First doc ID e.g. word:1"},
                 "doc_id_b": {"type": "string", "description": "Second doc ID e.g. word:2"},
             }, "required": ["action", "doc_id_a", "doc_id_b"]}),

        # --- ai_format ---
        Tool(name="ai_format", description="AI-powered intelligent formatting for WPS Word. Use when user asks to: analyze document formatting, suggest improvements, apply professional templates, auto-apply formatting from natural language, generate table of contents, add multi-level heading numbers, validate formatting quality, generate/summarize/rewrite/expand/translate content, or run quality supervision to auto-fix layout issues. Actions: analyze/suggest/apply_template/reformat/auto_toc/auto_numbering/validate/generate_content/summarize_document/rewrite_paragraph/expand_section/translate_section/supervise",
             inputSchema={"type": "object", "properties": {
                 "action": {"type": "string", "description": "analyze/suggest/apply_template/reformat/auto_toc/auto_numbering/validate/generate_content/summarize_document/rewrite_paragraph/expand_section/translate_section/supervise (auto-fix layout/cover/table issues)"},
                 "template_name": {"type": "string", "description": "official/thesis/report/resume/custom"},
                 "instructions": {"type": "string", "description": "Natural language formatting instructions for reformat"},
                 "doc_index": {"type": "integer"},
             }, "required": ["action"]}),

        # --- presentation ---
        Tool(name="presentation", description="Create and manipulate WPS Presentations (PPT). Use when user asks to: create/open/save PPT, add slides, set titles/body text, insert images/tables, format text, add speaker notes, or export slides as images. Actions: create/open/list/save/close/slide_count/slide_info/add_slide/delete_slide/set_title/set_body/add_textbox/format_text/insert_image/insert_table/fill_cell/apply_theme/add_notes",
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
        Tool(name="excel", description="Create and manipulate WPS Excel spreadsheets. Use when user asks to: create/open/save Excel workbooks, switch/add/rename sheets, read/write cell values, write tabular data, set formulas (=SUM etc.), format cells (font/color/borders), adjust column widths, merge cells, create charts, sort data, filter, conditional format, freeze panes. Actions: create/open/list/save/close/sheet_list/sheet_activate/sheet_add/sheet_copy/sheet_delete/sheet_move/cell_read/cell_write/range_read/range_write/font_set/interior_set/borders_set/column_width/auto_fit/merge_cells/formula_set/chart_add/chart_set_source/chart_set_title/sort/auto_filter/remove_filter/conditional_format/freeze_panes/get_used_range",
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
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict):
    action = arguments.get("action", "")
    doc_index = arguments.get("doc_index")

    try:
        if name == "document":
            if action == "info":
                result = document.doc_info(doc_index)
            elif action == "list":
                result = document.doc_list()
            elif action == "open":
                result = document.doc_open(arguments["filepath"])
            elif action == "create":
                result = document.doc_create(arguments.get("filepath"))
            elif action == "save":
                result = document.doc_save(doc_index, arguments.get("filepath"))
            elif action == "close":
                result = document.doc_close(doc_index, arguments.get("save_changes", False))
            elif action == "activate":
                result = document.doc_activate(doc_index if doc_index is not None else 1)
            elif action == "export_pdf":
                result = document.doc_export_pdf(doc_index, arguments.get("output_path"))
            elif action == "insert_image":
                result = document.insert_image(arguments["filepath"], arguments.get("width"), arguments.get("height"), arguments.get("position", "end"), doc_index)
            elif action == "doc_properties":
                result = document.doc_properties(doc_index)
            elif action == "set_doc_properties":
                result = document.set_doc_properties(arguments.get("author"), arguments.get("title"), arguments.get("subject"), doc_index)
            else:
                result = {"error": f"Unknown document action: {action}"}

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
            else:
                result = {"error": f"Unknown format action: {action}"}

        elif name == "style":
            if action == "list":
                result = formatting.list_styles(doc_index)
            elif action == "get":
                result = formatting.get_style(arguments["name"], doc_index)
            elif action == "create":
                result = formatting.create_style(doc_index=doc_index, **{k: v for k, v in arguments.items() if k not in ("action", "doc_index")})
            elif action == "modify":
                result = formatting.modify_style(doc_index=doc_index, **{k: v for k, v in arguments.items() if k not in ("action", "doc_index")})
            else:
                result = {"error": f"Unknown style action: {action}"}

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
            else:
                result = {"error": f"Unknown table action: {action}"}

        elif name == "search":
            if action == "find":
                result = search.find_text(arguments["query"], arguments.get("match_case", False), arguments.get("whole_word", False), doc_index)
            elif action == "replace":
                result = search.replace_text(arguments["find_text"], arguments["replace_text"], arguments.get("match_case", False), arguments.get("replace_all", False), doc_index)
            elif action == "find_format":
                result = search.find_format(arguments.get("font_name"), arguments.get("font_size"), arguments.get("bold"), arguments.get("style_name"), doc_index)
            elif action == "goto_heading":
                result = search.goto_heading(arguments.get("text"), arguments.get("level"), doc_index)
            else:
                result = {"error": f"Unknown search action: {action}"}

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
            else:
                result = {"error": f"Unknown layout action: {action}"}

        elif name == "review":
            if action == "track_changes_toggle":
                result = review.track_changes_toggle(arguments.get("enable", False), doc_index)
            elif action == "track_changes_status":
                result = review.track_changes_status(doc_index)
            elif action == "comments_list":
                result = review.comments_list(doc_index)
            elif action == "comment_add":
                result = review.comment_add(arguments["text"], arguments.get("para_index"), arguments.get("range_start"), arguments.get("range_end"), doc_index)
            elif action == "revisions_list":
                result = review.revisions_list(doc_index)
            elif action == "revisions_accept_all":
                result = review.revisions_accept_all(doc_index)
            elif action == "revisions_reject_all":
                result = review.revisions_reject_all(doc_index)
            else:
                result = {"error": f"Unknown review action: {action}"}

        elif name == "reference":
            if action == "add_footnote":
                result = document.add_footnote(arguments.get("para_index"), arguments.get("text", ""), doc_index)
            elif action == "add_endnote":
                result = document.add_endnote(arguments.get("para_index"), arguments.get("text", ""), doc_index)
            elif action == "list_footnotes":
                result = document.list_footnotes(doc_index)
            elif action == "add_bookmark":
                result = document.add_bookmark(arguments["name"], arguments.get("para_index"), doc_index)
            elif action == "goto_bookmark":
                result = document.goto_bookmark(arguments["name"], doc_index)
            elif action == "list_bookmarks":
                result = document.list_bookmarks(doc_index)
            elif action == "insert_field":
                result = document.insert_field(arguments.get("para_index"), arguments.get("field_code", "PAGE"), doc_index)
            else:
                result = {"error": f"Unknown reference action: {action}"}

        elif name == "docspace":
            if action == "list_all":
                result = docspace.list_all()
            elif action == "activate":
                result = docspace.activate(arguments["doc_id"])
            elif action == "close_all":
                result = docspace.close_all()
            elif action == "save_all":
                result = docspace.save_all()
            else:
                result = {"error": f"Unknown docspace action: {action}"}

        elif name == "template":
            result = _handle_template(action, arguments, doc_index)

        elif name == "transfer":
            if action == "copy_paragraphs":
                result = transfer.copy_paragraphs(arguments["source_doc_id"], arguments["from_start"], arguments["from_end"], arguments["target_doc_id"], arguments.get("target_position", "end"))
            elif action == "copy_table":
                result = transfer.copy_table(arguments["source_doc_id"], arguments["table_index"], arguments["target_doc_id"], arguments.get("target_position", "end"))
            elif action == "copy_range":
                result = transfer.copy_range(arguments["source_doc_id"], arguments["start_pos"], arguments["end_pos"], arguments["target_doc_id"], arguments.get("target_position", "end"))
            else:
                result = {"error": f"Unknown transfer action: {action}"}

        elif name == "migrate":
            if action == "word_table_to_excel":
                result = migrate.word_table_to_excel(arguments["word_doc_id"], arguments["table_index"], arguments["excel_doc_id"], arguments.get("target_cell", "A1"), arguments.get("keep_format", False))
            elif action == "excel_range_to_word_table":
                result = migrate.excel_range_to_word_table(arguments["excel_doc_id"], arguments["range_start"], arguments["range_end"], arguments["word_doc_id"], arguments.get("position", "end"), arguments.get("keep_format", False))
            elif action == "word_outline_to_ppt":
                result = migrate.word_outline_to_ppt(arguments["word_doc_id"])
            else:
                result = {"error": f"Unknown migrate action: {action}"}

        elif name == "compare":
            if action == "text_diff":
                result = compare.text_diff(arguments["doc_id_a"], arguments["doc_id_b"])
            elif action == "format_diff":
                result = compare.format_diff(arguments["doc_id_a"], arguments["doc_id_b"])
            else:
                result = {"error": f"Unknown compare action: {action}"}

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
            else:
                result = {"error": f"Unknown ai_format action: {action}"}

        elif name == "excel":
            result = _handle_excel(action, arguments)

        elif name == "presentation":
            result = _handle_presentation(action, arguments)

        else:
            result = {"error": f"Unknown tool: {name}"}

    except Exception as e:
        logger.exception(f"Tool error: {name}/{action}")
        result = {"error": str(e), "tool": name, "action": action}

    return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]


def _ai_analyze(doc_index, template=None):
    try:
        from intelligence.layout_analyzer import analyze
        return analyze(doc_index)
    except Exception as e:
        logger.exception("ai_format.analyze")
        outline_data = content.outline(doc_index)
        return {
            "document": document.doc_info(doc_index),
            "outline": outline_data,
            "error": str(e),
            "note": "LLM analysis requires valid API key in config.yaml"
        }


def _ai_suggest(doc_index):
    try:
        from intelligence.format_suggester import suggest
        return suggest(doc_index)
    except Exception as e:
        logger.exception("ai_format.suggest")
        outline_data = content.outline(doc_index)
        return {
            "outline_count": len(outline_data),
            "headings": outline_data[:20],
            "error": str(e),
            "note": "LLM suggestions require valid API key in config.yaml"
        }


def _ai_apply_template(template_name, doc_index):
    from intelligence.chinese_rules import CHINESE_FORMATTING
    from wps_bridge.app import get_app, get_doc
    from wps_bridge.utils import com_property, com_set, WDALIGNMENT, WDLINESPACING
    tmpl = CHINESE_FORMATTING.get(template_name)
    if not tmpl:
        return {"error": f"Template not found: {template_name}", "available": list(CHINESE_FORMATTING.keys())}
    applied = []
    doc = get_doc(doc_index)
    first_para_done = False
    for level_name, level_rules in tmpl.items():
        if level_name == "page":
            from wps_bridge import layout as lay
            lay.page_setup(doc_index=doc_index, **level_rules)
            applied.append("page_setup")
            continue
        is_cover = level_rules.pop("is_cover", False)
        outline_level = level_rules.get("outline_level", 10)
        para_count = 0
        for i in range(1, doc.Paragraphs.Count + 1):
            try:
                p = doc.Paragraphs.Item(i)
                pl = com_property(p.Format, "OutlineLevel", 10)
                is_body = (level_name == "正文" and pl >= 9 and com_property(p.Range, "Text", "").strip())
                is_heading = (pl == outline_level and 1 <= outline_level <= 9)
                should_apply = is_heading or is_body
                if is_cover and not first_para_done:
                    should_apply = (i == 1)
                    first_para_done = True
                if should_apply:
                    r = p.Range
                    com_set(r.Font, "ColorIndex", 1)
                    if "font_name" in level_rules:
                        com_set(r.Font, "Name", level_rules["font_name"])
                    if "font_name_fallback" in level_rules and not com_set(r.Font, "Name", level_rules["font_name"]):
                        com_set(r.Font, "Name", level_rules["font_name_fallback"])
                    if "font_size" in level_rules:
                        com_set(r.Font, "Size", level_rules["font_size"])
                    if "bold" in level_rules:
                        com_set(r.Font, "Bold", level_rules["bold"])
                    if "alignment" in level_rules:
                        com_set(r.ParagraphFormat, "Alignment", WDALIGNMENT.get(level_rules["alignment"], 3))
                    if "first_line_indent_chars" in level_rules:
                        indent = level_rules["first_line_indent_chars"] * level_rules.get("font_size", 14)
                        com_set(r.ParagraphFormat, "FirstLineIndent", indent)
                    elif "first_line_indent_pt" in level_rules:
                        com_set(r.ParagraphFormat, "FirstLineIndent", level_rules["first_line_indent_pt"])
                    elif "first_line_indent" in level_rules:
                        com_set(r.ParagraphFormat, "FirstLineIndent", level_rules["first_line_indent"])
                    if "line_spacing_rule" in level_rules:
                        com_set(r.ParagraphFormat, "LineSpacingRule", WDLINESPACING.get(level_rules["line_spacing_rule"], 0))
                    if "line_spacing" in level_rules:
                        com_set(r.ParagraphFormat, "LineSpacing", level_rules["line_spacing"])
                    if "space_before" in level_rules:
                        com_set(r.ParagraphFormat, "SpaceBefore", level_rules["space_before"])
                    if "space_after" in level_rules:
                        com_set(r.ParagraphFormat, "SpaceAfter", level_rules["space_after"])
                    para_count += 1
            except Exception:
                continue
        level_rules["is_cover"] = is_cover
        if para_count > 0:
            applied.append(f"{level_name}({para_count}段)")
    return {"template": template_name, "applied_to": applied}


def _ai_reformat(instructions, doc_index):
    if not instructions:
        return {"error": "No instructions provided"}
    from intelligence.layout_analyzer import generate_reformat_actions
    actions = generate_reformat_actions(instructions, doc_index)
    if not actions:
        return {
            "instructions": instructions,
            "note": "LLM analysis required for natural language parsing. API key needed in config.yaml.",
            "manual_hint": "Use specific tools: format/set_font, format/set_paragraph_format, style/create, layout/page_setup, etc."
        }

    def _execute_batch(actions_list):
        executed = []
        failed = []
        for act in actions_list:
            tool_name = act.get("tool", "format")
            act_args = {k: v for k, v in act.items() if k not in ("tool", "reason")}
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
    else:
        result = {"error": f"Unknown presentation action: {action}"}
    return result


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
        result = delete(args["template_name"])
    elif action == "export":
        result = export_template(args["template_name"], args["filepath"])
    elif action == "import":
        result = import_template(args["filepath"])
    elif action == "compare":
        result = compare_with_template(doc_index, args["template_name"])
    else:
        result = {"error": f"Unknown template action: {action}"}
    return result


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options(
            notification_options=NotificationOptions()
        ))


if __name__ == "__main__":
    asyncio.run(main())
