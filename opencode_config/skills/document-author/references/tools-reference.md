# WPS Agent MCP Tool Reference — Complete Action Cheat Sheet

> Auto-generated from v2.0 tool suite. 20 tools. 200+ actions.

---

## Core Word Tools

### document — Document management
`info` | `list` | `open` | `create` | `save` | `close` | `activate` | `export_pdf` | `insert_image` | `doc_properties` | `set_doc_properties` | `health_check`

### content — Read/write with Run-level precision
`full_text` (read all) | `paragraph` (read one) | `paragraphs` (read range) | `outline` (headings) | `runs_detail` (individual Run font/size) | `document_structure` (paragraphs+style) | `full_structure` (everything) | `semantic_structure` (20+ roles) | `query_by_role` (by semantic role) | `insert_text` | `insert_paragraph` | `delete_range` | `replace_range` | `create_cover` (with lines array) | `snapshot` | `rollback` | `batch` | `cache_status`

### format — Font and paragraph formatting with Run-level precision
Font: `get_font` | `set_font` | `get_run_font` | `set_run_font`
Paragraph: `get_paragraph_format` | `set_paragraph_format`
Style: `apply_style` | `clear_formatting` | `copy_format`
Special: `set_text_effect` (shadow/glow/outline/emboss/engrave/reflection)
Utility: `add_watermark` | `remove_watermark` | `add_hyperlink` | `set_tab_stops` | `set_bullet_list`
Meta: `batch` | `resolve_format` | `resolve_run_format`

**set_font params**: name, name_far_east, size, bold, italic, underline, color_index, color_rgb, highlight, superscript, subscript, strike_through, spacing, scaling, kerning, caps, small_caps, shadow, outline, emboss, vanish

**set_paragraph_format params**: alignment (left/center/right/justify), first_line_indent (pt), left_indent, right_indent, line_spacing_rule (single/1.5lines/double/exactly/multiple), line_spacing, space_before, space_after, outline_level (1-9=heading, 10=body), widow_control, keep_with_next, keep_together

### style — Style management
`list` | `get` | `create` | `modify`

### table — Table operations
`count` | `info` | `read` | `create` | `delete` | `set_cell_text` | `format_cell` | `set_header` | `format_borders` | `merge_cells` | `auto_fit` | `set_column_width` | `alternate_rows` | `set_cell_shading` | `batch_read` | `table_dimensions`

### search — Find & replace
`find` | `replace` | `find_format` | `goto_heading`

### layout — Page layout + v2.0 auto-fix
Page: `page_setup` (width/height/margins/orientation) | `section_info` | `page_dimensions` | `page_border`
Sections: `add_section_break` | `columns`
Headers: `header_footer` | `page_numbers`
Objects: `page_break` | `image_wrap` | `line_numbers`
**v2.0:** `fix_widow_orphan` (COM WidowControl on all) | `auto_fix_layout` (offline analysis + fix)

### review — Track changes & comments
`track_changes_toggle` | `track_changes_status` | `comments_list` | `comment_add` | `revisions_list` | `revisions_accept_all` | `revisions_reject_all`

### reference — Footnotes, bookmarks, fields
`add_footnote` | `add_endnote` | `list_footnotes` | `add_bookmark` | `goto_bookmark` | `list_bookmarks` | `insert_field`

---

## v2.0 New Tools

### surgical — Context-aware surgical editing 🔥
| Action | Params | Description |
|--------|--------|-------------|
| `select` | `para_indices: [int]` or `sr: str + filepath` | Capture target paragraphs + neighbors, return session_id |
| `modify` | `session_id, mutations: [{para, run?, font_name?, size?, bold?, alignment?, ...}]` | Queue mutations |
| `commit` | `session_id` | Execute all + verify + return results |
| `rollback` | `session_id` | Restore to pre-capture state |

**Mutation keys**: para (required), run (for run-level), font_name, size, bold, italic, underline, color_index, color_rgb, highlight, superscript, subscript, strike_through, caps, small_caps, alignment, first_line_indent, left_indent, right_indent, line_spacing_rule, line_spacing, space_before, space_after, outline_level, text

---

## Advanced Tools

### content_control — Content Control operations
`count` | `list_controls` | `info` | `add` (RICH_TEXT/TEXT/CHECKBOX/DROPDOWN_LIST/DATE_PICKER) | `set_text` | `set_checkbox` | `select_dropdown` | `delete` | `set_tag` | `find_by_tag`

### field_codes — Advanced field code operations
`insert_field` | `insert_quote` | `insert_doc_property` | `insert_seq` | `insert_style_ref` | `insert_ref` | `insert_if` | `list_fields` | `update_fields` | `unlink_field` | `find_field_by_code`

### docspace — Multi-document management
`list_all` | `activate` | `close_all` | `save_all`

### transfer — Cross-document copy
`copy_paragraphs` | `copy_table` | `copy_range`

### migrate — Cross-app migration
`word_table_to_excel` | `excel_range_to_word_table` | `word_outline_to_ppt`

### compare — Document comparison
`text_diff` | `format_diff`

---

## Offline Tools

### offline_docx — Offline document processing (no WPS needed)
**Build:** `build` (from JSON structure) | `build_cover`
**Analyze:** `analyze` (type+stats+quality) | `validate` | `full_structure` | `semantic_structure` | `cross_references` | `detect_semantic_roles` | `analyze_layout`
**Transform:** `auto_format` | `apply_template` | `add_numbering` | `replace_text`
**IO:** `get_text` | `get_statistics` | `read_model` | `write_model`

**Template names**: thesis_cn, report_official, resume_professional

---

## Excel & PPT Tools

### excel — Excel COM operations
**Workbook:** `create` | `open` | `save` | `close` | `list` | `get_used_range`
**Sheets:** `sheet_list` | `sheet_add` | `sheet_activate` | `sheet_copy` | `sheet_delete` | `sheet_move`
**Data:** `cell_read` | `cell_write` | `range_read` | `range_write`
**Format:** `font_set` | `interior_set` | `borders_set` | `column_width` | `auto_fit` | `merge_cells`
**Logic:** `formula_set` | `chart_add` | `chart_set_source` | `chart_set_title`
**Analysis:** `sort` | `auto_filter` | `remove_filter` | `conditional_format` | `freeze_panes`

### presentation — PPT COM operations
`create` | `open` | `save` | `close` | `list` | `slide_count` | `slide_info` | `add_slide` | `delete_slide` | `set_title` | `set_body` | `add_textbox` | `format_text` | `insert_image` | `insert_table` | `fill_cell` | `apply_theme` | `add_notes`

---

## AI Tools

### ai_format — AI-powered intelligent formatting
**Analysis:** `analyze` | `suggest` | `validate` | `detect_type` | `detect_role` | `batch_detect_roles`
**Templates:** `apply_template` | `reformat`
**Auto:** `auto_toc` | `auto_numbering` | `auto_enhance` (7-stage pipeline)
**Content:** `generate_content` | `summarize_document` | `rewrite_paragraph` | `expand_section` | `translate_section`

---

## operation_log — Audit log
`summary` | `recent` | `errors` | `replay_last` | `clear` | `dump`
