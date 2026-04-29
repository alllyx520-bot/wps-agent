# WPS Agent

> An MCP server that turns WPS Office into an AI-powered document assistant. 18 tools, 130+ actions, covering Word/Excel/PPT via COM automation.

## Features

- **Word**: full document CRUD, paragraph/range text operations, outline, styles, tables, search/replace, layout, track changes, comments, footnotes, bookmarks, fields, watermarks, images, document properties
- **Excel**: workbook/sheet CRUD, cell/range read/write, formatting, charts, sort, filter, conditional format, formulas, freeze panes
- **PPT**: presentation CRUD, slide management, text/shape/table/image operations, speaker notes
- **Cross-document**: copy paragraphs/tables/ranges between Word docs, Word↔Excel data migration, Word outline→PPT generation, document text/format diff
- **AI formatting**: analyze document structure, suggest improvements, apply 12 built-in Chinese templates (official/thesis/report/resume/contract/letter/proposal/meeting_minutes/press_release/manual/exam/bid), natural language reformatting, auto TOC, auto heading numbering, error self-healing
- **AI content**: generate/summarize/rewrite/expand/translate document content via LLM
- **Template system**: extract formatting from documents, save/load/export/import templates, compare documents against templates

## Architecture

```
opencode (AI Agent)
    ↕ MCP stdio
mcp_server.py (18 tools)
    ↕ Python import
wps_bridge/          intelligence/
├── app.py           ├── chinese_rules.py (12 presets)
├── document.py      ├── content_generator.py (5 AI write ops)
├── content.py       ├── format_suggester.py
├── formatting.py    ├── layout_analyzer.py
├── table.py         ├── llm_client.py
├── layout.py        └── template_manager.py
├── search.py
├── review.py        COM (pywin32)
├── docspace.py          ↕
├── transfer.py      WPS Office (Windows)
├── migrate.py       ├── Kwps.Application (Word)
├── compare.py       ├── Ket.Application (Excel)
├── excel_app.py     └── Kwpp.Application (PPT)
├── ppt_app.py
└── utils.py
```

## Prerequisites

- **Windows** with WPS Office installed (COM automation only works on Windows)
- **Python 3.11+**
- **Conda** (recommended for environment isolation)

## Installation

```bash
# 1. Clone
git clone https://github.com/YOUR_USERNAME/wps-agent.git
cd wps-agent

# 2. Create conda environment
conda create -n wps-agent python=3.11 -y
conda activate wps-agent

# 3. Install dependencies
pip install -r requirements.txt

# 4. Configure
copy config.yaml.example config.yaml
# Edit config.yaml and set your LLM API key (or use WPS_AGENT_LLM_KEY env var)
```

## Configuration

**LLM API Key** (choose one):

| Method | How |
|--------|-----|
| Environment variable | Set `WPS_AGENT_LLM_KEY=sk-xxx` |
| config.yaml | Add `api_key: "sk-xxx"` in the `llm:` section |
| MCP environment | Pass via MCP client config (see below) |

**config.yaml**:
```yaml
llm:
  provider: "deepseek"
  endpoint: "https://api.deepseek.com/v1"
  model: "deepseek-chat"
  api_key: ""  # or set WPS_AGENT_LLM_KEY env var
```

Supports any OpenAI-compatible endpoint (DeepSeek, Aliyun DashScope, OpenAI, etc.).

## MCP Setup (opencode / Claude Desktop)

Add to `opencode.jsonc` or `claude_desktop_config.json`:

```json
{
  "mcp": {
    "wps-agent": {
      "type": "local",
      "command": ["C:\\path\\to\\python.exe", "C:\\path\\to\\wps-agent\\mcp_server.py"],
      "environment": {
        "WPS_AGENT_LLM_KEY": "your-api-key-here"
      }
    }
  }
}
```

Restart your MCP client. WPS Office must be running.

## Tool Reference

### Word Tools

| Tool | Action | Description |
|------|--------|-------------|
| `document` | info/list/open/create/save/close/activate/export_pdf/insert_image/doc_properties/set_doc_properties | Document lifecycle, metadata, images |
| `content` | full_text/paragraph/paragraphs/selection/range/outline/insert_text/delete_range/replace_range/batch | Read/write document text |
| `format` | get_font/set_font/get_paragraph_format/set_paragraph_format/apply_style/clear_formatting/copy_format/batch/add_watermark/remove_watermark | Font & paragraph formatting |
| `style` | list/get/create/modify | Style management (标题 1, 正文, etc.) |
| `table` | count/info/read/create/delete/set_cell_text/format_cell/set_header/format_borders/merge_cells/auto_fit/set_column_width/alternate_rows/batch_read | Table operations |
| `search` | find/replace/find_format/goto_heading | Search & replace |
| `layout` | page_setup/section_info/add_section_break/columns/header_footer/page_numbers | Page layout, headers, page numbers |
| `review` | track_changes_toggle/track_changes_status/comments_list/comment_add/revisions_list/revisions_accept_all/revisions_reject_all | Track changes & comments |
| `reference` | add_footnote/add_endnote/list_footnotes/add_bookmark/goto_bookmark/list_bookmarks/insert_field | Footnotes, bookmarks, fields |

### Cross-Application Tools

| Tool | Action | Description |
|------|--------|-------------|
| `docspace` | list_all/activate/close_all/save_all | Unified document management across Word/Excel/PPT |
| `transfer` | copy_paragraphs/copy_table/copy_range | Copy content between Word documents |
| `migrate` | word_table_to_excel/excel_range_to_word_table/word_outline_to_ppt | Data migration between applications |
| `compare` | text_diff/format_diff | Document comparison |

### Excel Tool

| Tool | Action | Description |
|------|--------|-------------|
| `excel` | create/open/list/save/close/sheet_list/sheet_activate/sheet_add/sheet_copy/sheet_delete/sheet_move/cell_read/cell_write/range_read/range_write/font_set/interior_set/borders_set/column_width/auto_fit/merge_cells/formula_set/chart_add/chart_set_source/chart_set_title/sort/auto_filter/remove_filter/conditional_format/freeze_panes/get_used_range | Full Excel automation |

### PPT Tool

| Tool | Action | Description |
|------|--------|-------------|
| `presentation` | create/open/list/save/close/slide_count/slide_info/add_slide/delete_slide/set_title/set_body/add_textbox/format_text/insert_image/insert_table/fill_cell/apply_theme/add_notes | Full PowerPoint automation |

### AI Tools

| Tool | Action | Description |
|------|--------|-------------|
| `template` | extract/save/load/list/delete/export/import/compare | Template management (12 built-in presets) |
| `ai_format` | analyze/suggest/apply_template/reformat/auto_toc/auto_numbering/validate/generate_content/summarize_document/rewrite_paragraph/expand_section/translate_section | AI-powered formatting & content generation |

## Usage Examples

### Natural Language Commands (via opencode)

```
"把第一段改成黑体三号加粗居中"
"用学术论文模板格式化当前文档"
"自动生成目录和标题编号"
"在A1到D5写入销售数据并画柱状图"
"根据文档大纲生成一份10页的演示文稿"
"总结全文内容"
"把A文档的第3-5段复制到B文档末尾"
"对比这两个文档的排版差异"
```

### Batch Operations

```json
// format.batch — modify multiple paragraphs in one call
{"action": "batch", "operations": [
  {"type": "set_font", "para_index": 1, "name": "黑体", "size": 16, "bold": true},
  {"type": "set_font", "para_index": 5, "name": "黑体", "size": 16, "bold": true}
]}

// content.batch — read multiple items in one call
{"action": "batch", "items": [
  {"type": "paragraph", "para_index": 1},
  {"type": "outline"}
]}
```

### AI Content Generation

```
"在第3章之后写一段关于技术风险的补充说明"
"把这段内容翻译成英文"
"润色第5段"
"扩展第2章，补充更多细节"
```

## Built-in Templates

| Template | For |
|----------|-----|
| `official` | Government documents (GB/T 9704) |
| `thesis` | Academic papers |
| `report` | Business reports |
| `resume` | Resumes/CVs |
| `contract` | Contracts/agreements |
| `letter` | Official letters |
| `proposal` | Project proposals |
| `meeting_minutes` | Meeting minutes |
| `press_release` | Press releases |
| `manual` | User manuals |
| `exam` | Exam papers |
| `bid` | Bid documents |

## Project Structure

```
wps-agent/
├── mcp_server.py          # MCP server entry point (18 tools)
├── config.yaml.example    # Configuration template
├── requirements.txt       # Python dependencies
├── README.md
├── .gitignore
├── wps_bridge/            # COM automation layer
│   ├── app.py             # Word COM singleton
│   ├── document.py        # Document CRUD + advanced features
│   ├── content.py         # Text read/write
│   ├── formatting.py      # Font & paragraph formatting
│   ├── table.py           # Table operations
│   ├── layout.py          # Page layout
│   ├── search.py          # Find & replace
│   ├── review.py          # Track changes & comments
│   ├── docspace.py        # Unified document space
│   ├── transfer.py        # Cross-document copy
│   ├── migrate.py         # Word↔Excel migration
│   ├── compare.py         # Document diff
│   ├── excel_app.py       # Excel COM automation
│   ├── ppt_app.py         # PPT COM automation
│   └── utils.py           # COM helpers
├── intelligence/          # AI & template layer
│   ├── llm_client.py      # LLM API client
│   ├── chinese_rules.py   # 12 built-in templates
│   ├── template_manager.py # Template extract/save/load
│   ├── content_generator.py # AI content generation
│   ├── format_suggester.py # Format suggestion
│   └── layout_analyzer.py # Document analysis
└── logs/                  # Test & debug files
```

## Technical Notes

- WPS COM ProgIDs: `Kwps.Application` (Word), `Ket.Application` (Excel), `Kwpp.Application` (PPT)
- COM connection persists across MCP calls with automatic reconnection on failure
- Style names use Chinese in WPS (标题 1, 正文) — not English (Heading 1)
- Supports DeepSeek, OpenAI, and any OpenAI-compatible LLM API
- Python COM calls run on a single thread (COM STA requirement)
- Tested on WPS 12.0 / Windows 11 / Python 3.11

## License

MIT
