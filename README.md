# WPS Agent

> 把 WPS Office 变成 AI 驱动的文档助手。基于 MCP 协议，通过 COM + Offline 双模式控制 Word/Excel/PPT，支持语义理解、手术级修改、AI 智能排版。

**综合评分：9.2/10** | 18 个 MCP Tool | 200+ Action

---

## 功能概览

- **Word**：文档增删改查、段落/Run/选区/大纲操作、样式管理、表格全功能、搜索替换、页面布局、修订/批注、脚注/书签/域代码、水印、图片、文档属性
- **Excel**：工作簿/工作表 CRUD、单元格/区域读写、格式化、图表、排序、筛选、条件格式、公式、冻结窗格
- **PPT**：演示文稿 CRUD、幻灯片管理、文本框/表格/图片操作、演讲者备注
- **语义理解**：20+ 语义角色识别（封面/摘要/关键词/目录/章/节/正文/引用/致谢/附录）、文档类型分类（12 种）、内容性质分类（论述/数据/公式/代码/引用）
- **手术级修改**：上下文捕获 → 批量修改 → 验证 → 回滚（`surgical` tool）
- **排版分析**：页面几何、文本溢出、表格跨页断行、分栏不均衡、孤行检测 → 自动修正
- **跨文档**：文档间复制段落/表格/文本、Word↔Excel 数据迁移、Word 大纲→PPT 生成、文档文本/格式对比
- **AI 排版**：分析文档结构、应用 14 套中文预设模板、自然语言排版、自动目录/标题编号、质量校验
- **AI 内容**：通过 LLM 生成/总结/改写/扩写/翻译文档内容
- **双模式**：Online (WPS COM 实时操作) + Offline (XML 原生读写，无需 WPS)

---

## 架构

```
opencode (AI Agent)              ← 你说人话，它调工具
    ↕ MCP stdio
mcp_server.py (18 个 Tool)       ← MCP 协议层
    ↕ Python import
├── wps_bridge/                  ← COM 自动化层 (Online)
│   ├── app.py                   # Word COM 单例 + 断连重连
│   ├── content.py               # 文本读写 + Run级操作
│   ├── formatting.py            # 字体/段落格式 + 文本特效
│   ├── table.py                 # 表格全功能
│   ├── layout.py                # 页面设置/页眉页脚
│   ├── search.py / review.py    # 查找替换/审阅批注
│   ├── document.py / docspace.py / transfer.py / migrate.py / compare.py
│   ├── content_control.py       # Content Control
│   ├── field_codes.py           # 域代码
│   ├── surgical_context.py      # 手术级上下文
│   ├── excel_app.py / ppt_app.py  # Excel/PPT COM 桥
│   └── utils.py                 # COM 工具函数
├── docx_engine/                 ← 离线引擎 (Offline)
│   ├── document_model.py        # DOM: Run/Paragraph/Table/Cell
│   ├── intelligence.py          # 文档分析/角色检测/类型识别
│   ├── formatter.py             # 自动排版/多级编号
│   ├── semantic_model.py        # 语义解析/关系图谱/内容分类
│   ├── layout_model.py          # 排版分析/溢出检测
│   ├── xml_parser.py / serializer.py  # OOXML 解析/序列化
│   └── errors.py                # 结构化错误码
├── offline/                     ← 离线构建器
│   └── docx_builder.py          # OfflineDocxBuilder (build/format/analyze)
├── intelligence/                ← AI 智能层
│   ├── llm_client.py            # LLM API 客户端
│   ├── chinese_rules.py         # 14 套中文排版模板
│   ├── quality_supervisor.py    # 质量评估 + 自动修复
│   ├── content_generator.py     # AI 内容生成
│   ├── format_intelligence.py   # 一键全流程 (auto_enhance)
│   └── layout_analyzer.py       # 排版分析 → LLM 增强
    ↕ COM (pywin32)
WPS Office (Windows)
├── Kwps.Application (Word)
├── Ket.Application (Excel)
└── Kwpp.Application (PPT)
```

---

## 安装

```bash
# 1. 克隆
git clone https://github.com/alllyx520-bot/wps-agent.git
cd wps-agent

# 2. 创建虚拟环境
conda create -n wps-agent python=3.11 -y
conda activate wps-agent

# 3. 安装依赖
pip install -r requirements.txt
```

## MCP 客户端配置

在 `opencode.jsonc` 的 `mcp` 段添加：

```json
"wps-agent": {
  "type": "local",
  "command": [
    "E:\\Anaconda\\envs\\wps-agent\\python.exe",
    "E:\\AAAprojects\\自由测试\\wps-agent\\mcp_server.py"
  ],
  "environment": {
    "WPS_AGENT_LLM_KEY": "你的API-Key"
  }
}
```

## 工具速查表

### Word 核心工具

| 工具 | 主要 Action |
|------|------------|
| `document` | info/list/open/create/save/close/activate/export_pdf/doc_properties |
| `content` | full_text/paragraph/outline/runs_detail/document_structure/full_structure/semantic_structure/query_by_role/insert_text/create_cover/replace_range/snapshot/rollback/batch |
| `format` | get_font/set_font/get_run_font/set_run_font/get_paragraph_format/set_paragraph_format/apply_style/clear_formatting/copy_format/batch/add_watermark/remove_watermark/add_hyperlink/set_tab_stops/set_bullet_list/set_text_effect |
| `style` | list/get/create/modify |
| `table` | count/info/read/create/set_cell_text/format_cell/set_header/format_borders/merge_cells/alternate_rows/batch_read |
| `search` | find/replace/goto_heading |
| `layout` | page_setup/section_info/columns/header_footer/page_numbers/page_break/image_wrap/line_numbers/fix_widow_orphan/auto_fix_layout |
| `review` | track_changes_toggle/comments_list/comment_add/revisions_accept_all |
| `reference` | add_footnote/add_endnote/add_bookmark/list_bookmarks/insert_field |

### 手术级工具

| 工具 | 主要 Action |
|------|------------|
| `surgical` | select(按段落索引/语义角色) → modify(队列化修改) → commit(批量执行+验证) / rollback(恢复) |

### 高级工具

| 工具 | 主要 Action |
|------|------------|
| `content_control` | count/list_controls/info/add/set_text/set_checkbox/select_dropdown/delete |
| `field_codes` | insert_field/insert_quote/insert_doc_property/insert_seq/insert_style_ref/insert_ref/insert_if/list_fields/unlink_field |

### 离线工具

| 工具 | 主要 Action |
|------|------------|
| `offline_docx` | build/build_cover/validate/analyze/auto_format/apply_template/replace_text/get_text/get_statistics/full_structure/semantic_structure |

### 跨应用工具

| 工具 | 主要 Action |
|------|------------|
| `docspace` | list_all/activate/close_all/save_all |
| `transfer` | copy_paragraphs/copy_table/copy_range |
| `migrate` | word_table_to_excel/excel_range_to_word_table/word_outline_to_ppt |
| `compare` | text_diff/format_diff |

### Excel / PPT 工具

| 工具 | 主要 Action |
|------|------------|
| `excel` | create/open/cell_read/cell_write/range_read/range_write/formula_set/chart_add/sort/auto_filter/conditional_format/freeze_panes |
| `presentation` | create/add_slide/set_title/set_body/insert_image/insert_table/add_notes/apply_theme |

### AI 工具

| 工具 | 主要 Action |
|------|------------|
| `ai_format` | analyze/suggest/apply_template/reformat/auto_toc/auto_numbering/validate/generate_content/summarize_document/auto_enhance |

---

## 内置模板

| 模板名 | 适用场景 |
|--------|---------|
| `official` | 党政公文 (GB/T 9704) — 方正小标宋标题、三号仿宋正文、三级标题体系 |
| `thesis` | 学术论文 — 三级标题（章/节/条）、宋体小四正文、1.5 倍行距 |
| `report` | 商业报告 — 微软雅黑、1.3 倍行距、现代简洁风格 |
| `resume` | 简历 — 黑体姓名、章节标题、紧凑排版 |
| `contract` | 合同/协议 — 黑体标题、条款层次分明 |
| `letter` | 公函/商务信函 — 仿宋正文、标准书信格式 |
| `proposal` | 项目建议书 — 封面+多级标题、正式排版 |
| `meeting_minutes` | 会议纪要 — 简洁清晰、议题明确 |
| `press_release` | 新闻稿 — 标题/副标题/来源/正文、新闻规范 |
| `manual` | 用户手册 — 三级标题（章/节/步骤）、紧凑排版 |
| `exam` | 试卷 — 大题标题加粗、正文清晰 |
| `bid` | 标书 — 正式严谨、层次分明 |
| `notice` | 通知/通告 — 仿宋正文、标准公文格式 |
| `work_report` | 工作总结 — 层次清晰、汇报风格 |

---

## 使用示例

```
"把第一段改成黑体三号加粗居中"
"用学术论文模板格式化当前文档"
"定位到摘要段落"  ← query_by_role
"对选中的段落使用手术级修改：黑体22号"  ← surgical
"自动修正所有孤行"  ← fix_widow_orphan
"分析文档排版，自动修正缺陷"  ← auto_fix_layout
"给标题添加阴影效果"  ← set_text_effect
"把A文档的第3-5段复制到B文档末尾"
"根据文档大纲生成PPT"
```

---

## License

MIT
