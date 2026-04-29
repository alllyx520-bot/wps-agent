# WPS AI Agent 全能进化路线图

> 目标：从"排版助手"进化为超越人类的全能 Office AI 专家

---

## 一、总体架构升级

```
                          ┌───────────────────────────────┐
                          │       AI Agent (opencode)      │
                          │   自然语言 → MCP Tool 调用     │
                          └─────────────┬─────────────────┘
                                        │ MCP stdio
                          ┌─────────────┴─────────────────┐
                          │      WPS MCP Server (Python)   │
                          │                                │
                          │  ┌──────────────────────────┐  │
                          │  │    跨文档控制器            │  │
                          │  │  DocSpace: 统一管理所有    │  │
                          │  │  打开的 Word/Excel/PPT 文档 │  │
                          │  └──────────────────────────┘  │
                          │                                │
              ┌───────────┼───────────┬───────────────────┤
              ▼           ▼           ▼                   ▼
        Word COM    Excel COM    PPT COM           Intelligence
        Kwps.App    Ket.App     Kwpp.App          LLM Client
              │           │           │                   │
              ▼           ▼           ▼                   ▼
         ┌─────────────────────────────────────────────────┐
         │              WPS Office (Windows)               │
         │   实时 COM 操作 → 用户可见修改 → 云端同步       │
         └─────────────────────────────────────────────────┘
```

**核心新增：DocSpace 跨文档控制器**
- 不再按应用类型隔离文档，统一用 `doc_id` (格式: `word:1`, `excel:1`, `ppt:1`) 索引
- 支持跨文档操作：复制/迁移/对比
- 统一文档状态管理，COM 断连自动重连

---

## 二、Phase 3 — 多文档互操作

### 2.1 DocSpace 统一文档空间

**目标**：用统一 ID 管理所有打开的 Word/Excel/PPT 文档，消除"必须先激活才能操作"的限制。

**新增模块**：`wps_bridge/docspace.py`

```python
# 统一文档引用格式
doc_id = "word:2"   # 第2个Word文档
doc_id = "excel:1"  # 第1个Excel工作簿
doc_id = "ppt:1"    # 第1个PPT演示文稿

# 所有现有 tool 的 doc_index 参数改为接受统一 doc_id
# 例如:
document.info("word:2")       # 第2个Word文档的信息
excel.cell_read("B5", "excel:2")  # 第2个Excel的B5单元格
```

**新增 MCP Tool: `docspace`**

| action | 参数 | 说明 |
|---|---|---|
| `list_all` | — | 列出所有应用的所有打开文档 |
| `activate` | `doc_id` | 激活指定文档窗口 |
| `close_all` | — | 关闭所有文档 |
| `save_all` | — | 保存所有文档 |

### 2.2 跨文档内容操作

**新增 MCP Tool: `transfer`**

| action | 参数 | 说明 |
|---|---|---|
| `copy_paragraphs` | `source_doc_id`, `from_start`, `from_end`, `target_doc_id`, `target_position` | 从源文档复制段落范围到目标文档 |
| `copy_table` | `source_doc_id`, `table_index`, `target_doc_id`, `target_position` | 复制表格到另一个文档 |
| `copy_range` | `source_doc_id`, `start_pos`, `end_pos`, `target_doc_id`, `target_position` | 复制指定文本范围 |

**示例场景：**
> "把A文档的第3-5段复制到B文档末尾"  
> "把合同模板.docx的第2个表格复制到当前文档"  
> "把销售数据.xlsx的A1:D10区域作为表格插入到报告.docx的第5段之后"

### 2.3 跨应用数据迁移

**新增 MCP Tool: `migrate`**

| action | 参数 | 说明 |
|---|---|---|
| `word_table_to_excel` | `word_doc_id`, `table_index`, `excel_doc_id`, `target_cell`, `keep_format` | Word 表格 → Excel 区域 |
| `excel_range_to_word_table` | `excel_doc_id`, `range_start`, `range_end`, `word_doc_id`, `position`, `keep_format` | Excel 区域 → Word 表格 |
| `word_outline_to_ppt` | `word_doc_id`, `ppt_doc_id` | Word 大纲 → PPT 幻灯片 |

**示例场景：**
> "把月报.docx里的销售数据表导出到 sales.xlsx"  
> "把数据.xlsx 的 A1:F20 作为格式化表格插入到报告.docx"  
> "根据当前文档大纲生成 PPT 演示文稿"

### 2.4 文档对比

**新增 MCP Tool: `compare`**

| action | 参数 | 说明 |
|---|---|---|
| `text_diff` | `doc_id_a`, `doc_id_b` | 逐段对比两个文档的文本差异 |
| `format_diff` | `doc_id_a`, `doc_id_b` | 对比两个文档的格式差异 |
| `merge_changes` | `source_doc_id`, `target_doc_id`, `accept` | 将源文档修改合并到目标文档 |

---

## 三、Phase 4 — PowerPoint 集成

### 3.1 PPT COM 桥接层

**新增模块**：`wps_bridge/ppt_app.py`

```
Kwpp.Application
├── Presentations (.Add / .Open / .Item)
├── Presentation
│   ├── Slides (.Add / .Item / .Count)
│   ├── SlideShowSettings
│   └── PageSetup
├── Slide
│   ├── Shapes (.AddTextbox / .AddPicture / .AddTable)
│   ├── Layout
│   ├── SlideIndex
│   └── NotesPage
├── Shape
│   ├── TextFrame.TextRange (Font / ParagraphFormat)
│   ├── Width / Height / Left / Top
│   ├── Fill (背景填充)
│   ├── Line (边框线)
│   └── AnimationSettings
```

### 3.2 MCP Tool: `presentation`

| action | 参数 | 说明 |
|---|---|---|
| `create` | — | 新建 PPT |
| `open` | `filepath` | 打开 PPT |
| `save` | `filepath` | 保存 |
| `close` | — | 关闭 |
| `slide_count` | — | 获取幻灯片数量 |
| `slide_info` | `slide_index` | 获取某页详情 |
| `add_slide` | `layout_index` | 添加幻灯片 |
| `delete_slide` | `slide_index` | 删除幻灯片 |
| `set_title` | `slide_index`, `text` | 设置标题 |
| `set_body` | `slide_index`, `text` | 设置正文 |
| `add_textbox` | `slide_index`, `text`, `left`, `top`, `width`, `height` | 添加文本框 |
| `format_text` | `slide_index`, `shape_index`, `font_name`, `font_size`, `bold`, `color` | 格式化形状文本 |
| `insert_image` | `slide_index`, `image_path`, `left`, `top`, `width`, `height` | 插入图片 |
| `insert_table` | `slide_index`, `rows`, `cols` | 插入表格 |
| `fill_cell` | `slide_index`, `table_index`, `row`, `col`, `text` | 填充表格 |
| `apply_theme` | `theme_name` | 应用主题模板 |
| `add_notes` | `slide_index`, `text` | 添加演讲者备注 |
| `export_images` | `output_dir` | 每页导出为 PNG |
| `word_outline_to_slides` | `word_doc_id` | 从 Word 大纲生成 PPT |

### 3.3 AI PPT 生成

**`ai_format` 扩展 action: `generate_ppt`**

```
用户: "根据当前文档生成一份8页的演示文稿，每章一页，包含关键要点"

工作流:
1. 读取 Word 文档大纲
2. LLM 分析每章核心内容 → 生成幻灯片文本
3. 逐页创建幻灯片，填充标题+要点
4. 应用主题模板
```

---

## 四、Phase 5 — Word 高级功能扩展

### 4.1 新增 Bridge 函数

**`wps_bridge/document.py` 扩展：**

| 函数 | COM API | 说明 |
|---|---|---|
| `insert_image(filepath, width, height, position)` | `InlineShapes.AddPicture` | 插入图片，支持缩放 |
| `add_footnote(para_index, text)` | `Footnotes.Add` | 插入脚注 |
| `add_endnote(para_index, text)` | `Endnotes.Add` | 插入尾注 |
| `list_footnotes()` | `Footnotes` | 列出所有脚注 |
| `add_bookmark(name, para_index)` | `Bookmarks.Add` | 添加书签 |
| `goto_bookmark(name)` | `Bookmarks.Item(name).Select` | 跳转书签 |
| `list_bookmarks()` | `Bookmarks` | 列出所有书签 |
| `insert_field(para_index, field_code)` | `Fields.Add` | 插入域代码（PAGE/DATE/FILENAME等） |
| `add_watermark(text, font_size, color, layout)` | `Shapes.AddTextEffect` | 添加文字水印 |
| `remove_watermark()` | `Sections(1).Headers(1).Shapes` | 删除水印 |
| `add_caption(label, title, para_index)` | `CaptionsLabels.Add` | 添加题注 |
| `insert_cross_reference(ref_type, ref_text)` | `InsertCrossReference` | 交叉引用 |
| `doc_properties()` | `BuiltInDocumentProperties` | 读取文档属性（作者/标题/页数等） |
| `set_doc_properties(author, title, subject)` | `BuiltInDocumentProperties` | 设置文档属性 |
| `mail_merge(data_source, output_path)` | `MailMerge` | 邮件合并 |

### 4.2 MCP Tool 扩展

**`format` Tool 新增 action：**
- `insert_image`: 插入图片
- `add_watermark` / `remove_watermark`: 水印管理

**新 Tool: `reference`**
| action | 说明 |
|---|---|
| `add_footnote` | 插入脚注 |
| `add_endnote` | 插入尾注 |
| `list_footnotes` | 列出脚注 |
| `add_bookmark` | 添加书签 |
| `goto_bookmark` | 跳转书签 |
| `add_caption` | 添加题注（图表/表格/公式） |
| `insert_cross_ref` | 交叉引用 |
| `insert_field` | 插入域代码 |

**新 Tool: `mailmerge`**
| action | 说明 |
|---|---|
| `set_datasource` | 设置数据源（Excel/CSV） |
| `insert_field` | 插入合并域 |
| `execute` | 执行合并，输出到新文档 |

---

## 五、Phase 6 — Excel 高级功能扩展

### 5.1 新增 Bridge 函数

**`wps_bridge/excel_app.py` 扩展：**

| 函数 | COM API | 说明 |
|---|---|---|
| `sort_range(start, end, key_col, order)` | `Range.Sort` | 区域排序 |
| `auto_filter(start, end, field, criteria)` | `Range.AutoFilter` | 自动筛选 |
| `remove_filter()` | `AutoFilterMode = False` | 清除筛选 |
| `add_conditional_format(start, end, rule_type, formula, interior_color)` | `FormatConditions.Add` | 条件格式 |
| `list_conditional_formats(sheet_name)` | `FormatConditions` | 列出条件格式 |
| `add_data_validation(cell_ref, validation_type, formula)` | `Validation.Add` | 数据验证 |
| `add_named_range(name, cell_ref)` | `Names.Add` | 命名区域 |
| `goto_named_range(name)` | `Range(name).Select` | 跳转命名区域 |
| `list_named_ranges()` | `Names` | 列出所有命名区域 |
| `pivot_table_create(source_range, dest_cell, row_fields, data_fields)` | `PivotCaches.Create` | 创建数据透视表 |
| `chart_set_source(chart_index, range_start, range_end)` | `Chart.SetSourceData` | 设置图表数据源 |
| `chart_set_title(chart_index, title)` | `Chart.HasTitle / ChartTitle.Text` | 设置图表标题 |
| `chart_set_style(chart_index, style_id)` | `Chart.ChartStyle` | 图表样式 |
| `sheet_copy(name, before_after)` | `Worksheets(name).Copy` | 复制工作表 |
| `sheet_delete(name)` | `Worksheets(name).Delete` | 删除工作表 |
| `sheet_move(name, before_after)` | `Worksheets(name).Move` | 移动工作表 |
| `export_csv(filepath, sheet_name)` | `SaveAs(filepath, xlCSV)` | 导出 CSV |
| `import_csv(filepath, sheet_name)` | `QueryTables.Add` | 导入 CSV |
| `get_used_range(sheet_name)` | `UsedRange.Address` | 获取已用区域地址 |
| `freeze_panes(cell_ref)` | `FreezePanes = True` | 冻结窗格 |

### 5.2 MCP Tool 扩展

**`excel` Tool 新增 action：**
- `sort`: 排序
- `auto_filter` / `remove_filter`: 筛选
- `conditional_format`: 条件格式
- `data_validation`: 数据验证
- `named_range`: 命名区域管理
- `pivot_table`: 数据透视表
- `chart_set_source` / `chart_set_title`: 图表完善
- `sheet_copy` / `sheet_delete` / `sheet_move`: 工作表管理
- `export_csv` / `import_csv`: CSV 导入导出
- `freeze_panes`: 冻结窗格
- `get_used_range`: 获取已用范围

### 5.3 AI Excel 分析

**`ai_format` 扩展 action: `analyze_data`**

```
用户: "分析销售数据表，找出季度销售趋势并生成柱状图"

工作流:
1. Excel COM 读取数据区域
2. LLM 分析数据结构，理解列含义
3. LLM 生成分析结论 + 图表类型建议
4. COM 创建图表并绑定数据源
5. 返回分析报告
```

---

## 六、Phase 7 — 模板智能提取与预设系统

### 6.1 模板提取

**新 MCP Tool: `template`**

| action | 参数 | 说明 |
|---|---|---|
| `extract` | `doc_id` | 从指定文档提取所有格式规则 |
| `save` | `template_name`, `rules` | 保存为预设模板 |
| `load` | `template_name` | 加载预设模板 |
| `list` | — | 列出所有自定义预设 |
| `delete` | `template_name` | 删除预设 |
| `export` | `template_name`, `filepath` | 导出为 JSON |
| `import` | `filepath` | 从 JSON 导入 |
| `compare` | `doc_id`, `template_name` | 对比文档与预设的差异 |

**`template.extract` 工作流：**

```
1. 采样文档中每种 OutlineLevel 的段落
2. 提取字体规则 (name, size, bold, italic, color)
3. 提取段落规则 (alignment, indent, line_spacing, space)
4. 提取样式规则 (各样式名对应的完整格式)
5. 提取页面规则 (paper, margins, orientation)
6. 提取表格规则 (header font, border style, alternating colors)
7. 提取编号规则 (numbering scheme)
8. LLM 分析 → 生成带中文描述的完整规则 JSON
9. 保存到 intelligence/templates/<name>.json
```

**提取的规则 JSON 示例：**
```json
{
  "name": "XX公司内部报告模板",
  "extracted_from": "报告模板.docx",
  "extracted_at": "2026-04-29",
  "doc_type": "report",
  "description": "XX公司内部报告格式：一级标题黑体二号加粗居中，正文宋体小四两端对齐...",
  "rules": {
    "page": { "paper": "A4", "top_margin": 72, "bottom_margin": 72, "left_margin": 90, "right_margin": 90 },
    "styles": {
      "标题 1": { "font_name": "黑体", "font_size": 22, "bold": true, "alignment": "center", "space_before": 17, "space_after": 16.5 },
      "标题 2": { "font_name": "黑体", "font_size": 16, "bold": true, "alignment": "left" },
      "正文": { "font_name": "宋体", "font_size": 12, "first_line_indent": 28, "line_spacing_rule": "multiple", "line_spacing": 1.5 }
    },
    "outline_levels": {
      "1": { "font_name": "黑体", "font_size": 22, "bold": true, "alignment": "center" },
      "2": { "font_name": "黑体", "font_size": 16, "bold": true },
      "3": { "font_name": "黑体", "font_size": 14, "bold": true }
    },
    "table_defaults": { "header_font": "黑体", "header_size": 10.5, "header_bold": true, "border_style": 1, "alternate_colors": ["FFFFFF", "F2F2F2"] },
    "numbering": { "level_1": "一、", "level_2": "(一)", "level_3": "1." }
  }
}
```

### 6.2 扩展预设模板库

`intelligence/chinese_rules.py` 从 4 套扩展到 10+ 套：

| 模板名 | 适用场景 |
|---|---|
| `official` | 党政公文 (GB/T 9704) |
| `thesis` | 学术论文 |
| `report` | 商业报告 |
| `resume` | 简历 |
| `contract` | 合同/协议 |
| `bid` | 标书/投标文件 |
| `exam` | 试卷/考题 |
| `press_release` | 新闻稿 |
| `meeting_minutes` | 会议纪要 |
| `manual` | 用户手册/操作指南 |
| `letter` | 公函/商务信函 |
| `proposal` | 项目建议书/立项报告 |

### 6.3 样式迁移

**`template` Tool 新增 action: `transfer_style`**

```
用户: "把A文档的排版风格应用到B文档"

工作流:
1. template.extract(A) → 获取规则 JSON
2. 对 B 文档执行 apply_template(规则)
3. 报告迁移结果（哪些段落匹配/哪些未匹配）
```

---

## 七、Phase 8 — AI 深度增强

### 7.1 内容生成

**`ai_format` 扩展 action:**

| action | 说明 |
|---|---|
| `generate_content` | 在指定位置生成内容（段落/章节/摘要） |
| `expand_section` | 扩写某个章节 |
| `summarize_document` | 全文摘要 |
| `translate_section` | 翻译指定段落 |
| `rewrite_paragraph` | 润色/改写段落 |

**工作流示例：**
```
用户: "在第3章之后写一段关于技术风险的补充说明"

工作流:
1. content.outline → 获取文档上下文
2. content.paragraph(第3章) → 获取前后文
3. LLM: 基于上下文生成补充内容
4. content.insert_text(生成的内容, after=第3章最后一段)
```

### 7.2 智能图表

**`ai_format` 扩展 action: `smart_chart`**

```
用户: "用销售数据画一个柱状图，按部门汇总"

工作流:
1. excel.range_read → 获取原始数据
2. LLM 分析: 理解列含义，选择聚合方式
3. excel.formula_set → 添加汇总公式
4. excel.chart_add + chart_set_source → 创建图表
5. 格式化图表标题/图例/颜色
```

### 7.3 Few-shot Tool 调用优化

**改进 `intelligence/llm_client.py` 的 `parse_natural_language_instructions`：**

```python
# 当前：无示例，LLM 可能生成参数名不匹配
# 优化后：注入完整 tool schema + few-shot 示例

FEW_SHOT_EXAMPLES = [
    {
        "instruction": "把第一段改成黑体三号加粗居中",
        "actions": [
            {"tool": "format", "action": "set_font", "para_index": 1, "name": "黑体", "size": 16, "bold": True},
            {"tool": "format", "action": "set_paragraph_format", "para_index": 1, "alignment": "center"}
        ]
    },
    {
        "instruction": "在A1到D5写入销售数据",
        "actions": [
            {"tool": "excel", "action": "range_write", "start": "A1", "data": [["姓名","销售额"],["张三",120000],["李四",95000]]}
        ]
    },
    # ... 10+ 示例
]
```

### 7.4 错误自愈

当 `_ai_reformat` 执行失败时：
1. 记录失败原因
2. LLM 分析失败原因
3. LLM 生成修正后的 Action
4. 重试（最多 2 次）
5. 仍然失败 → 返回详细诊断信息

---

## 八、Phase 9 — 体验与工程质量

### 8.1 安全加固

- [ ] API Key 从 `config.yaml` 移至环境变量 `WPS_AGENT_LLM_KEY`
- [ ] `config.yaml` 仅保留非敏感配置
- [ ] 添加 `.gitignore` 排除含密钥的配置文件

### 8.2 鲁棒性

- [ ] COM 断连自动重连：`get_app()` 增加心跳检测，WPS 崩溃后自动 `Dispatch` 重连
- [ ] 分级错误码：
  - `E_CONNECTION`: WPS 未运行
  - `E_DOCUMENT`: 文档未找到/已关闭
  - `E_PARAGRAPH`: 段落索引越界
  - `E_PERMISSION`: 权限不足
- [ ] 操作日志：记录每次 MCP 调用的 tool/action/参数/耗时/结果

### 8.3 性能优化

- [ ] `outline()` 优化：仅遍历有 Heading 样式的段落（跳过正文）
- [ ] `paragraphs()` 支持批量获取（一次 COM 调用返回多个段落）
- [ ] 格式读取缓存：连续 get_font 调用复用上次结果（通过段落 hash 判断脏数据）
- [ ] Excel `range_write`：支持 >26 列（`chr(ord('A') + col - 1)` → `openpyxl.utils.get_column_letter` 逻辑）

### 8.4 依赖管理

- [ ] 添加 `requirements.txt`
- [ ] 添加 `pyproject.toml`

### 8.5 批量操作

**新增 action: `batch` 在各自 tool 中**

```
# 一次调用修改 10 个段落的字体
format.batch = [
  {"action": "set_font", "para_index": 1, "name": "黑体", "size": 16},
  {"action": "set_font", "para_index": 5, "name": "黑体", "size": 16},
  {"action": "set_font", "para_index": 11, "name": "黑体", "size": 16},
]

# 一次调用读取 3 个表格
table.batch_read = [1, 2, 3]
```

---

## 九、完整 MCP Tool 全景图（终态）

| Tool Group | Actions | 层级 |
|---|---|---|
| `docspace` | list_all, activate, close_all, save_all | 跨文档统一管理 |
| `document` | info, list, open, create, save, close, activate, export_pdf, insert_image | Word 文档 |
| `content` | full_text, paragraph, paragraphs, selection, range, outline, insert_text, delete_range, replace_range | Word 内容 |
| `format` | get_font, set_font, get_paragraph_format, set_paragraph_format, apply_style, clear_formatting, copy_format, insert_image, add_watermark, remove_watermark | Word 格式 |
| `style` | list, get, create, modify, delete, apply_to_all_like | Word 样式 |
| `table` | count, info, read, create, delete, set_cell_text, format_cell, set_header, format_borders, merge_cells, auto_fit, set_column_width, alternate_rows | Word 表格 |
| `search` | find, replace, find_format, goto_heading | Word 查找 |
| `layout` | page_setup, section_info, add_section_break, columns, header_footer, page_numbers | Word 布局 |
| `review` | track_changes_toggle, track_changes_status, comments_list, comment_add, revisions_list, revisions_accept_all, revisions_reject_all | Word 审阅 |
| `reference` | add_footnote, add_endnote, list_footnotes, add_bookmark, goto_bookmark, add_caption, insert_cross_ref, insert_field | Word 引用 |
| `mailmerge` | set_datasource, insert_field, execute | Word 邮件合并 |
| `transfer` | copy_paragraphs, copy_table, copy_range | 跨文档复制 |
| `migrate` | word_table_to_excel, excel_range_to_word_table, word_outline_to_ppt | 跨应用迁移 |
| `compare` | text_diff, format_diff, merge_changes | 文档对比 |
| `template` | extract, save, load, list, delete, export, import, compare, transfer_style | 模板智能管理 |
| `excel` | create, open, list, save, close, sheet_list, sheet_add, sheet_activate, sheet_copy, sheet_delete, sheet_move, cell_read, cell_write, range_read, range_write, font_set, interior_set, borders_set, column_width, auto_fit, merge_cells, formula_set, chart_add, chart_set_source, chart_set_title, chart_set_style, sort, auto_filter, remove_filter, conditional_format, data_validation, named_range, pivot_table, export_csv, import_csv, freeze_panes, get_used_range | Excel 全功能 |
| `presentation` | create, open, save, close, slide_count, slide_info, add_slide, delete_slide, set_title, set_body, add_textbox, format_text, insert_image, insert_table, fill_cell, apply_theme, add_notes, export_images, word_outline_to_slides | PPT 全功能 |
| `ai_format` | analyze, suggest, apply_template, reformat, auto_toc, auto_numbering, validate, generate_content, expand_section, summarize_document, translate_section, rewrite_paragraph, smart_chart, analyze_data, generate_ppt | AI 排版与分析 |

---

## 十、实施优先级

| 优先级 | Phase | 内容 | 预估工时 |
|---|---|---|---|
| 🔴 P0 | 3 | DocSpace + 多文档互操作 | 2天 |
| 🔴 P0 | 3 | 跨应用迁移 (Word↔Excel) | 1天 |
| 🔴 P0 | 8 | AI few-shot 优化 + 错误自愈 | 1天 |
| 🔴 P0 | 9 | 安全加固 (API Key) + 鲁棒性 (重连) | 0.5天 |
| 🟡 P1 | 4 | PPT 基础支持 | 2天 |
| 🟡 P1 | 5 | Word 高级功能 (图片/脚注/书签/域) | 1.5天 |
| 🟡 P1 | 7 | 模板提取 + 预设库扩展 | 1.5天 |
| 🟡 P1 | 7 | AI 内容生成 (摘要/扩写/翻译) | 1天 |
| 🟢 P2 | 6 | Excel 高级功能 (排序/筛选/透视表) | 1.5天 |
| 🟢 P2 | 6 | Excel AI 分析 | 1天 |
| 🟢 P2 | 4 | AI PPT 生成 (Word→PPT) | 1天 |
| 🟢 P2 | 7 | 样式迁移 + 智能图表 | 1天 |
| ⚪ P3 | 9 | 批量操作 + 性能优化 | 1天 |
| ⚪ P3 | 9 | 依赖管理 + 项目工程化 | 0.5天 |

---

## 十一、兼容性注意事项

1. **COM API 版本差异**：WPS 12.0 的 COM 对象模型高度兼容 Office VBA，但部分高级 API（如 PivotTable、MailMerge）可能存在细微差异，需要实测验证
2. **字体名称**：WPS 字体名使用中文（内部存储为 NameFarEast），`Font.Name` 在某些场景可能返回英文名
3. **PPT Shape API**：WPS PPT 的形状操作可能与 Office PowerPoint 有差异，需要特别验证 `Shapes.AddTextbox` 等接口
4. **大文档性能**：段落数 >1000 的文档，`outline()` 遍历需要优化（改为只遍历有样式的段落）
5. **并发安全**：MCP stdio server 是单线程的，不需要加锁，但 COM STA 约束确保所有 COM 调用在同一线程
