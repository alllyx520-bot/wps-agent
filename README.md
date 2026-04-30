# WPS Agent

> 把 WPS Office 变成 AI 驱动的文档助手。基于 MCP 协议，通过 COM 自动化控制 Word/Excel/PPT，支持 LLM 智能排版与内容生成。

## 功能概览

- **Word**：文档增删改查、段落/选区/大纲操作、样式管理、表格、搜索替换、页面布局、修订/批注、脚注/书签/域代码、水印、图片、文档属性
- **Excel**：工作簿/工作表 CRUD、单元格/区域读写、格式化、图表、排序、筛选、条件格式、公式、冻结窗格
- **PPT**：演示文稿 CRUD、幻灯片管理、文本框/表格/图片操作、演讲者备注
- **跨文档**：文档间复制段落/表格/文本、Word↔Excel 数据迁移、Word 大纲→PPT 生成、文档文本/格式对比
- **AI 排版**：分析文档结构、提改进建议、应用 12 套中文预设模板（公文/论文/报告/简历/合同/公函/项目建议书/会议纪要/新闻稿/手册/试卷/标书）、自然语言排版的、自动目录/标题编号、错误自愈
- **AI 内容**：通过 LLM 生成/总结/改写/扩写/翻译文档内容
- **模板系统**：提取文档格式、保存/加载/导出/导入模板、文档与模板对比

## 架构

```
opencode (AI Agent)          ← 你说人话，它调工具
    ↕ MCP stdio
mcp_server.py (18 个工具)    ← MCP 协议层
    ↕ Python import
wps_bridge/                  ← COM 自动化层
├── app.py                   # Word COM 单例
├── document.py / content.py / formatting.py / table.py
├── layout.py / search.py / review.py
├── docspace.py / transfer.py / migrate.py / compare.py
├── excel_app.py / ppt_app.py / utils.py
intelligence/                ← AI 智能层
├── chinese_rules.py         # 12 套预设模板
├── content_generator.py     # AI 内容生成
├── llm_client.py            # LLM API 客户端
├── template_manager.py      # 模板提取管理
    └── layout_analyzer.py       # 文档分析
    ↕ COM (pywin32)
WPS Office (Windows)
└── opencode_config/          ← Agent 智能行为层
    ├── AGENTS.md              # 自动触发规则：WPS Word 操作时加载 document-author
    └── skills/document-author/ # 4-Phase 类人工作流（理解→规划→执行→验证）
├── Kwps.Application (Word)
├── Ket.Application (Excel)
└── Kwpp.Application (PPT)
```

## 环境要求

- **Windows** + WPS Office（COM 自动化仅支持 Windows）
- **Python 3.11+**
- **Conda**（推荐）

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

# 4. 配置 LLM API Key（三选一）
# 方式一：环境变量
set WPS_AGENT_LLM_KEY=sk-xxx
# 方式二：编辑 config.yaml 的 llm.api_key 字段
# 方式三：在 MCP 客户端配置中注入（见下方）
```

## MCP 客户端配置

在 `opencode.jsonc` 或 `claude_desktop_config.json` 中添加：

```json
{
  "mcp": {
    "wps-agent": {
      "type": "local",
      "command": ["python路径", "wps-agent安装路径\\mcp_server.py"],
      "environment": {
        "WPS_AGENT_LLM_KEY": "你的API-Key"
      }
    }
  }
}
```

重启 MCP 客户端。确保 WPS Office 已启动。

## 工具速查表

### Word 工具

| 工具 | 常用操作 |
|------|---------|
| `document` | info / list / open / create / save / close / export_pdf / insert_image / doc_properties / set_doc_properties |
| `content` | full_text / paragraph / paragraphs / outline / insert_text / batch |
| `format` | get_font / set_font / get_paragraph_format / set_paragraph_format / apply_style / copy_format / batch / add_watermark / remove_watermark |
| `style` | list / get / create / modify |
| `table` | count / info / read / create / set_cell_text / format_cell / set_header / format_borders / merge_cells / batch_read |
| `search` | find / replace / find_format / goto_heading |
| `layout` | page_setup / section_info / add_section_break / columns / header_footer / page_numbers |
| `review` | track_changes_toggle / comments_list / comment_add / revisions_accept_all |
| `reference` | add_footnote / add_endnote / add_bookmark / goto_bookmark / list_bookmarks / insert_field |

### 跨应用工具

| 工具 | 常用操作 |
|------|---------|
| `docspace` | list_all（查看所有打开的文档） / activate / close_all / save_all |
| `transfer` | copy_paragraphs / copy_table / copy_range（文档间复制） |
| `migrate` | word_table_to_excel / excel_range_to_word_table / word_outline_to_ppt |
| `compare` | text_diff / format_diff（文档对比） |

### Excel / PPT 工具

| 工具 | 常用操作 |
|------|---------|
| `excel` | create / open / cell_read / cell_write / range_write / formula_set / chart_add / sort / auto_filter / conditional_format / freeze_panes |
| `presentation` | create / add_slide / set_title / set_body / insert_image / insert_table / add_notes |

### AI 工具

| 工具 | 常用操作 |
|------|---------|
| `template` | extract（提取当前文档格式）/ list（12套预设）/ save / load / compare |
| `ai_format` | analyze / apply_template / reformat / auto_toc / auto_numbering / generate_content / summarize_document / rewrite_paragraph / translate_section |

## 使用示例

```
"把第一段改成黑体三号加粗居中"
"用学术论文模板格式化当前文档"
"自动生成目录"
"把A文档的第3-5段复制到B文档末尾"
"根据文档大纲生成PPT"
"总结全文"
"把这段翻译成英文"
```

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

## 智能化文档操作：Skills + AGENTS.md 协同

`opencode_config/` 目录存放 opencode AI agent 的行为配置，通过 Skill 与 AGENTS.md 的协同实现类人智能化文档操作：

```
用户说"把参考文献格式改成国标"
        │
        ▼
┌─ AGENTS.md §10.1.2 ───────────────────┐
│ 检测到 WPS Word 操作 → 自动加载          │
│ document-author skill                   │
└───────────┬────────────────────────────┘
            ▼
┌─ document-author Skill ─────────────────┐
│ Phase 1: 理解 → batch 读全文+大纲+格式    │
│ Phase 2: 规划 → 输出修改计划+影响分析      │
│ Phase 3: 执行 → 逐步操作，记录状态         │
│ Phase 4: 验证 → 重读+一致性检查+自动修正   │
└───────────┬────────────────────────────┘
            ▼
     WPS MCP 工具 (content/format/table/...)
```

**核心能力：**

| 能力 | 说明 |
|------|------|
| **文档风格发现** | 读 20% 内容后自动推断文档自身的格式规律，不盲套标准模板 |
| **语义角色标注** | 自动识别段落类型（封面/标题/正文/参考文献...），用语义引用而非数字索引 |
| **一致性守护** | 每次修改后自动对比同类元素格式，不一致立即修正 |
| **影响预判** | 操作前自动分析牵影响（目录/页码/交叉引用） |
| **意图澄清** | 模糊指令不瞎猜，先分析候选方案再确认 |
| **分层打磨** | Pass 1 内容正确 → Pass 2 格式统一 → Pass 3 细节到位 → Pass 4 视觉润色 |

**部署方式**：将 `opencode_config/` 下的文件复制到 `~/.config/opencode/` 对应位置即可启用。

## 项目结构

```
wps-agent/
├── mcp_server.py          # MCP 服务入口（18 个工具）
├── config.yaml            # 配置文件
├── requirements.txt       # Python 依赖
├── README.md
├── .gitignore
├── wps_bridge/            # COM 自动化桥接层
│   ├── app.py             # Word COM 单例（含断连重连）
│   ├── document.py        # 文档 CRUD + 图片/书签/水印等
│   ├── content.py         # 文本读写 + 批量操作
│   ├── formatting.py      # 字体/段落格式 + 格式刷 + 批量
│   ├── table.py           # 表格全功能
│   ├── layout.py          # 页面设置/页眉页脚/页码
│   ├── search.py          # 查找替换
│   ├── review.py          # 修订/批注
│   ├── docspace.py        # 统一文档空间
│   ├── transfer.py        # 跨文档复制
│   ├── migrate.py         # Word↔Excel 迁移 + Word→PPT
│   ├── compare.py         # 文档对比
│   ├── excel_app.py       # Excel 全功能 COM 桥接
│   ├── ppt_app.py         # PPT COM 桥接
│   └── utils.py           # COM 辅助函数
└── intelligence/          # AI 智能层
    ├── llm_client.py      # LLM API 客户端（支持 DeepSeek/OpenAI）
    ├── chinese_rules.py   # 12 套中文排版预设
    ├── template_manager.py # 模板提取/保存/加载/对比
    ├── content_generator.py # AI 生成/总结/改写/扩写/翻译
    ├── format_suggester.py # 格式建议
    └── layout_analyzer.py # 文档分析 + 自然语言解析
└── opencode_config/       # opencode AI 配置（skills / commands / agents）
    ├── AGENTS.md           # Agent 行为规则（含 document-author 4-Phase 工作流）
    ├── skills/             # 9 个 skill（含 document-author 文档智能化）
    ├── commands/           # 4 个自定义命令
    └── agents/             # 2 个自定义 agent
```

## 技术说明

- WPS COM ProgID：`Kwps.Application`（Word）、`Ket.Application`（Excel）、`Kwpp.Application`（PPT）
- COM 连接跨 MCP 调用保持，断连自动重连
- WPS 中样式名使用中文（标题 1、正文），非英文
- 支持 DeepSeek、OpenAI 及任何兼容 OpenAI 接口的 LLM API
- Python COM 调用在单线程运行（COM STA 约束）
- 实测环境：WPS 12.0 / Windows 11 / Python 3.11

## License

MIT
