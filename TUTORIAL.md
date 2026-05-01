# WPS AI Agent 使用教程 v2

> 一个深度集成 WPS Office 的 AI 排版专家，精通 Word 和 Excel，通过 MCP 协议与 AI Agent 实时交互。配备 `document-author` Skill，实现类人 4-Phase 文档工作流（理解→规划→执行→验证）。

---

## 一、环境准备

### 1.1 前置条件

| 组件 | 要求 |
|---|---|
| Windows | Windows 10/11 |
| WPS Office | 12.0+ (已安装并注册 COM) |
| Python | 3.11+ (conda env: `wps-agent`) |

### 1.2 安装

```bash
conda create -n wps-agent python=3.11 -y
pip install -r requirements.txt
```

### 1.3 MCP 配置

在 `opencode.jsonc` 添加：
```json
"wps-agent": {
  "type": "local",
  "command": ["E:\\Anaconda\\envs\\wps-agent\\python.exe", "你的路径\\mcp_server.py"]
}
```

保存后重启 opencode。

### 1.5 部署 opencode 智能化配置（推荐）

本仓库 `opencode_config/` 目录包含让 AI Agent 像人类文档专家一样工作的配置：

```
opencode_config/
├── AGENTS.md              # Agent 行为规则（含 document-author 自动触发）
├── skills/                # 9 个 Skill
│   ├── document-author/   # ★ 类人文档智能化（4-Phase 工作流）
│   ├── docx/              # .docx 离线创建/编辑
│   ├── xlsx/              # .xlsx 电子表格
│   ├── pptx/              # .pptx 演示文稿
│   └── ...                # code-review / debug / agent-creator 等
├── commands/              # 自定义快捷命令
└── agents/                # 自定义 Agent
```

**核心 Skill：`document-author`**

这个 Skill 颠覆了传统的"逐一调用 MCP 工具"模式，让 Agent 像人类一样思考和操作文档：

```
Phase 1: 理解     → 读写文档全文+大纲+格式，构建"文档心智模型"
Phase 2: 规划     → 自然语言输出修改计划+影响分析，再动手
Phase 3: 执行     → 逐步操作，每步记录状态（已改了什么、还剩什么）
Phase 4: 验证     → 重读修改区域，一致性检查，有问题立即修正
```

**部署方法**：将 `opencode_config/` 下所有文件复制到 `~/.config/opencode/`：

```powershell
robocopy opencode_config\ $env:USERPROFILE\.config\opencode\ /E
```

重启 opencode 后，所有 WPS Word 操作将自动触发 4-Phase 工作流。

---

## 二、快速开始

```
在 WPS 中新建一个空白文档
查看当前文档信息
```

---

## 三、Word 排版操作

### 3.1 字体格式

```
把第1段改成黑体三号加粗
把第3段改成楷体小四号
给标题添加阴影效果               ← set_text_effect
给文字添加外框效果
```

**Run 级精确控制：**
```json
{"tool": "format", "action": "set_run_font", "para_index": 3, "run_index": 2, "name": "黑体", "size": 16, "bold": true}
```

**文本特效：**
```json
{"tool": "format", "action": "set_text_effect", "para_index": 1, "effect": "shadow"}
{"tool": "format", "action": "set_text_effect", "para_index": 1, "effect": "glow", "color_rgb": 255}
```
支持: shadow, outline, emboss, engrave, glow, reflection

**中文字号对照表：**
| 字号 | 初号 | 小初 | 一号 | 小二 | 二号 | 三号 | 四号 | 小四 | 五号 | 小五 |
|------|------|------|------|------|------|------|------|------|------|------|
| 磅值 | 42 | 36 | 26 | 18 | 22 | 16 | 14 | 12 | 10.5 | 9 |

### 3.2 段落格式

```
把第2段设置为两端对齐，首行缩进2字符，1.5倍行距
```
段落属性: alignment, first_line_indent, left_indent, right_indent, line_spacing, space_before, space_after, outline_level

### 3.3 样式管理

```
列出所有样式 → 170+ WPS 内置样式
创建名为"我的标题"的新样式，字体微软雅黑14号加粗
把第5段应用"标题 2"样式
```

---

## 四、手术级修改 (surgical) 🔥

这是最精确的修改方式——先捕获上下文，再批量修改，最后验证+回滚。

### 4.1 按索引选择

```
选中第2段和第3段进行手术级修改
```
```json
{"tool": "surgical", "action": "select", "para_indices": [2, 3]}
→ {"session_id": "2097188697808", "context": {...}}
```

### 4.2 队列化修改

```json
{"tool": "surgical", "action": "modify",
 "session_id": "2097188697808",
 "mutations": [
   {"para": 2, "run": 1, "font_name": "黑体", "size": 22, "bold": true},
   {"para": 2, "alignment": "center"},
   {"para": 3, "first_line_indent": 24, "space_after": 12}
 ]}
→ {"mutations_queued": 3}
```

### 4.3 提交验证

```json
{"tool": "surgical", "action": "commit", "session_id": "2097188697808"}
→ {"committed": true, "applied": [...], "verified": true}
```

### 4.4 回滚恢复

```json
{"tool": "surgical", "action": "rollback", "session_id": "2097188697808"}
→ {"rolled_back": true, "restored_paragraphs": 2}
```

---

## 五、语义理解与内容分类

### 5.1 按角色定位 (query_by_role)

```
定位到文档摘要段落
定位到参考文献
定位到封面信息
定位到目录
```
```json
{"tool": "content", "action": "query_by_role", "sr": "abstract", "filepath": "doc.docx"}
→ {"role": "abstract", "matched": 1, "paragraphs": [{"index": 3, "role": "abstract_label", ...}]}
```

支持的角色: abstract, cover, keywords, toc, references, acknowledgements, appendix

### 5.2 内容性质分类

自动识别段落类型：论述型 / 数据型 / 公式型 / 代码型 / 引用型

---

## 六、排版分析与自动修正

### 6.1 孤行修正

```
自动修正文档中的所有孤行
```
```json
{"tool": "layout", "action": "fix_widow_orphan"}
→ {"fixed_paragraphs": 39, "total_paragraphs": 39}
```

### 6.2 排版全量自动修正

```
分析文档排版，自动修正所有缺陷
```
```json
{"tool": "layout", "action": "auto_fix_layout", "filepath": "doc.docx"}
→ {"fixed": 2, "fixes": [...], "issues_found": 3}
```

检测维度：文本溢出、孤行标题、表格跨页断行、分栏不均衡、图片位置

---

## 七、表格操作

```
在文档末尾插入一个4行3列的表格
在表格第1行第1列填入"姓名"
将表格第1行设为标题行（黑体加粗）
合并表格1中第1行第2列到第1行第3列
```

---

## 八、页面布局

```
将页面设置为A4纸，上下边距2.54cm
将当前节设置为两栏
设置页眉为"XX公司年度报告"
添加居中页码
```

---

## 九、查找替换与审阅

```
在文档中搜索"项目"
把所有"旧词"替换为"新词"
开启修订模式
在第3段添加批注："此处需要补充数据"
接受所有修订
```

---

## 十、AI 智能排版

### 10.1 自动格式化

```
分析当前文档的排版结构
用学术论文模板格式化当前文档
给文档标题添加多级编号
生成文档目录
```

### 10.2 一键全流程 (auto_enhance)

```
对文档进行全流程增强
```
工作流：Parse → Semantic Analyze → Format → Numbering → Quality Check → Layout Fix → Save

7 阶段全自动，无需手动干预。

---

## 十一、Excel 操作

```
新建 Excel 工作簿
在A1到D5写入数据
在C6设置公式=SUM(C2:C5)
给A1:D1表头设置黑体加粗灰色背景
创建柱状图
排序/筛选/条件格式
```

---

## 十二、PPT 操作

```
新建 PPT
添加幻灯片
设置第1页标题为"年度总结"
添加演讲者备注
```

### 场景4：类人智能文档修改（document-author）

启用 `document-author` skill 后，Agent 会像人类专家一样操作：

```
1. 打开需要修改的文档
2. "把参考文献格式改成 GB/T 7714 国标"
3. Agent 自动：
   - Phase 1: 读取全文+大纲+格式，发现当前参考文献是字母序排列
   - Phase 2: 输出规划："修改 ref[1]-ref[15] 为 GB/T 7714...
                 → 注意 TOC 页码可能变化"
   - Phase 3: 逐条调整，每步记录进度
   - Phase 4: 重读参考文献区域，检查所有条目格式一致
4. "修改第三章的标题格式和图注编号"
5. Agent 发现文档风格 → 自动匹配现有标题格式，不盲套标准
```

---

## 十三、Offline 模式（无需 WPS）

```json
{"tool": "offline_docx", "action": "analyze", "filepath": "doc.docx"}
{"tool": "offline_docx", "action": "auto_format", "filepath": "doc.docx"}
{"tool": "offline_docx", "action": "build", "structure": {"paragraphs": [...]}}
{"tool": "offline_docx", "action": "apply_template", "template_name": "thesis_cn", "filepath": "doc.docx"}
{"tool": "offline_docx", "action": "replace_text", "old_text": "旧", "new_text": "新", "filepath": "doc.docx"}
```

---

## 十四、全部 MCP Tool 一览

| 工具 | Action 数 | 核心能力 |
|------|----------|---------|
| `document` | 11 | 文档 CRUD + 导出 PDF + 属性 |
| `content` | 28 | 文本读写 + Run级操作 + 语义查询 + 封面 + 快照回滚 |
| `format` | 18 | 字体/段落格式 + 文本特效 + 水印/超链接/制表位 |
| `style` | 4 | 样式 CRUD |
| `table` | 15 | 表格全功能 + 格式化 |
| `search` | 4 | 查找替换 + 格式查找 |
| `layout` | 13 | 页面设置 + 页眉页脚 + 孤行修正 + 自动修正 |
| `review` | 7 | 修订/批注全功能 |
| `reference` | 5 | 脚注/尾注/书签/域代码 |
| `docspace` | 4 | 多文档统一管理 |
| `transfer` | 3 | 跨文档复制 |
| `migrate` | 3 | Word↔Excel↔PPT 数据迁移 |
| `compare` | 2 | 文档文本/格式对比 |
| `presentation` | 18 | PPT 全功能 COM 操作 |
| `excel` | 27 | Excel 全功能 COM 操作 |
| `ai_format` | 10 | AI 分析/排版/内容生成 |
| `offline_docx` | 12 | 离线 build/analyze/format/template |
| `content_control` | 10 | Content Control CRUD |
| `field_codes` | 12 | 域代码全功能 |
| `surgical` | 4 | 手术级上下文->修改->验证->回滚 |
| `operation_log` | 5 | 操作审计日志 |

---

## 十五、常见问题

**Q: Agent 无法连接 WPS？**
A: 确保 WPS Office 已启动且 COM 组件已注册。

**Q: 修改格式没有生效？**
A: COM 操作是同步的，修改立即反映。若使用 Offline 模式需检查文件锁。

**Q: 中文样式名报错？**
A: WPS 内置样式使用中文名（"标题 1"而非"Heading 1"）。"标题 2" → ✅，"Heading 2" → ❌。

**Q: 云端文档能操作吗？**
A: 可以，只要在 WPS 中打开了云端文档，Agent 就能完全操作。

**Q: LLM 分析不工作？**
A: 检查 `config.yaml` 中 API Key 是否正确，网络是否能访问 API 端点。

**Q: 如何切换 LLM 模型？**
A: 修改 `config.yaml` 中的 `endpoint` 和 `model`，支持所有 OpenAI 兼容 API。

---

## 十三、项目结构

```
wps-agent/
├── mcp_server.py              # MCP Server 入口（10个Tool Group）
├── config.yaml                # 配置文件
├── wps_bridge/                # COM 桥接层
│   ├── app.py                 # WPS Word Application
│   ├── document.py            # 文档管理
│   ├── content.py             # 内容读写
│   ├── formatting.py          # 格式/样式
│   ├── table.py               # 表格操作
│   ├── layout.py              # 页面布局
│   ├── search.py              # 查找替换
│   ├── review.py              # 修订批注
│   ├── excel_app.py           # WPS Excel Application
│   └── utils.py               # COM 工具函数
├── intelligence/              # AI 智能层
│   ├── llm_client.py          # LLM API 客户端
│   ├── layout_analyzer.py     # 排版分析器
│   ├── format_suggester.py    # 格式建议器
│   └── chinese_rules.py       # 中文排版规则库
├── opencode_config/              # ★ opencode 智能化配置
│   ├── AGENTS.md                # Agent 行为规则（自动触发 document-author）
│   ├── skills/                  # 9个 Skill（含 document-author 4-Phase 工作流）
│   ├── commands/                # 自定义快捷命令
│   └── agents/                  # 自定义 Agent
└── logs/                      # 日志
```
