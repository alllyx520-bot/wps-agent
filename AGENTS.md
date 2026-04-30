# AGENTS.md — WPS Agent MCP Server

> 本文件定义 wps-agent 项目的 AI Agent 操作规范。所有规则适用于本项目代码库的操作。

---

## 1. 项目概要

- **语言**：Python 3.11+
- **框架**：MCP (Model Context Protocol) stdio server
- **COM 层**：pywin32 → WPS Kwps.Application / Ket.Application / Kwpp.Application
- **AI 层**：LLM API (DeepSeek/OpenAI) 用于智能排版和内容生成
- **GitHub**：github.com/alllyx520-bot/wps-agent

---

## 2. 运行与测试

### 2.1 启动 MCP Server

```bash
# 先启动 WPS Office（必须是运行状态）
# 然后从 opencode 或其他 MCP 客户端连接
```

### 2.2 手动测试

```bash
# 语法检查
python -c "import py_compile; py_compile.compile('mcp_server.py', doraise=True); print('OK')"

# 导入检查（不启动 COM）
python -c "from wps_bridge import content, formatting; print('import OK')"

# 完整功能测试（需要 WPS 运行中）
python logs/test_e2e.py
```

### 2.3 验证规范
- 每次修改后至少跑语法检查和导入检查
- 涉及 COM 调用的修改需在 WPS 运行状态下验证
- 报错必须在原文件修复，禁止创建 `_simple.py` 等临时文件

---

## 3. 添加新工具 / Action 的标准流程

### 3.1 4 步清单

| 步骤 | 文件 | 操作 |
|------|------|------|
| 1. 实现函数 | `wps_bridge/<module>.py` | 添加核心逻辑函数 |
| 2. 注册路由 | `mcp_server.py` → `call_tool()` | 在对应 `elif action == "xxx":` 分支添加调用 |
| 3. 声明 Schema | `mcp_server.py` → `list_tools()` | 在对应 Tool 的 `inputSchema.properties` 中添加参数 |
| 4. 更新描述 | `mcp_server.py` → Tool `description` | 在 Actions 列表中追加新 action 名 |

### 3.2 函数签名规范

```python
def new_action(param1: type, param2: Optional[type] = None, doc_index: Optional[int] = None) -> Dict:
    """简短说明。"""
    doc = get_doc(doc_index)
    # logic here
    return {"result": "..."}
```

- 必传参数在前，可选参数在后
- `doc_index` 永远放在最后一个参数，默认 `None`
- 返回值统一用 `Dict`

### 3.3 COM 操作规范

```python
from .utils import com_property, com_set, com_set_batch, WDALIGNMENT, WDLINESPACING

# 读属性（安全）
value = com_property(obj, "PropertyName", default_value)

# 写单个属性
com_set(obj, "PropertyName", value)

# 批量写
failed = com_set_batch(obj, {"Prop1": val1, "Prop2": val2})
```

- **禁止** 直接用 `obj.PropName` 或 `setattr()`，COM 异常会导致 Python 崩溃
- 中文属性名使用 `NameFarEast`（而非 `Name`）设置中文字体
- WPS 样式名使用中文（标题 1、正文等），非英文

---

## 4. 封面生成规则（血泪教训）

> `content.py:266 create_cover()` 是唯一的封面创建函数，以下规则必须遵守：

| 禁止 | 原因 | 替代 |
|------|------|------|
| 多次 `insert_text` 逐段创建 | 文本合并到一个段落，格式设置失效 | `content action=create_cover lines=[...]` |
| 用 `delete_range` 清空文档时只传 `start_pos` | 不传 `end_pos`（或传 0）范围无效，段落标记残留 | 用 `create_cover` 的 `clear_existing: true` |
| 依赖 `format batch` 设置封面格式 | 参数格式容易出错 | `create_cover` 内置逐段格式化 |
| 不清除继承的段落间距 | WPS Normal 样式自带段前/段后间距，未显式设置时文本间距失控 | `_apply_line_format` 强制 `SpaceBefore=0, SpaceAfter=0`（line 指定值时覆盖） |
| `doc.Range(0, Content.End).Delete()` 清空 | Word/WPS 拒绝删除末尾段落标记，抛出异常后 fallback 可能残留内容导致多页 | `doc.Content.Text = ""` 可靠清空

### 标准封面调用

```json
{
  "action": "create_cover",
  "clear_existing": true,
  "lines": [
    {"text": "项目标题", "font_name": "黑体", "font_size": 26, "bold": true, "alignment": "center", "space_before": 120, "space_after": 24},
    {"text": "副标题", "font_name": "宋体", "font_size": 16, "alignment": "center", "space_after": 6},
    {"text": "2026年4月", "font_name": "宋体", "font_size": 14, "alignment": "center", "space_after": 6}
  ]
}
```

- `lines` 数组每项支持的字段：`text / font_name / font_size / bold / italic / alignment / space_before / space_after / line_spacing_rule / line_spacing / first_line_indent / left_indent / right_indent / outline_level / underline / strike_through`

---

## 5. Quality Supervisor 规范

`intelligence/quality_supervisor.py` 在每次 `reformat` 或 `generate_content` 后自动运行。

评估维度：
- 段落数量合理性
- 是否所有内容挤在一个段落
- 封面质量（标题字号/居中/间距）
- 表格宽度和表头格式
- 内容顺序（正文在前，表格在后）

修改 `_check_*` 函数时：
- 必须同时更新 `evaluate()` 中的计分逻辑
- 修复阈值写在函数内部，不要硬编码在调用处

---

## 6. Git 规范

- 不主动 commit，等用户要求
- commit message 说明"为什么"而非"改了什么"
- **禁止** force push 到 main/master
- 提交前检查：不包含 `config.yaml`（含 API Key）、不包含 `logs/`、不包含 `__pycache__/`
- `.gitignore` 已排除：`config.yaml`、`logs/`、`__pycache__/`、`*.log`

---

## 7. 中文排版模板体系

`intelligence/chinese_rules.py` 中 `CHINESE_FORMATTING` 字典定义了 14 套模板。

新增模板时：
- 模板 key 使用英文标识符（如 `report`）
- 风格名称使用中文（如 `一级标题`、`正文`）
- 必须包含 `page` 子项定义纸张和页边距
- `is_cover: True` 仅用于封面专属样式
- `font_size` 单位为 pt（point），间距单位为 pt

---

## 8. 已归档的血泪教训

### 8.1 COM 线程模型
- WPS COM 调用必须在 STA 线程中执行
- `mcp_server.py` 中所有 COM 操作默认在单线程运行，不要引入多线程

### 8.2 MCP 返回值
- 所有 tool 返回值必须是 `json.dumps(result, ensure_ascii=False)` 可序列化的 `Dict`
- 不能返回 COM 对象、指针、或不可序列化的 Python 对象

### 8.3 段落索引
- WPS COM 的 `Paragraphs.Item(i)` 是 **1-based**，不是 0-based
- `doc.Content.End` 返回字符位置（含段落标记），删除/插入时注意 -1 偏移

### 8.4 字体名称
- 设置中文字体必须同时设 `Name` 和 `NameFarEast`
- 常用中文字体：`黑体`、`宋体`、`仿宋`、`楷体`、`微软雅黑`

---

## 9. opencode_config：Skills + AGENTS.md 协同体系

> `opencode_config/` 目录存放 opencode AI agent 的行为配置，通过 Skill 与 AGENTS.md 的协同实现类人智能化文档操作。

### 9.1 架构分层

```
用户说"把参考文献格式改成国标"
        │
        ▼
┌─ AGENTS.md（opencode 客户端配置）─────────┐
│ 检测到 WPS Word 操作 → 自动加载            │
│ document-author skill                      │
└───────────┬──────────────────────────────┘
            ▼
┌─ document-author Skill ──────────────────┐
│ Phase 1: 理解 → batch 读全文+大纲+格式     │
│ Phase 2: 规划 → 输出修改计划+影响分析       │
│ Phase 3: 执行 → 逐步操作，记录状态          │
│ Phase 4: 验证 → 重读+一致性检查+自动修正    │
└───────────┬──────────────────────────────┘
            ▼
     WPS MCP 工具 (content/format/table/...)
```

### 9.2 核心能力

| 能力 | 说明 |
|------|------|
| **文档风格发现** | 读 20% 内容后自动推断文档自身格式规律，不盲套标准模板 |
| **语义角色标注** | 自动识别段落类型（封面/标题/正文/参考文献），用语义引用而非数字索引 |
| **一致性守护** | 每次修改后自动对比同类元素格式，不一致立即修正 |
| **影响预判** | 操作前自动分析牵影响（目录/页码/交叉引用） |
| **意图澄清** | 模糊指令不瞎猜，先分析候选方案再确认 |
| **分层打磨** | Pass 1 内容 → Pass 2 格式 → Pass 3 细节 → Pass 4 视觉 |

### 9.3 目录结构

```
opencode_config/
├── AGENTS.md              # Agent 配置（自动触发 document-author）
├── skills/                # 9 个 Skill
│   ├── document-author/   # ★ 类人文档智能化（4-Phase 工作流）
│   │   ├── SKILL.md       #     4-Phase 强制工作流 + 一致性守护 + 分层打磨
│   │   └── references/
│   │       └── conventions.md  # GB/T 9704 / 学术论文 / 实验报告格式参考
│   ├── docx/              # .docx 离线创建/编辑 Skill
│   ├── xlsx/              # .xlsx 电子表格 Skill
│   ├── pptx/              # .pptx 演示文稿 Skill
│   └── ...                # code-review / debug 等其它 Skill
├── commands/              # 自定义快捷命令
└── agents/                # 自定义 Agent
```

### 9.4 部署

```powershell
cd wps-agent
robocopy opencode_config\ %USERPROFILE%\.config\opencode\ /E
```

重启 opencode 后生效。WPS Word 操作将自动走 4-Phase 工作流。

### 9.5 修改 Skill 时注意

- `document-author/SKILL.md` 不含硬编码格式规则，仅含思维框架
- 格式常识在 `references/conventions.md`，作为参考而非强制规则
- `AGENTS.md` 中的触发规则位于 `§2 document-author 智能化操作`
- 修改任一 Skill 后需告知用户重新部署（robocopy 覆盖）
