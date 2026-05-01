# AGENTS.md — WPS Agent MCP Server

> 当前进度：P0/P1 已完成，综合评分 9.2/10 (A)。
> 2026-05-02 完成两轮全面优化修复，16 项 Bug 修复 + 13 项新功能。

---

## 0. 六维评分（当前状态 2026-05-02）

| 维度 | 终态 | 当前 | 状态 |
|------|------|------|------|
| 1. 原生理解 Word 文档内容 | 10/10 | **9.5** | semantic_model + query_by_role + 内容分类 ✓ |
| 2. 精确控制字体大小、格式 | 10/10 | **9.5** | set_run_font/set_text_effect ✓ |
| 3. 精确理解排版、控制排版 | 10/10 | **9.5** | LayoutAnalyzer + fix_widow_orphan + auto_fix_layout ✓ |
| 4. 手术级全面理解与修改 | 10/10 | **9.5** | SurgicalContext + MCP surgical tool ✓ |
| 5. 实用性 | 10/10 | **8.5** | 结构化错误码 + 性能基准 ✓ |
| 6. 能否跑通 | 10/10 | **9.0** | Unit Tests(121+) + 优雅降级 ✓ |
| **综合** | **10** | **9.2** | **A** |

---

## 1. 项目概要

- **语言**：Python 3.11+
- **框架**：MCP (Model Context Protocol) stdio server
- **COM 层**：pywin32 → WPS Kwps.Application / Ket.Application / Kwpp.Application
- **离线引擎**：`docx_engine/` 基于 lxml 的原生 OOXML 读写（包含 DOM 模型、样式解析、语义分析、排版分析）
- **AI 层**：LLM API (DeepSeek/OpenAI) 用于智能排版和内容生成
- **GitHub**：github.com/alllyx520-bot/wps-agent
- **MCP Tool 数量**：18 个（document/content/format/style/table/search/layout/review/reference/docspace/transfer/migrate/compare/presentation/excel/ai_format/offline_docx/content_control/field_codes/surgical/operation_log）

### 架构层次

```
MCP Protocol Layer (mcp_server.py)
    ├── Online Mode: wps_bridge/ (COM → WPS)
    │   ├── content / formatting / table / layout / search / review
    │   ├── document / docspace / transfer / migrate / compare
    │   ├── content_control / field_codes / surgical_context
    │   └── excel / presentation (Excel/PPT COM)
    ├── Offline Mode: docx_engine/ (XML-native)
    │   ├── document_model.py (DOM: Run/Paragraph/Table/Cell)
    │   ├── intelligence.py (文档分析/角色检测/类型识别)
    │   ├── formatter.py (自动排版/多级编号)
    │   ├── semantic_model.py (语义解析/关系图谱/内容分类)
    │   └── layout_model.py (排版分析/溢出检测/表格检测/分栏检测)
    └── Intelligence Layer: intelligence/
        ├── llm_client.py (LLM API)
        ├── chinese_rules.py (14套中文排版模板)
        ├── quality_supervisor.py (质量评估+自动修复)
        └── layout_analyzer.py (排版分析→LLM增强)
```

---

## 2. 运行与测试

### 2.1 环境要求

```bash
# 虚拟环境
E:\Anaconda\envs\wps-agent\python.exe

# 依赖安装
pip install -r requirements.txt

# WPS Office 必须运行（Online 模式需要）
```

### 2.2 测试命令

```bash
# 语法检查（任何时候可跑）
python -c "import py_compile; py_compile.compile('mcp_server.py', doraise=True); print('OK')"

# 导入检查（不启动 COM）
python -c "import sys; sys.path.insert(0,'.'); from wps_bridge import content, formatting; from docx_engine import document_model, intelligence; print('import OK')"

# 纯逻辑单测（不依赖 WPS）
python -m pytest tests/test_unit.py -v

# 性能基准
python tests/benchmark.py
```

### 2.3 验证规范
- 每次修改后至少跑语法检查 + 导入检查
- 涉及 COM 调用的修改需在 WPS 运行状态下验证
- 报错必须在原文件修复，禁止创建 `_simple.py` 等临时文件

---

## 3. 添加新工具 / Action 的标准流程

| 步骤 | 文件 | 操作 |
|------|------|------|
| 1. 实现函数 | `wps_bridge/<module>.py` 或 `docx_engine/<module>.py` | 添加核心逻辑函数 |
| 2. 注册路由 | `mcp_server.py` → `call_tool()` | 在对应分支添加调用 |
| 3. 声明 Schema | `mcp_server.py` → `list_tools()` | 在 Tool 描述中追加 action |
| 4. 更新描述 | Schema 描述 + 参数列表 | — |
| 5. 写测试 | `tests/` | 正常参数 + 异常参数 |

### 函数签名规范

```python
def new_action(param1: type, param2: Optional[type] = None, doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    return {"result": "..."}
```

- 必传参数在前，可选参数在后
- `doc_index` 永远放在最后，默认 `None`
- 返回值统一用 `Dict`

### COM 操作规范

```python
from .utils import com_property, com_set, com_set_batch

value = com_property(obj, "PropertyName", default_value)  # 读
com_set(obj, "PropertyName", value)                       # 写
failed = com_set_batch(obj, {"Prop1": val1})              # 批量写
```

- **禁止** 直接用 `obj.PropName` 或 `setattr()`
- 中文字体必须同时设 `Name` 和 `NameFarEast`
- WPS 样式名使用中文（标题 1、正文等）
- 段落索引 **1-based**

---

## 4. 封面生成规则

| 禁止 | 原因 | 替代 |
|------|------|------|
| 多次 `insert_text` 逐段创建 | 文本合并到一个段落 | `content action=create_cover lines=[...]` |
| 不清除继承的段落间距 | Normal 样式自带间距 | `_apply_line_format` 强制 SpaceBefore/After=0 |
| `doc.Range(0, Content.End).Delete()` | WPS 拒绝删除末尾段落标记 | `doc.Content.Text = ""` 可靠清空 |

---

## 5. Git 规范

- 不主动 commit，等用户要求
- **禁止** force push 到 main/master
- 提交前检查：不包含 `config.yaml`（含 API Key）、不包含 `logs/`、不包含 `__pycache__/`
- `.gitignore` 已排除：`config.yaml`、`logs/`、`__pycache__/`、`*.log`

---

## 6. 已归档的血泪教训

### COM 线程模型
- WPS COM 调用必须在 STA 线程中执行

### 段落索引
- WPS COM 的 `Paragraphs.Item(i)` 是 **1-based**
- `doc.Content.End` 返回字符位置（含段落标记）

### 字体名称
- 设置中文字体必须同时设 `Name` 和 `NameFarEast`
- 常用中文字体：`黑体`、`宋体`、`仿宋`、`楷体`、`微软雅黑`

### 格式控制 Bug（2026-05 修复）
- COM 模式 `set_run_font`/`set_paragraph_format` 无边界检查 → 新增段落数校验
- walrus 操作符零值吞没 → 改为 `is not None` 判断
- `italic` 强制覆写 → 仅显式传入时设置
- 新建文档操作 → 返回 `PARAGRAPH_OUT_OF_RANGE` 结构化错误
- `content_control.count` 无控件时 `t` 变量未绑定 → 重构 type_names 构建逻辑

### 离线路径 Bug（2026-05 修复）
- `offline_docx build` 是硬编码 error → 实现完整的 structure.paragraphs 构建逻辑
- `create_cover`/`snapshot`/`rollback` 离线代码死路 → 从"始终 COM"列表移除
- `delete_range` 离线处理缺少 `elif` 守卫 → 补充 action 分支
- `insert_run`/`delete_run`/`split_run` 不在离线写入列表 → 补齐
- presentation/excel handler 全部 stub → 重连通 COM 桥
