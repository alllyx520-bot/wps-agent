# WPS Agent Rules — 追加到你的主 AGENTS.md

> **部署方式**：将本文文件内容**追加**到 `~/.config/opencode/AGENTS.md` 末尾，不要覆盖原文件。
> 原主 AGENTS.md 的 Anti-Fluff、Vibe Protection、Env Enforcement 等 13 节规则不受影响。
> 命令行追加：`type opencode_config\wps-agent-rules.md >> %USERPROFILE%\.config\opencode\AGENTS.md`

---

## WPS Agent MCP Tool 路由矩阵

WPS Office 操作必须优先使用 `wps-agent` MCP（COM 实时操作），离线 skill 仅作不可用时的后备。

### Word 操作（20+ tools）

| 需求关键词 | 优先 MCP Tool |
|-----------|--------------|
| 文档/创建/打开/保存/关闭/PDF | `document` |
| 内容/段落/大纲/读取/写入 | `content` |
| 字体/大小/加粗/颜色/格式/段落 | `format` |
| 样式/标题1/正文 | `style` |
| 表格/行/列/单元格/合并 | `table` |
| 搜索/查找/替换 | `search` |
| 页面/边距/页眉页脚/页码/分节 | `layout` |
| 修订/批注/审阅 | `review` |
| 脚注/尾注/书签/域代码 | `reference` |
| Content Control/下拉框/日期 | `content_control` |
| 域代码/Date/Seq/StyleRef | `field_codes` |
| 多文档/切换/全部关闭 | `docspace` |
| 跨文档复制/迁移/对比 | `transfer` / `migrate` / `compare` |

### 智能操作（v2.0 新增）

| 需求关键词 | 优先 MCP Tool | 说明 |
|-----------|--------------|------|
| 手术/精确/批量修改/上下文 | `surgical` | select→modify→commit→rollback |
| 定位/摘要/参考文献/封面/查找角色 | `content.query_by_role` | 按语义角色精确定位段落 |
| 自动修正/孤行/排版缺陷 | `layout.fix_widow_orphan` / `layout.auto_fix_layout` | 排版自动检测+修正 |
| 文字特效/阴影/发光/外框 | `format.set_text_effect` | 6种文本特效 |
| 离线批量/构建/分析 | `offline_docx` | 无需WPS的离线文档操作 |

### 后备离线 Skill（WPS MCP 不可用时）

| 需求 | Skill | 条件 |
|------|-------|------|
| Word/docx 文档 | `docx` | 仅当 wps-agent document/content/format 全部不可用 |
| PPT/pptx 演示文稿 | `pptx` | 仅当 wps-agent presentation 不可用 |
| Excel/xlsx 表格 | `xlsx` | 仅当 wps-agent excel 不可用 |

### 工具选择决策树

```
用户说"把标题改大" 或 "排版这个文档"
    │
    ├─ wps-agent MCP 可用？
    │   ├─ YES → 加载 `document-author` skill
    │   │        Phase 1: wps-agent_content batch / semantic_structure / query_by_role
    │   │        Phase 2: 输出计划（含 surgical.select 定位）
    │   │        Phase 3: wps-agent_format set_font / set_text_effect
    │   │        Phase 4: wps-agent_layout auto_fix_layout 验证
    │   │
    │   └─ NO → 加载 `docx` skill（离线 XML/JS 操作）
    │            ↓
    └─ 完成 → 输出操作摘要
```

---

## document-author 智能操作

当涉及 WPS MCP Word 工具时，**必须优先加载 `document-author` skill**，严格遵循 4 Phase 工作流。

### v2.0 增强 Phase 1 读取

在原有 batch read 基础上，**必须额外调用**：

```json
{"tool": "wps-agent_content", "action": "semantic_structure", "filepath": "..."}
{"tool": "wps-agent_content", "action": "query_by_role", "sr": "abstract"}
```

目的：自动识别 20+ 语义角色、构建 DocumentGraph 关系网络。

### v2.0 增强 Phase 4 验证

排版质量评估必须加入：

```json
{"tool": "wps-agent_layout", "action": "auto_fix_layout", "filepath": "..."}
```

自动检测并修正：文本溢出、孤行标题、表格跨页、分栏不均衡。

---

## 封面制作规则（血泪教训）

| ❌ 错误 | ✅ 正确 |
|--------|--------|
| 多次 `insert_text` | `content action=create_cover lines=[...]` |
| 不清除段落间距 | `create_cover` 强制 SpaceBefore/After=0 |
| `doc.Range(0,End).Delete()` | `doc.Content.Text = ""` |

```json
{"tool": "wps-agent_content", "action": "create_cover", "clear_existing": true,
 "lines": [
   {"text": "标题", "font_name": "黑体", "font_size": 26, "bold": true, "alignment": "center", "space_before": 120, "space_after": 24},
   {"text": "副标题", "font_name": "宋体", "font_size": 16, "alignment": "center", "space_after": 6},
   {"text": "日期", "font_name": "宋体", "font_size": 14, "alignment": "center", "space_after": 6}
 ]}
```

---

## 环境

- Python：`E:\Anaconda\envs\wps-agent\python.exe`
- 禁止 `conda activate` 或裸 `python` 命令
- Offline 模式需先关闭 WPS 中文档（避免文件锁 `[WinError 32]`）

---

## 文档更新日志

| 日期 | 更新 |
|------|------|
| 2026-05-02 | 新增 v2.0 surgical/query_by_role/auto_fix_layout/set_text_effect 路由规则 |
