# AGENTS.md — 启用 WPS Agent 智能化功能

> 将本文件复制到 `~/.config/opencode/` 以启用 WPS Agent 的类人智能化文档操作。
> 部署命令：`robocopy opencode_config\ %USERPROFILE%\.config\opencode\ /E`

---

## 1. WPS Office 文档操作优先级

WPS Office 相关需求**必须优先使用 wps-agent MCP**，它直接通过 COM 控制实时 WPS 窗口，效果远好于离线脚本生成。

| 用户需求 | 优先工具 | 备用方案 |
|---|---|---|
| 编写/编辑/排版 Word 文档 | `wps-agent` MCP（document/content/format/table/style/layout/review/reference 等工具） | `docx` skill（仅当 WPS 不可用时） |
| 制作 PPT 演示文稿 | `wps-agent` MCP（presentation 工具） | `pptx` skill |
| 创建/编辑 Excel 表格、数据分析 | `wps-agent` MCP（excel 工具） | `xlsx` skill |
| AI 智能排版、模板应用、目录生成 | `wps-agent` MCP（ai_format/template 工具） | — |
| Word→PPT 大纲转换 | `wps-agent` MCP（migrate 工具） | — |
| 跨文档复制/迁移 | `wps-agent` MCP（transfer/migrate 工具） | — |
| 文档对比 | `wps-agent` MCP（compare 工具） | — |
| 统一管理多个打开的文档 | `wps-agent` MCP（docspace 工具） | — |

**规则：**
- 用户提及"Word/文档/排版/论文/报告/合同/目录/书签/页眉页脚/脚注/批注/水印"等 → 优先用 `wps-agent` MCP
- 用户提及"PPT/幻灯片/演示文稿"等 → 优先用 `wps-agent` MCP 的 `presentation` 工具
- 用户提及"Excel/表格/图表/公式/数据"等 → 优先用 `wps-agent` MCP 的 `excel` 工具
- **只有在 WPS Agent MCP 明确不可用或连接失败时**，才回退到 `docx`/`pptx`/`xlsx` skill
- 不要同时加载 wps-agent 和离线 skill，避免冲突

---

## 2. document-author 智能化操作

当涉及 WPS MCP Word 工具操作时，**必须优先加载 `document-author` skill**，然后严格遵循 4 Phase 工作流：

1. **Phase 1: 理解** — 任何工具调用前必须先用 `batch` 读取文档全文、大纲结构、格式样本，构建"文档心智模型"
2. **Phase 2: 规划** — 必须先输出自然语言修改计划（含影响分析、步骤分解），再动手
3. **Phase 3: 执行** — 逐步操作，每步记录状态（已完成/进行中/待完成）
4. **Phase 4: 验证** — 全部完成后重读修改区域，执行一致性检查（Consistency Guard），发现问题立即修正

**铁律：**
- **禁止盲写盲改**：任何 `insert_text` / `set_font` / `delete_range` 等写操作前必须先读过文档
- **禁止无规划执行**：必须先输出计划再动手
- **禁止改完就走**：必须验证修改结果
- **禁止无脑套标准**：格式决策必须基于文档自身风格发现 + conventions 参考，优先匹配文档现有格式

---

## 3. 封面制作注意事项

| 错误 | 现象 | 根因 | 正确做法 |
|------|------|------|----------|
| 多次 `insert_text` 不分段 | 所有文本挤在同一段落，后面格式设置失效 | 逐次调用时文本合并 | **必须用 `content action=create_cover`**，传入 `lines` 数组，一行调用搞定 |
| `delete_range` 只传 `start_pos` 不传 `end_pos` | 文档清不干净，段落标记残留 | 不传 end_pos（或传 0）范围无效 | `delete_range` 现已支持省略 `end_pos`（默认文档末尾），但仍优先用 `create_cover` 自带 `clear_existing` |
| 不清除继承的段落间距 | WPS Normal 样式自带段前/段后间距，文本间距失控 | 未显式设置间距时走样式默认值 | `create_cover` 强制 `SpaceBefore=0, SpaceAfter=0`（line 有值时覆盖） |
| `doc.Range(0, Content.End).Delete()` 清空 | 文档残留旧内容混排 | Word/WPS 拒绝删除末尾段落标记 | 用 `doc.Content.Text = ""` 可靠清空 |

**封面标准流程：**
```json
{
  "content": {
    "action": "create_cover",
    "clear_existing": true,
    "lines": [
      {"text": "标题", "font_name": "黑体", "font_size": 26, "bold": true, "alignment": "center", "space_before": 120, "space_after": 24},
      {"text": "副标题", "font_name": "宋体", "font_size": 16, "alignment": "center", "space_after": 6},
      {"text": "日期", "font_name": "宋体", "font_size": 14, "alignment": "center", "space_after": 6}
    ]
  }
}
```

---

## 4. 环境与 Python

- 运行 Python 必须用 conda 环境的绝对路径：`E:\Anaconda\envs\wps-agent\python.exe`
- **禁止** `conda activate` 或直接用 `python` 命令
- wps-agent 的 conda 环境名：`wps-agent`

---

## 5. 其他可用 Skill

| 用户需求 | 加载的 skill |
|---|---|
| 代码审查、安全检查、性能分析 | `code-review-and-quality` |
| 报错排查、Bug 定位修复 | `debugging-and-error-recovery` |
| 创建新 agent | `agent-creator` |
| 创建新 command | `command-creator` |
| 创建新 skill | `skill-creator` |
