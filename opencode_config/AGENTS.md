# AGENTS.md — Vibe Archive / VibeFlow

> 本文件是最高优先级指令。所有规则均为硬性约束，除非架构师明确覆盖。

---

## 0. Quick Reference (紧急情况速查)

| 场景 | 规则 |
|------|------|
| 修改 `.tsx/.css` | 只改逻辑，Tailwind 类名和 Framer Motion 动效 **绝对不动** |
| 运行 Python | 用绝对路径 `E:\Anaconda\envs\<env>\python.exe`，**禁止** `conda activate` |
| 训练中断 | 优先 `resume`，**禁止** 截断/降 Epoch/简化版本 |
| 脚本报错 | **原地修复**，禁止创建 `*_simple.py` 等临时文件 |
| 5+ 文件重构 | 先输出 `.md` 执行计划，**等审批**再动手 |
| 不确定环境名 | **直接问**，禁止猜测 |
| 完全卡住 | 列 2-3 个方案 + 优缺点，等架构师决策 |
| 后台训练 | **必须用 `pythonw.exe`** 启动训练，opencode 关闭不会杀子进程 |

---

## 1. Anti-Fluff & Anti-Docs Policy

### 1.1 输出纪律
- **零开场白**：不写"好的"、"我来帮你"、"接下来我将"等任何铺垫
- **零总结**：不写"以上就是..."、"总结一下"等收尾语
- **直接给结果**：代码、命令、文件路径，仅此而已

### 1.2 文档禁令
- **禁止创建**：README、技术路线图、项目日志、CHANGELOG、架构文档等任何 `.md` 文件
- **唯一触发词**：只有当我明确说出 **"DOCUMENT"** 时，才允许创建文档
- **SOLE EXCEPTION（编排规则）**：涉及 **5 个以上文件修改** 的全局重构前，必须先输出 `.md` 格式的执行计划，并等待我的明确审批（回复"批准"或"执行"）

### 1.3 沟通语言
- 对话用 **中文**
- 技术术语、代码、变量名、日志、文件路径 **全部用英文**
- 回复极度简洁，能用一行不用两段

---

## 2. Ultimate Vibe & UI Protection

### 2.1 Tailwind 类名铁律
在修改 `.tsx` / `.jsx` / `.css` / `.scss` 文件时：
- **绝对禁止** 删除、简化、合并、重排现有的 Tailwind 类名
- **绝对禁止** 用"更简洁的写法"替换现有类名组合
- 只能修改与 **逻辑或数据变量** 直接相关的部分

### 2.2 Framer Motion 动效保护
- **绝对禁止** 修改 `animate`、`transition`、`whileHover`、`whileTap`、`layout` 等动效属性
- **绝对禁止** 删除 `motion.div`、`motion.button` 等组件包裹
- 只能修改动效组件内部的 **业务逻辑** 或 **数据绑定**

### 2.3 视觉 Vibe 校验清单
每次修改 UI 相关文件后，自检：
- [ ] Tailwind 类名数量未减少
- [ ] 颜色值（`zinc-950`、`emerald-500` 等）未改变
- [ ] 间距值（`p-4`、`gap-2` 等）未改变
- [ ] 圆角/阴影/边框类名未改变
- [ ] Framer Motion 配置未改变
- [ ] 只改了 `if/else`、变量名、函数调用、数据流

---

## 3. Environment & Runtime Enforcement

### 3.1 虚拟环境优先原则
- **绝对禁止** 在系统环境（base/system Python）中直接运行项目或安装依赖
- **执行前必须** 先查找与项目相关的 conda 虚拟环境
- **必须** 使用 conda 环境的绝对路径执行 Python：
  ```bash
  # ✅ 正确
  E:\Anaconda\envs\dnncnn-env\python.exe train.py

  # ❌ 错误
  python train.py
  conda activate dnncnn-env && python train.py
  ```

### 3.2 指纹校验（运行前必做）
执行任何核心脚本或训练循环前，**必须** 先运行：
```bash
[absolute_python_path] -c "import sys; print(sys.executable)"
```

### 3.3 Fail-Fast 机制
以下情况视为 **FATAL ERROR**，立即停止并报告：
- Python 版本不匹配项目要求
- `torch.cuda.is_available()` 返回 `False` 但项目需要 GPU
- 关键依赖缺失且无法自动安装
- 环境变量未正确加载

### 3.4 环境名确认
- 如果存在多个相似环境，**直接问我** 正确名称
- **禁止** 根据命名规律猜测
- **禁止** 尝试多个环境直到碰对

---

## 4. Execution & Verification Workflow

### 4.1 颗粒度确认
复杂任务必须 **分步确认**：
```
Step 1: 安装依赖 → 完成（贴终端日志）
Step 2: 数据预处理 → 完成（贴终端日志）
Step 3: 启动训练 → 完成（贴终端日志）
```

### 4.2 零简化原则
- **禁止** 提供 `TODO`、`pass`、`# placeholder` 等占位符
- **禁止** 提供"示例代码"替代核心逻辑
- 每一行代码必须是 **生产就绪** 的完整实现

### 4.3 断点续训协议（Resume Protocol）
训练任务因超时/报错中断时：
1. **优先方案**：查找 checkpoint，使用 `resume` 参数续训
2. **次优方案**：从最近一次成功 epoch 重新开始
3. **禁止**（除非架构师明确授权）：
   - 减少 Epoch 数量
   - 降低 batch size 强行跑通
   - 使用简化版数据集
   - 跳过验证步骤

### 4.4 根进程终止协议（opencode 关闭不杀后台训练）

**血泪教训：** 用 `Start-Process` 或 `python.exe` 启动的训练进程是 opencode 的子进程，opencode 关闭时会被一起杀光。

**正确做法：**
- **必须** 使用 `pythonw.exe`（无控制台窗口的 Python）启动训练
- **必须** 用 `cmd /c start /B` 或直接调用 `pythonw.exe` 来脱离进程树
- **命令模板：**
  ```powershell
  cmd.exe /c start /B /D "E:\AAAprojects\PIX2PIX" "" "E:\Anaconda\envs\<env>\pythonw.exe" -u train.py > logs/output.log 2> logs/error.log
  ```
- `pythonw.exe` 启动的进程在 opencode 关闭后**不会**被终止
- 可用 `Get-Process -Name pythonw` 验证进程存活
- **对应快速启动脚本**：项目根目录放 `start_training.bat` 双击即可启动

### 4.5 自动验证与自修复
- 每次修改后，自动运行相关验证命令
- 如果命令失败，**立即修复**，不询问权限
- 修复后重新验证，直到通过

---

## 5. Coding Standards

### 5.1 代码风格
- 遵循项目现有代码风格为 **绝对参考**
- 变量命名：英文，驼峰或蛇形，与项目现有风格一致
- 注释：只在逻辑复杂处加简短英文注释

### 5.2 工具使用
- 搜网/查资料/查 API 用法时，使用 `brave-search` 工具
- 需要浏览器测试 UI 或抓取页面时，使用 `playwright` 工具
- GitHub 仓库搜索/Issue/PR 操作，使用 `github` 工具
- 不确定 API 用法时，先查证再写代码

### 5.3 阻塞处理
完全卡住时，按格式输出：
```
阻塞原因：[一句话描述]
方案 A：[描述] | 优点：[x] | 缺点：[y]
方案 B：[描述] | 优点：[x] | 缺点：[y]
方案 C：[描述] | 优点：[x] | 缺点：[y]
请架构师决策。
```

---

## 6. File System Discipline

### 6.1 原地修复铁律
- 当 `train.py` 或任何核心脚本报错时，**必须** 直接在原文件上修改修复
- **绝对禁止** 通过创建新文件来逃避 Debug

### 6.2 禁止临时文件
- **禁止** 创建 `train_simple.py`、`test_basic.py`、`debug.py`、`temp.py` 等降级/临时文件
- 如果确实需要极简测试脚本验证 API，验证完成后 **立即删除**

### 6.3 根目录卫生
项目根目录只允许存在架构规划内定义的核心文件：
- **日志** → 统一放 `logs/`
- **模型权重** → 统一放 `checkpoints/`
- **测试截图** → 统一放 `screenshots/` 或 `tests/`
- **临时文件** → 用完即删
- **禁止** 在根目录散落任何非核心文件

### 6.4 文件创建审批
创建新文件前自检：
- [ ] 是否真的需要新文件？能否原地修改？
- [ ] 新文件是否符合项目目录结构规范？
- [ ] 是否属于核心架构规划内的文件？
- 任一答案为否，先问我。

### 6.5 配置备份策略
- **MCP 配置修改**：必须等待用户明确回复"重启成功"或"MCP 加载正常"后再执行备份，因为 MCP 错误会导致 opencode 无法启动，需先验证。
- **其他配置修改**（Skills、Agents、Commands、AGENTS.md 等）：修复验证通过后即可直接备份，无需等待用户确认重启。
- 备份统一复制 `~/.config/opencode` → `~/.config/opencode - 备份`。

---

## 7. Security & Secrets

- **禁止** 在代码中硬编码 API Key、密码、Token
- **禁止** 提交 `.env` 文件到版本控制
- 敏感信息统一从环境变量或 `.env`（已加入 `.gitignore`）读取
- 发现代码中有硬编码密钥，立即替换为环境变量引用
- **例外**：`opencode.jsonc`（MCP 配置）中的 Token 在本机非公开环境下可容忍明文存储，但切勿提交该文件

---

## 8. Decision Boundaries

### 8.1 自主决策（无需询问）
- 修复明确的 bug
- 补充缺失的类型注解
- 优化明显低效的代码（不改变行为）
- 添加缺失的错误处理
- 运行验证命令
- 删除不必要的冗余过程文件

### 8.2 必须询问
- 改变现有 API 接口签名
- 删除现有功能或文件
- 更换技术栈或核心依赖
- 不确定环境名称或配置
- 完全阻塞需要架构决策

### 8.3 建议但可执行
- 添加新的工具函数（不破坏现有逻辑）
- 重构单个文件内部结构（不跨文件）
- 补充单元测试（不修改业务代码）

---

## 9. Git Workflow

### 9.1 提交规范
- 只有我明确要求时才执行 git commit，不主动提交
- 禁止 force push，尤其禁止对 main/master 分支 force push
- 禁止使用 `--no-verify` 或 `--no-gpg-sign` 跳过钩子
- Commit message：1-2 句话，说"为什么"而不是"改了什么"
- 提交前先检查 `git status` 和 `git diff`，确认内容正确
- 如果暂存区包含 `.env`、`credentials.json` 等密钥文件，必须警告

### 9.2 PR 规范
- PR 描述要包含该分支上所有 commit 的变更总结，而不仅是最后一个 commit
- 提交前先检查变更范围是否正确

### 9.3 冲突处理
- 解决合并冲突时，先读两边代码的逻辑，再选择正确方案
- 不确定时问我

---

## 10. Skill & MCP Invocation Guide

当用户提出以下需求时，按优先级选择工具：

### 10.1 Office 文档操作（最高优先级 → wps-agent MCP）

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
- 不要同时加载 wps-agent 和 skill，避免冲突

### 10.1.1 WPS Agent 封面制作血泪教训

> 以下错误各造成 3-5 次阻塞，必须写入铁律。

| 错误 | 现象 | 根因 | 正确做法 |
|------|------|------|----------|
| 多次 `insert_text` 不分段 | 所有文本挤在同一段落，后面格式设置失效 | `insert_text` 的 `\n` 虽然是分段逻辑但 AI 不敢赌，逐次调用时文本合并 | **必须用 `content action=create_cover`**，传入 `lines` 数组，一行调用搞定 |
| `delete_range` 只传 `start_pos` 不传 `end_pos` | 文档清不干净，段落标记残留，后续插入错位 | `end_pos` 原为必传参数，不传等价于 0，范围无效 | `delete_range` 现已支持省略 `end_pos`（默认文档末尾），但仍优先用 `create_cover` 自带 `clear_existing` |
| `format batch` 操作数不确定 | batch 参数格式报错，不确定 `type` 字段写 `set_font` 还是 `set_paragraph_format` | AI 记忆偏差，误认为 batch 接口是"无描述语法" | `create_cover` 内置逐段格式化，无需手动 batch |
| 不清除继承的段落间距 | WPS Normal 样式自带段前/段后间距，未显式 `space_before/space_after` 时文本间距失控 | `_apply_line_format` 之前只在 `line` 指定了 `space_before/space_after` 时才设置，未指定则走样式默认值 | `_apply_line_format` 现已强制 `SpaceBefore=0, SpaceAfter=0`（`line` 有值时覆盖） |
| `doc.Range(0, Content.End).Delete()` 清空 | 封面有两页、文档残留旧内容混排 | Word/WPS 拒绝删除末尾段落标记，`Delete()` 抛异常后 fallback `doc.Range(1, Content.End).Delete()` 无效，旧内容残留 | 用 `doc.Content.Text = ""` 可靠清空 |

**封面标准流程（一调用搞定）：**
```json
{
  "tools": ["content"],
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

- `create_cover` 自动管理段落创建、格式设置、清空旧内容
- `lines` 数组每项配置：`text/font_name/font_size/bold/alignment/space_before/space_after/line_spacing_rule/line_spacing/first_line_indent`

### 10.1.2 document-author 智能化操作

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

### 10.2 其他 Skill（非 Office 场景）

| 用户需求 | 加载的 skill |
|---|---|
| 代码审查、安全检查、性能分析 | `code-review` |
| 报错排查、Bug 定位修复 | `debug-troubleshoot` |
| 创建新 agent（"帮我写个 agent"、"创建一个 agent 做..."） | `agent-creator` |
| 创建新 command（"写个命令"、"帮我加个快捷指令"） | `command-creator` |
| 创建新 skill（"把这个流程做成 skill"） | `skill-creator` |

**规则：**
- 如果需求跨多个 skill，按主次顺序依次加载
- 不要在用户没提时自作主张调用

---

## 11. MCP Configuration Rules

配置 MCP Server 时，**必须使用 opencode 格式，不是 Claude Desktop 格式**。

**正确格式（opencode）：**
```json
{
  "mcp": {
    "name": {
      "type": "local",
      "command": ["npx", "-y", "package-name"],
      "environment": {
        "KEY": "value"
      }
    }
  }
}
```

**常见错误：**
- ❌ `"type": "stdio"` → 应为 `"type": "local"`
- ❌ `"command": "npx"` + `"args": [...]` → 应合并为 `"command": ["npx", "-y", "..."]`
- ❌ `"env": { ... }` → 应为 `"environment": { ... }`

**验证规则：**
- 写入 MCP 配置前，必须用 `npx -y package-name` 验证包真实存在，禁止凭记忆或猜测填写包名（曾把 `@github/mcp-server-github` 错写成 `@modelcontextprotocol/server-github`，把不存在的 `@modelcontextprotocol/server-fetch` 填入配置导致 404）
- 写入后提醒用户重启测试。MCP 配置错误会导致 opencode 无法启动（"连接不上本地服务器"）
- **网络环境校验**：MCP 包能下载 ≠ API 能访问。校园网/防火墙可能拦截某些域名（fetch MCP 无法访问 GitHub/npm，brave-search API 也被拦截），但 github MCP 通过 REST API 直连不受影响。安装后必须实际调用一次工具验证网络连通性
- **失效 MCP 替代方案**：
  - `fetch` → **永久不可用**：TUN 模式下域名被解析到虚拟网卡 IP（`198.18.0.x`），触发 opencode 内置 SSRF 保护
  - `brave-search` → **TUN 依赖**：TUN 开启时内置 `brave-search` MCP 正常工作；TUN 关闭时 Node.js 不走系统代理导致失效。后备：`/search` command → Python `web_tools.py search`
- **工具选择优先级**：当需要获取网页/网络信息时，优先使用 `brave-search` MCP，Playwright 作为低优先级替代即可

---

## 12. Verification Checklist (每次任务完成前自检)

- [ ] 没有创建任何未经授权的 `.md` 文档
- [ ] 回复中没有开场白/总结语/废话
- [ ] 技术术语和代码全部使用英文

---

## 13. 持续学习机制 (Self-Correction & Learning)

- 当 opencode 在执行任务中遇到错误，经过排查找到根因并修复后，**必须**将教训总结更新到 AGENTS.md 对应章节中。
- 更新形式：在相关规则下添加"曾犯错误"或"血泪教训"条目，说明错误场景、错误做法、正确做法。
- **禁止**以"这是常识"或"下次注意"为由不更新文档。任何导致任务阻塞或需要用户介入才能发现的错误，都值得被记录。
- 示例：在 MCP 配置规则中记录"曾把 `@github/mcp-server-github` 错写成 `@modelcontextprotocol/server-github`，把不存在的 `@modelcontextprotocol/server-fetch` 填入配置导致 404"。
