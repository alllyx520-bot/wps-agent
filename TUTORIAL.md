# WPS AI Agent 使用教程

> 一个深度集成 WPS Office 的 AI 排版专家，精通 Word 和 Excel，通过 MCP 协议与 AI Agent 实时交互。配备 `document-author` Skill，实现类人 4-Phase 文档工作流（理解→规划→执行→验证）。

---

## 一、环境准备

### 1.1 前置条件

| 组件 | 要求 |
|---|---|
| Windows | Windows 10/11 |
| WPS Office | 12.0+ (已安装并注册 COM) |
| Python | 3.11+ (conda env: `wps-agent`) |

### 1.2 安装依赖

```bash
# 创建专用 conda 环境
conda create -n wps-agent python=3.11 -y

# 安装核心依赖
pip install pywin32 pyyaml mcp httpx
```

### 1.3 配置 API Key

编辑 `config.yaml`，填入 LLM API 信息：

```yaml
llm:
  provider: "deepseek"
  endpoint: "https://api.deepseek.com/v1"
  model: "deepseek-chat"
  api_key: "sk-your-api-key"      # 替换为你的 API Key
  max_tokens: 4096
  temperature: 0.3
```

支持所有 OpenAI 兼容 API（DeepSeek、阿里云 DashScope、OpenAI 等），只需修改 `endpoint` 和 `model`。

### 1.4 接入 opencode

在 `C:\Users\<用户名>\.config\opencode\opencode.jsonc` 的 `mcp` 段添加：

```json
"wps-agent": {
  "type": "local",
  "command": [
    "E:\\Anaconda\\envs\\wps-agent\\python.exe",
    "E:\\AAAprojects\\自由测试\\wps-agent\\mcp_server.py"
  ]
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

### 2.1 打开或新建文档

在 WPS 中打开任意文档（包括云端文档），或让 Agent 创建：

```
在 WPS 中新建一个空白文档
```

Agent 会调用 `document create` 创建新文档，所有后续操作都实时反映在 WPS 窗口中。

### 2.2 查看文档信息

```
查看当前文档信息
列出所有打开的文档
```

示例返回：
```json
{
  "name": "报告.docx",
  "paragraph_count": 45,
  "table_count": 3,
  "section_count": 2,
  "saved": false
}
```

---

## 三、Word 排版操作

### 3.1 字体格式

**查看某段字体：**
```
查看第1段的字体格式
```

**修改字体：**
```
把第1段改成黑体三号加粗
把第3段改成楷体小四号
把选中的文字改成红色
```

MCP 工具调用示例（AI 自动转换自然语言为 COM 操作）：
```
tool: format
action: set_font
para_index: 1
name: 黑体
size: 16
bold: true
```

**支持的字体属性：**
| 属性 | 说明 | 示例 |
|---|---|---|
| name | 字体名 | 黑体, 宋体, 仿宋, 楷体, 微软雅黑 |
| name_far_east | 东亚字体 | 宋体, 黑体 |
| size | 字号(磅) | 42(初号), 22(二号), 16(三号), 12(小四) |
| bold | 粗体 | true/false |
| italic | 斜体 | true/false |
| underline | 下划线 | 0=无, 1=单线, 7=波浪线 |
| color_index | 颜色 | 0=自动, 1=黑, 2=蓝, 3=青, 4=绿, 5=品红, 6=红 |
| superscript | 上标 | true/false |
| subscript | 下标 | true/false |
| strike_through | 删除线 | true/false |

**中文字号对照表：**
| 字号 | 磅值 |
|---|---|
| 初号 | 42 |
| 小初 | 36 |
| 一号 | 26 |
| 小一 | 24 |
| 二号 | 22 |
| 小二 | 18 |
| 三号 | 16 |
| 小三 | 15 |
| 四号 | 14 |
| 小四 | 12 |
| 五号 | 10.5 |
| 小五 | 9 |

### 3.2 段落格式

**查看段落格式：**
```
查看第2段的段落格式
```

**修改段落格式：**
```
把第2段设置为两端对齐，首行缩进2字符，1.5倍行距
把所有正文段落设置为首行缩进2字符
```

**支持的段落属性：**
| 属性 | 说明 | 取值 |
|---|---|---|
| alignment | 对齐方式 | left, center, right, justify |
| first_line_indent | 首行缩进(磅) | 28≈2字符(14pt字) |
| left_indent | 左缩进 | 磅值 |
| right_indent | 右缩进 | 磅值 |
| line_spacing_rule | 行距规则 | single, 1.5lines, double, exactly, multiple |
| line_spacing | 行距值 | 配合规则使用 |
| space_before | 段前间距 | 磅值 |
| space_after | 段后间距 | 磅值 |
| outline_level | 大纲级别 | 1-9(标题级), 10(正文) |

### 3.3 样式管理

**查看样式：**
```
列出文档中的所有样式
查看"标题 1"样式的属性
```

**创建/修改样式：**
```
创建名为"我的标题"的新样式，字体微软雅黑14号加粗
修改"标题 2"样式为楷体三号
```

**应用样式：**
```
把第5段应用"标题 2"样式
把选中的段落应用"正文"样式
```

> 注意：WPS 内置样式名使用中文（"标题 1"而非"Heading 1"）。也可用数值索引访问：
> - Normal = -1
> - Heading 1 = -2, Heading 2 = -3, ..., Heading 9 = -10

**格式刷：**
```
把第3段的格式复制到第5、6、7段
```

### 3.4 内容读写

**读取文档：**
```
显示文档全文
读取第3段的内容
显示当前选中的文字
显示文档大纲
```

**修改内容：**
```
在第3段后面插入一段文字："这是新插入的内容"
在文档末尾追加一段总结
删除第5段
把第3段中的"旧文本"替换为"新文本"
```

---

## 四、表格操作

### 4.1 创建表格

```
在文档末尾插入一个4行3列的表格
```

### 4.2 读写单元格

```
在表格第1行第1列填入"姓名"
在表格第1行第2列填入"部门"
读取表格1的全部内容
```

### 4.3 格式化表格

```
将表格第1行设为标题行（黑体加粗）
给表格添加边框
将表格列宽自适应内容
设置第1列宽度为100磅
给表格应用隔行变色
```

### 4.4 合并单元格

```
合并表格1中第1行第2列到第1行第3列的单元格
```

---

## 五、页面布局

**页面设置：**
```
将页面设置为A4纸，上下边距2.54cm，左右边距3.17cm
设置页面为横向
```

**分栏：**
```
将当前节设置为两栏
```

**页眉页脚：**
```
设置页眉为"XX公司年度报告"
添加居中页码
```

**分节：**
```
在第5段之后插入分节符（下一页）
```

---

## 六、查找替换与审阅

### 6.1 查找替换

```
在文档中搜索"项目"
把所有"旧词"替换为"新词"
查找所有黑体加粗的文字
跳转到第2级标题
```

### 6.2 修订与批注

```
开启修订模式
在第3段添加批注："此处需要补充数据"
列出文档中的所有批注
列出所有修订
接受所有修订
拒绝所有修订
```

---

## 七、AI 智能排版（核心功能）

这是 WPS Agent 最具竞争力的功能，利用 LLM 实现超越人类的排版理解。

### 7.1 分析文档结构

```
分析当前文档的排版结构
```

AI 会：
1. 读取文档大纲和格式采样
2. 通过 LLM 识别文档类型（公文/论文/报告/通用）
3. 识别标题层级结构
4. 发现格式不一致问题
5. 生成改进建议

示例输出：
```
文档类型: report
标题层级:
  Level 1: 黑体 x3 章
  Level 2: 宋体/Calibri x4 节
不一致: 4处
  - 一级标题格式不统一，段落1使用宋体
  - 二级标题英文字体为Calibri，中文字体应为宋体
建议: 应用统一样式模板
```

### 7.2 生成排版建议

```
建议如何改进当前文档的排版
```

AI 会生成具体的格式化操作序列，每条建议包含：
- 操作工具、动作、参数
- 操作原因说明

### 7.3 应用预设模板

内建 4 套专业模板：

| 模板名 | 适用场景 | 标题 | 正文 |
|---|---|---|---|
| `official` | 公文 | 方正小标宋二号 居中 | 仿宋三号 2字符缩进 28磅行距 |
| `thesis` | 学术论文 | 黑体三号 居中 | 宋体小四 2字符缩进 1.5倍行距 |
| `report` | 商业报告 | 微软雅黑二号 居中 | 微软雅黑五号 1.3倍行距 |
| `resume` | 简历 | 黑体二号 居中 | 宋体五号 1.3倍行距 |

使用方法：
```
对当前文档应用论文模板
对当前文档应用公文模板
```

模板会匹配标题大纲级别，自动应用对应格式。页面设置（纸张、边距）也会一并调整。

### 7.4 AI 自然语言排版

**这是最强大的功能——直接用自然语言描述排版需求：**

```
将第一章和第二章的标题改为黑体三号加粗居中，正文全部改为宋体小四号两端对齐首行缩进两字符
```

```
把标题改大一点，正文行距调到1.5倍，所有表格加上边框
```

```
把文档按照学术论文格式排版，要有目录和多级编号
```

AI 工作流程：
1. 解析自然语言指令
2. 通过 LLM 转换为具体 COM 操作序列
3. 在 WPS 中执行所有操作
4. 返回执行结果

### 7.5 自动目录

```
生成文档目录
```

AI 会基于标题的 OutlineLevel 自动生成 1-3 级目录。

### 7.6 自动多级编号

```
给文档标题添加多级编号
```

AI 会根据标题大纲级别自动生成编号：
- 1级 → 1, 2, 3
- 2级 → 1.1, 1.2, 2.1
- 3级 → 1.1.1, 1.1.2

### 7.7 排版质量校验

```
校验当前文档的排版质量
```

检查项：
- 标题层级是否连贯（不跳跃）
- 格式一致性
- 结构完整性

---

## 八、Excel 操作

### 8.1 工作簿管理

```
新建一个 Excel 工作簿
打开 E:\data\销售报表.xlsx
保存工作簿
列出所有打开的工作簿
```

### 8.2 工作表操作

```
查看工作表列表
切换到"销售数据"工作表
新建一个名为"汇总"的工作表
```

### 8.3 读写数据

**写入数据：**
```
在A1到D5写入数据：
姓名  部门  销售额  季度
张三  销售  12万   Q1
李四  销售  9.5万  Q1
王五  市场  8万    Q1
赵六  技术  15万   Q1
```

**读取数据：**
```
读取A1到D5的数据
读取C2单元格的值
```

### 8.4 格式化

```
将A1:D1的表头设置为黑体12号加粗，灰色背景
给A1:D5添加边框
设置A列列宽为15
自动调整B到D列列宽
```

### 8.5 公式

```
在C6单元格设置公式=SUM(C2:C5)
在D2设置公式=C2*0.1
```

### 8.6 合并单元格

```
合并A1到D1单元格
```

---

## 九、云端文档处理

WPS Agent 天然支持 WPS 云端文档，无需特殊操作：

1. 在 WPS 中打开云端文档（WPS 自动下载到本地缓存）
2. Agent 通过 COM 连接到运行中的 WPS 实例
3. 所有修改实时生效
4. 保存时 WPS 自动同步回云端

**对用户来说完全透明，无需额外操作。**

---

## 十、所有 MCP 工具一览

| 工具名 | 常用操作 | 说明 |
|---|---|---|
| `document` | create/list/open/save/close/export_pdf | 文档管理 |
| `content` | full_text/paragraph/outline/insert/delete | 内容读写 |
| `format` | get_font/set_font/get_paragraph_format/set_paragraph_format | 格式操作 |
| `style` | list/get/create/modify/apply_style | 样式管理 |
| `table` | count/create/read/format_cell/set_header/merge/borders | 表格操作 |
| `search` | find/replace/find_format/goto_heading | 查找替换 |
| `layout` | page_setup/section_info/columns/header_footer/page_numbers | 页面布局 |
| `review` | track_changes/comments/revisions | 审阅批注 |
| `ai_format` | analyze/suggest/apply_template/reformat/auto_toc/auto_numbering/validate | AI 排版 |
| `excel` | create/open/cell_read/cell_write/range_read/range_write/font_set/formula_set | Excel 操作 |

---

## 十一、典型工作流

### 场景1：撰写论文

```
1. 在 WPS 中新建文档
2. "请帮我写一篇关于人工智能在医疗领域应用的论文大纲"
3. [手动填充各章节内容]
4. "对文档应用论文模板"
5. "给标题添加多级编号"
6. "生成文档目录"
7. "分析文档排版，检查格式问题"
8. "根据建议调整格式"
9. 导出 PDF
```

### 场景2：制作 Excel 报表

```
1. "新建 Excel 工作簿"
2. "在 A1 写入表头：月份,销售额,成本,利润"
3. [填入数据]
4. "在 E2 设置公式 =C2-D2，下拉到 E13"
5. "在 C14 设置公式 =SUM(C2:C13)"
6. "表头设为黑体加粗，灰色背景"
7. "给数据区域添加边框"
8. "自动调整所有列宽"
9. "保存为 月度报表.xlsx"
```

### 场景3：排版公文

```
1. 打开已有文档
2. "对文档应用公文模板"
3. "分析公文排版格式是否符合标准"
4. "将标题字体改为方正小标宋简体"
5. "检查所有段落首行缩进"
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

## 十二、常见问题

**Q: Agent 无法连接 WPS？**
A: 确保 WPS 已安装（非绿色版），COM 组件已注册。检查注册表 `HKCR\Kwps.Application` 是否存在。

**Q: 修改格式没有生效？**
A: 检查 WPS 窗口是否为可见状态。COM 操作是同步的，修改会立即反映。

**Q: 中文样式名找不到？**
A: WPS 内置样式使用中文名（"标题 1"而非"Heading 1"）。也可以使用数值索引。

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
