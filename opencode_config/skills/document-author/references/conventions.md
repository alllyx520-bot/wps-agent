# Chinese Document Formatting Conventions (Reference)

> **IMPORTANT**: These are REFERENCE only, NOT mandatory rules.
> Always discover the document's own style first, then use these conventions
> as fallback guidance when no existing pattern exists.

---

## GB/T 9704-2012 — Government Documents (公文)

Official Chinese government document formatting standard.

### Page Layout
| Property | Value |
|----------|-------|
| Paper size | A4 (210mm × 297mm) |
| Top margin | 37mm ± 1mm |
| Bottom margin | 35mm ± 1mm |
| Left margin | 28mm ± 1mm |
| Right margin | 26mm ± 1mm |

### Font Hierarchy
| Element | Font (中文) | Font Size | Size (pt) | Style |
|---------|-----------|-----------|-----------|-------|
| 大标题 (Main title) | 方正小标宋简体 | 二号 | 22 | Regular, centered |
| 一级标题 (H1) | 黑体 | 三号 | 16 | Regular |
| 二级标题 (H2) | 楷体 | 三号 | 16 | Regular |
| 三级标题 (H3) | 仿宋 | 三号 | 16 | Bold |
| 正文 (Body) | 仿宋 | 三号 | 16 | Regular |
| 页码 (Page number) | — | 四号 | 14 | — |

### Line Spacing
- Body text: Fixed 28pt (固定值 28 磅)
- Title spacing: Adjust by visual balance

### Numbering
- Level 1: 一、二、三、...
- Level 2: (一)(二)(三)...
- Level 3: 1. 2. 3. ...
- Level 4: (1)(2)(3)...

---

## Academic Thesis — University Standard (高校学位论文)

Common conventions found in Chinese university theses and dissertations.

### Cover Page (封面)
| Element | Font | Size (pt) | Style |
|---------|------|-----------|-------|
| University name | 黑体 | 一号 (26) or 小初 (36) | Bold, centered |
| Thesis title | 黑体 | 一号 (26) or 小初 (36) | Bold, centered |
| Subtitle | 黑体 | 二号 (22) | Centered |
| Author/Advisor/etc. | 宋体/黑体 | 三号 (16) or 四号 (14) | Centered |
| Date | 宋体 | 三号 (16) or 四号 (14) | Centered |

### Abstract (摘要)
| Element | Font | Size (pt) | Style |
|---------|------|-----------|-------|
| "摘要" heading | 黑体 | 三号 (16) | Bold, centered |
| Abstract body | 宋体 | 小四 (12) | Regular |
| Keywords label | 黑体 | 小四 (12) | Bold |
| Keywords content | 宋体 | 小四 (12) | Regular |

### Body Text Hierarchy
| Element | Font | Size (pt) | Style |
|---------|------|-----------|-------|
| 一级标题 (Chapter) | 黑体 | 三号 (16) | Bold, left-aligned |
| 二级标题 (Section) | 黑体 | 四号 (14) | Bold, left-aligned |
| 三级标题 (Subsection) | 黑体 | 小四 (12) | Bold, left-aligned |
| 正文 (Body) | 宋体 | 小四 (12) | Regular |
| 图注 (Figure caption) | 宋体 | 五号 (10.5) | Centered |
| 表头 (Table header) | 黑体 | 五号 (10.5) | Bold, centered |
| 表内文字 (Table body) | 宋体 | 五号 (10.5) | Regular |
| 页眉 (Header) | 宋体 | 五号 (10.5) | Regular |
| 页码 (Page number) | — | 五号 (10.5) | Centered |

### Paragraph Format
| Property | Value |
|----------|-------|
| Body text first-line indent | 2 Chinese chars (≈ 24pt at 小四) |
| Line spacing | 1.5 lines (1.5 倍行距) |
| Paragraph spacing before/after | 0pt |
| Chapter starting page | New page |

### References (GB/T 7714)
Common reference types:
- Journal: `作者. 题名[J]. 刊名, 年, 卷(期): 起止页码.`
- Book: `作者. 书名[M]. 出版地: 出版社, 年.`
- Thesis: `作者. 题名[D]. 学校, 年.`
- Patent: `作者. 专利名[P]. 专利号, 年.`

---

## Lab Report / Experiment Report (实验报告)

Common structure and format for university lab reports.

### Structure
```
封面 (Cover)
  ├── 实验名称 (Experiment name)
  ├── 课程名称 (Course name)
  ├── 姓名/学号 (Name/Student ID)
  ├── 班级 (Class)
  ├── 实验日期 (Date)
  └── 指导教师 (Advisor)

正文 (Body)
  ├── 一、实验目的 (Purpose)
  ├── 二、实验原理 (Principles)
  ├── 三、实验器材 (Equipment)
  ├── 四、实验步骤 (Procedure)
  ├── 五、实验数据 (Data)
  └── 六、实验结论 (Conclusion)
```

### Format Conventions
| Element | Font | Size (pt) | Style |
|---------|------|-----------|-------|
| Cover title | 黑体 | 二号 (22) or 一号 (26) | Bold, centered |
| Cover info | 宋体 | 四号 (14) | Centered |
| Section heading | 黑体 | 小三 (15) or 四号 (14) | Bold, left |
| Body text | 宋体 | 小四 (12) | Regular, first-line indent 2 chars |
| Table content | 宋体 | 五号 (10.5) | Regular |
| Figure captions | 宋体 | 五号 (10.5) | Centered |

---

## WPS Point Size Reference

| Chinese Size | Points | Approx. mm |
|-------------|--------|------------|
| 初号 | 42 | 14.82 |
| 小初 | 36 | 12.70 |
| 一号 | 26 | 9.17 |
| 小一 | 24 | 8.47 |
| 二号 | 22 | 7.76 |
| 小二 | 18 | 6.35 |
| 三号 | 16 | 5.64 |
| 小三 | 15 | 5.29 |
| 四号 | 14 | 4.94 |
| 小四 | 12 | 4.23 |
| 五号 | 10.5 | 3.70 |
| 小五 | 9 | 3.18 |

### Common Indent Conversions
- 2 Chinese chars at 小四 (12pt) ≈ 24pt first-line indent
- 2 Chinese chars at 三号 (16pt) ≈ 32pt first-line indent
- 2 Chinese chars at 五号 (10.5pt) ≈ 21pt first-line indent

### Common Margin Conversions
- Standard word margin 2.54cm ≈ 72pt
- GB/T 9704 top margin 3.7cm ≈ 104.9pt
- Typical thesis top margin 2.5cm ≈ 70.9pt
