# Chinese Document Formatting Conventions — v2.0 Reference

> **IMPORTANT**: These are REFERENCE only, NOT mandatory rules.
> Always discover the document's own style first (via Phase 1 format sampling),
> then use these conventions as fallback. The wps-agent formatter implements
> 14 templates in `intelligence/chinese_rules.py`.

---

## 14 Built-in Templates Quick Reference

| Template | Scene | Heading | Body | Spacing |
|----------|-------|---------|------|---------|
| `official` | 党政公文 (GB/T 9704) | 方正小标宋 二号/center | 仿宋 三号/2char indent | 28pt fixed |
| `thesis` | 学术论文 | 黑体 三号/center(H1) 四号/left(H2) | 宋体 小四/2char indent | 1.5x |
| `report` | 商业报告 | 微软雅黑 二号/center | 微软雅黑 五号 | 1.3x |
| `resume` | 简历 | 黑体 二号/center | 宋体 五号 | 1.3x |
| `contract` | 合同/协议 | 黑体 二号/center | 宋体 小四/justify | 1.5x |
| `bid` | 标书 | 黑体 二号/center | 宋体 小四/justify | 1.5x |
| `exam` | 试卷 | 黑体 三号/bold(H1) | 宋体 小四/left | 1.5x |
| `press_release` | 新闻稿 | 黑体 二号/center | 宋体 小四/2char indent | 1.5x |
| `meeting_minutes` | 会议纪要 | 黑体 三号/left | 仿宋 小四/2char indent | 28pt fixed |
| `manual` | 用户手册 | 黑体 三号/left(H1) | 宋体 五号/left | 1.15x |
| `letter` | 公函 | 仿宋 三号/left | 仿宋 三号/justify | 28pt fixed |
| `proposal` | 项目建议书 | 黑体 二号/center | 宋体 小四/justify | 1.5x |
| `notice` | 通知 | 仿宋 二号/center | 仿宋 三号/2char indent | 28pt fixed |
| `work_report` | 工作总结 | 黑体 二号/center | 仿宋 三号/justify | 28pt fixed |

---

## Chinese Font Size Reference

| Don't Say | Say | pt |
|-----------|-----|----|
| 初号 / 小初 | 42 / 36 | 42 / 36 |
| 一号 / 小一 | 26 / 24 | 26 / 24 |
| 二号 | 22 | 22 |
| 小二 / 三号 | 18 / 16 | 18 / 16 |
| 小三 / 四号 | 15 / 14 | 15 / 14 |
| 小四 / 五号 | 12 / 10.5 | 12 / 10.5 |
| 小五 | 9 | 9 |

**Indent rule**: 2 Chinese chars ≈ font_size × 2 (e.g., 12pt body → 24pt first_line_indent)

---

## GB/T 9704-2012 Government Document Detail

### Page
| Property | Value |
|----------|-------|
| Paper | A4 (210×297mm) |
| Top | 37mm ±1mm |
| Bottom | 35mm ±1mm |
| Left | 28mm ±1mm |
| Right | 26mm ±1mm |

### Numbering
| Level | Format |
|-------|--------|
| 1 | 一、二、三... |
| 2 | (一)(二)(三)... |
| 3 | 1. 2. 3. |
| 4 | (1)(2)(3)... |

---

## Thesis Standard Detail

### Page
| Property | Value (pt) |
|----------|-----------|
| Top/Bottom | 72pt (2.54cm) |
| Left | 90pt (3.17cm) |
| Right | 72pt (2.54cm) |

### Title Hierarchy
| Level | Font | Size | Style |
|-------|------|------|-------|
| Chapter (H1) | 黑体 | 16pt (三号) | Bold, Center |
| Section (H2) | 黑体 | 14pt (四号) | Bold, Left |
| Subsection (H3) | 黑体 | 12pt (小四) | Bold, Left |
| Body | 宋体 | 12pt (小四) | Justify, 2char indent |
| Caption | 宋体 | 10.5pt (五号) | Center |
| Footnote | 宋体 | 9pt (小五) | Left |

---

## Surgical Editing Patterns (v2.0)

### Pattern 1: Batch heading format unification
```
surgical.select(para_indices=[2,14,27])  // all H2 headings
surgical.modify(mutations=[
  {para: 2, run:1, font_name: "黑体", size: 14, bold: true},
  {para: 14, run:1, font_name: "黑体", size: 14, bold: true},
  {para: 27, run:1, font_name: "黑体", size: 14, bold: true},
])
surgical.commit()  // execute + verify in one call
```

### Pattern 2: Semantic role targeting
```
content.query_by_role(sr="abstract") → locate abstract paragraphs
surgical.select(para_indices=[4])     // target them precisely
surgical.modify(...)
surgical.commit()
```

---

## Layout Fix Patterns (v2.0)

### Auto-fix all issues
```
layout.auto_fix_layout(filepath="doc.docx")
→ {fixed: 3, fixes: [orphan_heading, text_overflow, ...]}
```

### Quick widow/orphan fix (COM)
```
layout.fix_widow_orphan()
→ {fixed_paragraphs: 39, total_paragraphs: 39}
```

---

## Offline vs COM Mode Decision

| Scenario | Use | Note |
|----------|-----|------|
| Live WPS open with doc | COM mode (default) | Real-time, all format tools available |
| Batch format/analyze | Offline mode | Faster, no WPS file lock issues |
| Build from scratch | Offline build | No WPS needed |
| Edit COM-opened file | COM mode only | File locked by WPS |
| Edit closed file | Offline mode | `offline_docx replace_text` |

Offline requires: file NOT opened in WPS (avoid `[WinError 32]`)
