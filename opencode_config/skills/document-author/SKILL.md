---
name: document-author
description: >
  智能化的 WPS Word 文档操作 agent。像人类文档专家一样工作：
  先阅读理解文档（含语义解析+排版分析），再规划操作步骤，
  然后逐步执行并记录状态，最后验证结果（含自动排版修正）。
  支持 18 个 Tool 的完整编排，包括 v2.0 新增的 surgical 手术级修改、
  query_by_role 语义定位、auto_fix_layout 排版自动修正。
  每当使用 WPS MCP Word 工具时自动加载。
---

# Document Author — v2.0 Human-like Document Intelligence

## Overview

This skill transforms the agent into a human-like document professional with full awareness of the v2.0 wps-agent tool suite. Instead of treating WPS MCP tools as isolated API calls, the agent reads/understands → plans (with semantic awareness) → executes (with surgical precision) → verifies (with layout auto-fix), just like a person working on a document.

**Core principle**: Every tool call is a conscious decision made after understanding the document state, not a blind operation.

---

## Mandatory 4-Phase Workflow

> **IRON RULE**: Never skip a phase. Never reverse the order.

---

### Phase 1: Understand (理解) — MUST DO FIRST

Before ANY modifying tool call, build a complete mental model of the document.

**Step 1.1: Structural Read (3 parallel calls)**

```json
{"tool": "wps-agent_content", "action": "full_text"}
{"tool": "wps-agent_content", "action": "document_structure"}
{"tool": "wps-agent_content", "action": "outline"}
```

**Step 1.2: Semantic Deep Read (v2.0)**

```json
{"tool": "wps-agent_content", "action": "semantic_structure"}
{"tool": "wps-agent_content", "action": "query_by_role", "sr": "abstract"}
{"tool": "wps-agent_content", "action": "query_by_role", "sr": "cover"}
```

Purpose: Auto-identify 20+ semantic roles (cover/abstract/keywords/toc/chapter/section/body/references/acknowledgements/appendix), build DocumentGraph relationships.

**Step 1.3: Format Sampling**

```json
{"tool": "wps-agent_format", "action": "batch", "operations": [
  {"type": "get_font", "para_index": 1},
  {"type": "get_paragraph_format", "para_index": 1},
  {"type": "get_font", "para_index": FIRST_HEADING},
  {"type": "get_paragraph_format", "para_index": FIRST_HEADING},
  {"type": "get_font", "para_index": FIRST_BODY},
  {"type": "get_paragraph_format", "para_index": FIRST_BODY}
]}
```

**Step 1.4: Build Mental Model**

```
DOCUMENT MENTAL MODEL:
  Type: [thesis/report/official/resume/contract/...] (12 types)
  Total paragraphs: [N] | Tables: [M] | Sections: [S]
  Semantic Structure:
    [1-3: cover], [4: abstract], [5-8: toc],
    [9: chapter_1 "绪论"], [10-45: body], ... (from semantic_structure)
  Discovered style rules:
    Chapter Title: 黑体/22pt/bold/center
    Section Title: 黑体/16pt/bold/left
    Body: 宋体/12pt/justify/first_line_indent=24pt
    Table: header bold, no borders
  Content Types (from classify_paragraph_content):
    Expository: 80% | Data: 10% | Reference: 10%
  User's task: [restate in own words]
```

**Step 1.5: Layout Analysis (v2.0)**

```json
{"tool": "wps-agent_layout", "action": "auto_fix_layout"}
```

Detect issues BEFORE planning: orphan headings, text overflow, table page-break risks, column imbalance.

---

### Phase 2: Plan (规划) — MUST OUTPUT BEFORE ACTION

> **IRON RULE for NEW documents**: When creating a new document from scratch (not modifying an existing one), you MUST use `content.action=build` with a complete `structure` JSON. This single call replaces 50-70 individual `insert_paragraph`/`insert_text`/`format`/`table create` calls, and eliminates paragraph merging, index shifting, table boundary bugs, and silent failures.

**Decision gate — Is this a NEW document?**

| Scenario | Primary Tool |
|----------|-------------|
| Create new document from scratch | `content.action=build` (single call) |
| Modify existing document | Phase 2 normal workflow (surgical/format/insert) |
| Add 1-2 paragraphs to existing doc | `insert_paragraph` / `surgical` |
| Create cover page on existing doc | `content.action=create_cover` |

**Build JSON template:**

```json
{
  "tool": "wps-agent_content",
  "action": "build",
  "structure": {
    "cover": {"lines": [
      {"text": "Title", "font_name": "黑体", "font_size": 26, "bold": true, "alignment": "center", "space_before": 120},
      {"text": "Subtitle", "font_name": "宋体", "font_size": 16, "alignment": "center"}
    ]},
    "sections": [
      {"heading": "Section Heading", "paragraphs": ["Body text 1", "Body text 2"]},
      {"heading": "Section With Table",
       "table": {"headers": ["Col1","Col2"], "rows": [["a","b"],["c","d"]], "header_bold": true}}
    ],
    "defaults": {"body_font": "宋体", "body_size": 12, "heading_font": "黑体", "heading_size": 16, "first_line_indent": 24},
    "page_setup": {"page_width": 595.3, "page_height": 841.9, "top_margin": 72, "bottom_margin": 72, "left_margin": 90, "right_margin": 90}
  },
  "output_path": "E:\\path\\to\\doc.docx"
}
```

Output a natural-language plan with precise tool references.

```
PLAN:
  1. Goal: [one sentence]
  2. Impact analysis:
     □ Paragraph count change? → affects subsequent indices
     □ Heading text change? → TOC may need update
     □ Content added/removed? → page numbers shift
     □ Format changes? → consistency check needed
     □ Layout issues detected in Phase 1.5? → include fix steps
  3. Execution steps (with precise tool mapping):
     Step A: [action] | Tool: [wps-agent_xxx] | Action: [yyy] | Target: [semantic role or para index]
     Step B: [action] | Tool: [wps-agent_surgical] | Action: [select/modify/commit] | Target: [...]
     Step C: ...
  4. For precise edits, PREFER surgical tool:
     surgical.select(para_indices=[...]) → modify(mutations=[...]) → commit()
  5. For semantic targeting, USE query_by_role:
     content.query_by_role(sr="abstract") → locates abstract paragraphs
  6. Expected result: [document state after changes]
  7. Verification: layout.auto_fix_layout + consistency guard
```

---

### Phase 3: Execute (执行) — STATE-AWARE + SURGICAL

**Step 3.0: For new document creation — USE build (one call)**

Execute the `content.build` plan from Phase 2. No individual insert/format calls needed. Verify the result with `document_structure`.

**Step 3.1: For simple single-paragraph changes**

Use direct format calls:
```json
{"tool": "wps-agent_format", "action": "set_font", "para_index": 1, "name": "黑体", "size": 26, "bold": true}
{"tool": "wps-agent_format", "action": "set_paragraph_format", "para_index": 1, "alignment": "center"}
```

**Step 3.2: For multi-element coordinated changes — USE SURGICAL**

```json
// Step A: Capture context
{"tool": "wps-agent_surgical", "action": "select", "para_indices": [2, 3, 4, 5]}
→ {session_id: "...", context: {...}}

// Step B: Queue mutations
{"tool": "wps-agent_surgical", "action": "modify", "session_id": "...",
 "mutations": [
   {"para": 2, "run": 1, "font_name": "黑体", "size": 16, "bold": true},
   {"para": 2, "alignment": "center", "space_before": 24},
   {"para": 3, "first_line_indent": 24, "space_after": 6}
 ]}

// Step C: Commit with auto-verification
{"tool": "wps-agent_surgical", "action": "commit", "session_id": "..."}
→ {committed: true, verified: true}
```

**Step 3.3: For text effects — USE set_text_effect**

```json
{"tool": "wps-agent_format", "action": "set_text_effect", "para_index": 1, "effect": "shadow"}
// effects: shadow, outline, emboss, engrave, glow, reflection
```

**During execution:**
1. State-tracking after each tool call: ✅ Step A done / 🔜 Step B next
2. Before format operations, re-read target paragraph to confirm state
3. Format decisions: discover existing style first → match it — NOT blindly apply templates
4. If unexpected → STOP and report

---

### Phase 4: Verify (验证) — SELF-CHECK + AUTO-FIX

**Step 4.1: Re-read modified area**

```json
{"tool": "wps-agent_content", "action": "full_text"}
{"tool": "wps-agent_content", "action": "document_structure"}
```

**Step 4.2: Consistency Guard**

```
CONSISTENCY CHECK:
  - All H1 same font/size/bold? [Yes/No]
  - All H2 same font/size/bold? [Yes/No]
  - All body text same indent/size? [Yes/No]
  - All tables same border/header? [Yes/No]
  - Captions format matches? [Yes/No]
```
If any "No" → fix immediately and re-check.

**Step 4.3: Layout Auto-Fix (v2.0)**

```json
{"tool": "wps-agent_layout", "action": "auto_fix_layout"}
```

Automatically: fix orphan headings (increase space_after), fix text overflow (adjust line_spacing), detect table page-break risks, detect column imbalance.

**Step 4.4: Plan Completion Check**

```
  ✅ Step A: [confirmed]
  ✅ Step B: [confirmed]
  ...
  Plan ↔ Result: [matched / discrepancies: ...]
```

---

## New Tool Quick Reference (v2.0)

### surgical — Context-aware precise editing
```
select(para_indices=[2,3]) → modify(mutations=[...]) → commit() / rollback()
```
Use when: batch editing multiple paragraphs, precise formatting across sections, need rollback safety.

### query_by_role — Semantic paragraph targeting
```
content.query_by_role(sr="abstract") → [{index: 4, role: "abstract_label", ...}]
```
Roles: abstract, cover, keywords, toc, references, acknowledgements, appendix

### auto_fix_layout — Layout quality auto-fix
```
layout.auto_fix_layout(filepath="doc.docx") → {fixed: N, issues_found: M}
```
Detects: orphan headings, text overflow, table page-break, column imbalance

### fix_widow_orphan — Widow/orphan quick fix
```
layout.fix_widow_orphan() → {fixed_paragraphs: N}
```
Sets WidowControl=True on all paragraphs via COM.

### set_text_effect — Text visual effects
```
format.set_text_effect(para_index=1, effect="glow", color_rgb=255, offset=2)
```
Effects: shadow, outline, emboss, engrave, glow, reflection

---

## Style Discovery Protocol

When modifying formatting, DISCOVER before APPLY:
1. Read 3-5 paragraphs of same semantic role
2. Identify common pattern (majority rule)
3. Apply that pattern to new content

If no existing elements → reference `references/conventions.md` for 14-template defaults.

---

## Intent Disambiguation Protocol

Vague request ("make it better") → Don't guess:
1. Read full document
2. Identify 3-5 specific improvable items
3. Present to user: "I found: [A/B/C]. Which should I address?"
4. After selection → 4-Phase workflow

---

## Advanced: Offline Mode Routing

When `wps-agent` MCP is unavailable:
1. Document analysis: `offline_docx analyze/validate`
2. Automatic formatting: `offline_docx auto_format` (report/thesis/resume/general)
3. Template application: `offline_docx apply_template` (thesis_cn/report_official/resume_professional)
4. Document building: `offline_docx build` (from JSON structure)
5. Text replacement: `offline_docx replace_text`
6. Fallback: `docx` skill (npm docx + pandoc)

Note: Offline mode write operations require file NOT locked by WPS COM.

---

## Self-Verification Checklist

- [ ] Did I read the document before changes?
- [ ] Did I run semantic_structure + query_by_role?
- [ ] Did I check layout with auto_fix_layout?
- [ ] Did I output a plan before executing?
- [ ] Did I consider using surgical for multi-element changes?
- [ ] Did I track state after each operation?
- [ ] Did I verify results after completion?
- [ ] Are formats consistent within same element type?
- [ ] Did I avoid imposing external standards?
- [ ] Would a human reader find the result professionally acceptable?

---

## Red Flags

- Modifying without reading document structure first
- Applying standard formats without style discovery
- Making changes without outputting a plan
- NOT using surgical for batch/multi-paragraph edits
- NOT checking query_by_role for intelligent targeting
- Skipping auto_fix_layout verification
- Guessing user intent instead of disambiguation protocol
