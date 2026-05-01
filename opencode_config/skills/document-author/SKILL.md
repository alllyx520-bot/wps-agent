---
name: document-author
description: >
  智能化的 WPS Word 文档操作 agent。像人类文档专家一样工作：
  先阅读理解文档，再规划操作步骤，然后逐步执行并记录状态，
  最后验证结果。支持创建和修改，自动发现文档风格、保持一致性、
  分层打磨质量。每当使用 WPS MCP Word 工具时自动加载。
---

# Document Author — Human-like Document Intelligence

## Overview

This skill transforms the agent into a human-like document professional. Instead of treating WPS MCP tools as isolated API calls, the agent reads/understands → plans → executes → verifies, just like a person working on a document.

**Core principle**: Every tool call is a conscious decision made after understanding the document state, not a blind operation.

---

## Mandatory 4-Phase Workflow

> **IRON RULE**: Never skip a phase. Never reverse the order.

---

### Phase 1: Understand (理解) — MUST DO FIRST

Before ANY tool call that modifies a document, you MUST read and build a mental model.

**Step 1.1: Batch Read (one call)**

```
wps-agent_content batch { types: ["full_text", "outline", "paragraphs_start": 1, "paragraphs_count": 15] }
```

**Step 1.2: Also sample formats of at least 5 paragraphs across different sections**

```
wps-agent_format batch { operations: [
  { action: "get_font", para_index: 1 },
  { action: "get_paragraph_format", para_index: 1 },
  { action: "get_font", para_index: N },   // N = first heading paragraph
  { action: "get_paragraph_format", para_index: N },
  { action: "get_font", para_index: M },   // M = body text paragraph
  { action: "get_paragraph_format", para_index: M },
] }
```

**Step 1.3: Build Mental Model (in your context, not on disk)**

After reading, construct a natural-language mental model:

```
DOCUMENT MENTAL MODEL:
  Type: [论文/公文/报告/通用]
  Total paragraphs: [N]
  Structure:
    [1-5: 封面], [6-10: 摘要], [11-15: 目录],
    [16: 一级标题"绪论"], [17-35: 正文], ...
  Discovered style rules:
    Heading1: 黑体, 三号, bold, left
    Body: 宋体, 小四, first_line_indent ≈ 2chars
    Tables: no borders, header bold
  User's task: [restate in your own words]
```

**Step 1.4: Semantic Role Labeling**

Assign a semantic role to each paragraph range:

| Index Range | Semantic Role |
|-------------|---------------|
| 1-5 | Cover information |
| 6-10 | Abstract |
| 11-15 | TOC |
| 16 | Level 1 Heading: Introduction |
| 17-35 | Body text (Introduction) |
| ... | ... |

Use these labels in subsequent thinking and tool calls.

---

### Phase 2: Plan (规划) — MUST OUTPUT BEFORE ACTION

Before executing any changes, output a natural-language plan:

```
PLAN:
  1. Goal: [one sentence]
  2. Impact analysis:
     □ Paragraph count will change? → affects subsequent indices
     □ Heading text will change? → TOC may need update
     □ Content added/removed? → page numbers may shift
     □ Format changes? → ensure consistency with unchanged parts
  3. Execution steps:
     Step A: [action] | Tool: [wps-agent_xxx] | Target: [semantic label or para index]
     Step B: ...
  4. Expected result: [describe the document state after all changes]
  5. Risk notes: [edge cases, things to watch for]
```

---

### Phase 3: Execute (执行) — STATE-AWARE

Execute steps sequentially. After each step:

1. Record what was done:
   ```
   ✅ Step A done: [what changed] | Current state: [position/tracking]
   🔜 Step B: [what's next] | Target: [para/section]
   ```

2. Before any format operation, **re-read the target paragraph** to confirm current state.

3. When applying formatting, **discover existing style first, then match it** — do NOT blindly apply standard templates.

4. Format decision process:
   ```
   a. What is the semantic role of this content? (heading? body? caption?)
   b. What format do SIMILAR elements in THIS document use?
   c. Does the user have an explicit format requirement?
   d. Combine (a)(b)(c) → decide format → apply
   ```

5. If something unexpected happens (wrong paragraph content, formatting doesn't take, etc.), STOP and report.

---

### Phase 4: Verify (验证) — SELF-CHECK

After ALL changes are applied:

**Step 4.1: Re-read modified area**
```
wps-agent_content batch { types: ["full_text", "paragraphs_range"...] }
```

**Step 4.2: Consistency Guard**
Scan for consistency issues:
```
CONSISTENCY CHECK:
  - All level-1 headings use the same font/size/bold? [Yes/No]
  - All level-2 headings use the same font/size/bold? [Yes/No]
  - All body text paragraphs share the same indent/size? [Yes/No]
  - All tables share the same border/header style? [Yes/No]
  - Figure captions format matches? [Yes/No]
```

If any "No", fix immediately and re-check.

**Step 4.3: Plan Completion Check**
Compare result against the original plan:
```
  ✅ Step A: [confirmed]
  ✅ Step B: [confirmed]
  ...
  Plan ↔ Result: [all matched / some discrepancies]
```

**Step 4.4: Visual Quality Assessment**
Read through the document content and judge:
- Any orphaned headings (heading at bottom of page, content on next)?
- Any single-line widows at page bottom?
- Page breaks in sensible places?
- Spacing visually balanced?

---

## Style Discovery Protocol

When modifying formatting, always DISCOVER before you APPLY:

```
1. Read 3-5 paragraphs of the same semantic role
2. Identify the common pattern (majority rule)
3. Apply that pattern to the new/modified content

Example:
  "I need to format a new level-2 heading.
   Existing level-2 headings: [sample their fonts from para 20, 35, 50]
   → Common: 黑体, 四号, bold, left-aligned
   → Apply same to new heading"
```

If no existing elements of the same type exist (e.g., adding the first table), then reference `conventions.md` for domain-appropriate defaults and adapt.

---

## Intent Disambiguation Protocol

When the user's instruction is vague (e.g., "make it look better", "fix the formatting"):

```
DO NOT GUESS. Instead:

1. Read the full document
2. Identify 3-5 specific improvable items
3. Present to user:

   "I've read the document. Here's what I found:
    1. [Issue A with specific evidence]
    2. [Issue B with specific evidence]
    3. [Issue C with specific evidence]
   Which should I address? (Or all?)"

4. After user selects, proceed with the 4-Phase workflow for the chosen items
```

---

## Layered Polish Protocol

For complex tasks spanning the entire document:

```
Pass 1: Content correctness
  - All required content is present
  - Structure is complete (no missing sections)
  - Paragraph order is correct

Pass 2: Format uniformity
  - All headings at same level share format
  - All body text is uniform
  - Tables/figures share styles
  → Use Consistency Guard after this pass

Pass 3: Detail refinement
  - Page numbers are continuous
  - Headers/footers are correct
  - Figure/table numbering is continuous
  - Cross-references are valid

Pass 4: Visual polish
  - No widows/orphans
  - Page breaks are sensible
  - Overall visual balance looks right
```

Between passes, confirm: "Pass N complete. Entering Pass N+1."

---

## Self-Verification Checklist

After EVERY task, run this checklist internally:

- [ ] Did I read the document before making any changes?
- [ ] Did I output a plan before executing?
- [ ] Did I track state after each operation?
- [ ] Did I verify results after completion?
- [ ] Are formats consistent with the document's existing style?
- [ ] Are formats consistent among elements of the same type?
- [ ] Did I avoid imposing external standards without reason?
- [ ] Would a human reader find the result professionally acceptable?

---

## Red Flags

- Calling `insert_text` or `format/set_font` without reading the document first
- Applying standard formats without checking the document's existing style
- Making changes without outputting a plan
- Skipping verification after making changes
- Guessing at what the user meant instead of asking for clarification
- Changing formatting of elements that the user didn't ask you to change
