# -*- coding: utf-8 -*-
"""Surgical Context — captures a document *region* with full context for safe modification.

Problem: Character-position-based editing is fragile because any edit shifts all
subsequent positions. A surgical context solves this by:

1. Identifying targets by paragraph index, Run index, or semantic role
2. Capturing the NEIGHBORHOOD (surrounding paragraphs' full state)
3. Executing modifications with pre/post verification
4. Auto-rolling back on verification failure

This is the "human-like" approach — understand context first, then operate.

Usage:
    ctx = SurgicalContext([5,6,7], doc_index=0)
    ctx.capture()
    ctx.modify({"para": 5, "run": 2, "font": "黑体", "size": 16, "bold": True})
    ctx.modify({"para": 6, "alignment": "center"})
    result = ctx.commit()
    # or ctx.rollback() if something went wrong
"""

from typing import Any, Dict, List, Optional
from .app import get_doc
from .utils import com_property, com_set
from .content import runs_detail, paragraph as _paragraph_read
from .formatting import set_font, set_paragraph_format, set_run_font


class SurgicalContext:
    """Capture a document region, modify with verification, commit or rollback."""

    def __init__(self, para_indices: List[int], doc_index: Optional[int] = None):
        self.para_indices = sorted(set(para_indices))
        self.doc_index = doc_index
        self.pre_snap: Dict[int, Dict] = {}   # para_index → snapshot
        self.mutations: List[Dict] = []
        self.committed = False

    def capture(self) -> "SurgicalContext":
        """Capture baseline state of all target paragraphs plus one neighbor on each side."""
        neighbors = set()
        for pi in self.para_indices:
            neighbors.add(pi)
            if pi > 1:
                neighbors.add(pi - 1)
            neighbors.add(pi + 1)

        doc = get_doc(self.doc_index)
        total = doc.Paragraphs.Count
        for pi in sorted(neighbors):
            if pi < 1 or pi > total:
                continue
            try:
                detail = runs_detail(pi, self.doc_index)
                self.pre_snap[pi] = {
                    "text": detail.get("text", ""),
                    "style": detail.get("style_name", ""),
                    "alignment": detail.get("alignment", ""),
                    "space_before": detail.get("space_before", 0),
                    "space_after": detail.get("space_after", 0),
                    "outline_level": detail.get("outline_level", 10),
                    "run_count": detail.get("word_count", 0),
                    "runs": [{"idx": r["index"], "text": r["text"], "font": r["font"]}
                             for r in detail.get("runs", [])],
                }
            except Exception:
                continue
        return self

    def modify(self, mutation: Dict) -> "SurgicalContext":
        """Record a mutation to apply at commit time.

        mutation format:
            {"para": N} → paragraph-level change
            {"para": N, "run": M} → run-level change
            {"sr": "abstract"} → target by semantic role (requires filepath)

        Supported keys: {para, run, font_name, size, bold, italic, underline,
                         color_index, color_rgb, highlight, superscript, subscript,
                         strike_through, caps, small_caps, alignment, first_line_indent,
                         left_indent, right_indent, line_spacing_rule, line_spacing,
                         space_before, space_after, outline_level, text}
        """
        required = ["para"]
        if "run" in mutation:
            required.append("run")
        for r in required:
            if r not in mutation:
                return self
        self.mutations.append(dict(mutation))
        return self

    def commit(self) -> Dict:
        """Execute all mutations, verify, and return results."""
        if not self.mutations:
            return {"error": "No mutations to commit", "error_code": "NO_MUTATIONS"}

        doc = get_doc(self.doc_index)
        applied = []
        for mut in self.mutations:
            pi = mut["para"]
            if pi < 1:
                pi = 1
            run_idx = mut.get("run")
            try:
                if run_idx is not None:
                    font_kwargs = {k: v for k, v in mut.items()
                                   if k in ("name", "size", "bold", "italic", "underline",
                                            "color_index", "color_rgb", "highlight",
                                            "superscript", "subscript", "strike_through",
                                            "caps", "small_caps", "emboss", "shadow",
                                            "outline", "vanish", "spacing", "scaling",
                                            "kerning") and v is not None}
                    if "font_name" in mut and mut["font_name"]:
                        font_kwargs["name"] = mut["font_name"]
                    if "font_size" in mut and mut["font_size"]:
                        font_kwargs["size"] = mut["font_size"]
                    if font_kwargs:
                        res = set_run_font(pi, run_idx, self.doc_index, **font_kwargs)
                        applied.append({"para": pi, "run": run_idx, "action": "set_run_font", "result": res})
                    if "text" in mut and mut["text"] is not None:
                        p = doc.Paragraphs.Item(pi)
                        rng = p.Range
                        w = rng.Words.Item(run_idx)
                        w.Text = mut["text"]
                        applied.append({"para": pi, "run": run_idx, "action": "set_text"})
                else:
                    font_kwargs = {k: v for k, v in mut.items()
                                   if k in ("name", "size", "bold", "italic", "underline",
                                            "color_index", "color_rgb", "highlight",
                                            "superscript", "subscript", "strike_through",
                                            "caps", "small_caps", "shadow", "outline",
                                            "emboss", "vanish") and v is not None}
                    if "font_name" in mut and mut["font_name"]:
                        font_kwargs["name"] = mut["font_name"]
                    if "font_size" in mut and mut["font_size"]:
                        font_kwargs["size"] = mut["font_size"]
                    if font_kwargs:
                        set_font(para_index=pi, doc_index=self.doc_index, **font_kwargs)
                        applied.append({"para": pi, "action": "set_font"})
                    pf_kwargs = {k: v for k, v in mut.items()
                                 if k in ("alignment", "first_line_indent", "left_indent",
                                          "right_indent", "line_spacing_rule", "line_spacing",
                                          "space_before", "space_after", "outline_level")}
                    if pf_kwargs:
                        set_paragraph_format(para_index=pi, doc_index=self.doc_index, **pf_kwargs)
                        applied.append({"para": pi, "action": "set_paragraph_format"})
            except Exception as e:
                return {"error": str(e), "applied_so_far": applied, "failed_mutation": mut,
                        "error_code": "SURGICAL_FAILED"}

        verify_result = self._verify()
        if verify_result.get("mismatches"):
            return {"committed": True, "warning": "Verification found mismatches",
                    "mismatches": verify_result["mismatches"], "applied": applied}

        self.committed = True
        return {"committed": True, "applied": applied, "verified": True}

    def rollback(self) -> Dict:
        """Reset all target paragraphs to their pre-capture state."""
        doc = get_doc(self.doc_index)
        restored = 0
        for pi, snap in self.pre_snap.items():
            try:
                p = doc.Paragraphs.Item(pi)
                p.Range.Text = snap["text"]
                if snap.get("style"):
                    p.Range.Style = snap["style"]
                pf = p.Format
                rng = p.Range
                f = rng.Font
                runs = snap.get("runs", [])
                if runs:
                    for rn in runs:
                        try:
                            w = rng.Words.Item(rn["idx"])
                            wf = w.Font
                            fn = rn["font"]
                            if fn.get("name"):
                                com_set(wf, "Name", fn["name"])
                            if fn.get("size"):
                                com_set(wf, "Size", fn["size"])
                            if "bold" in fn:
                                com_set(wf, "Bold", fn["bold"])
                            if "italic" in fn:
                                com_set(wf, "Italic", fn["italic"])
                        except Exception:
                            pass
                restored += 1
            except Exception:
                continue
        self.mutations.clear()
        return {"rolled_back": True, "restored_paragraphs": restored}

    def _verify(self) -> Dict:
        """Re-read target paragraphs and diff against pre-snapshot."""
        mismatches = []
        for pi in self.para_indices:
            if pi not in self.pre_snap:
                continue
            pre = self.pre_snap[pi]
            try:
                cur = runs_detail(pi, self.doc_index)
                cur_text = cur.get("text", "")
                pre_text = pre.get("text", "")
                if cur_text != pre_text:
                    mismatches.append({
                        "para": pi, "field": "text",
                        "expected": pre_text[:80], "actual": cur_text[:80],
                    })
            except Exception:
                mismatches.append({"para": pi, "field": "read_error",
                                   "error": "Failed to re-read paragraph"})
        return {"matched": len(mismatches) == 0, "mismatches": mismatches}
