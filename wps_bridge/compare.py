# -*- coding: utf-8 -*-
"""
Document comparison: text diff and format diff between two Word documents.
"""
import difflib
from typing import Any, Dict, List
from .app import get_doc
from .content import paragraphs, outline
from .formatting import get_font, get_paragraph_format
from .docspace import get_word_doc_by_id
from .utils import com_property


def _doc_by_id(doc_id: str):
    doc_id = doc_id.strip()
    if ":" not in doc_id:
        return get_doc(int(doc_id))
    return get_word_doc_by_id(doc_id)


def text_diff(doc_id_a: str, doc_id_b: str) -> Dict:
    doc_a = _doc_by_id(doc_id_a)
    doc_b = _doc_by_id(doc_id_b)
    text_a = com_property(doc_a.Content, "Text", "")
    text_b = com_property(doc_b.Content, "Text", "")
    lines_a = [l.strip() for l in text_a.split("\r") if l.strip()]
    lines_b = [l.strip() for l in text_b.split("\r") if l.strip()]
    differ = difflib.unified_diff(lines_a, lines_b, lineterm="",
                                   fromfile=f"doc_a({doc_id_a})", tofile=f"doc_b({doc_id_b})")
    diff_lines = list(differ)
    changes = {
        "added": sum(1 for l in diff_lines if l.startswith("+") and not l.startswith("+++")),
        "removed": sum(1 for l in diff_lines if l.startswith("-") and not l.startswith("---")),
    }
    return {
        "doc_a": {"id": doc_id_a, "name": com_property(doc_a, "Name", ""), "paragraphs": len(lines_a)},
        "doc_b": {"id": doc_id_b, "name": com_property(doc_b, "Name", ""), "paragraphs": len(lines_b)},
        "changes": changes,
        "diff": diff_lines[:100] if len(diff_lines) > 100 else diff_lines,
        "diff_truncated": len(diff_lines) > 100,
    }


def format_diff(doc_id_a: str, doc_id_b: str) -> Dict:
    doc_a = _doc_by_id(doc_id_a)
    doc_b = _doc_by_id(doc_id_b)
    count_a = doc_a.Paragraphs.Count
    count_b = doc_b.Paragraphs.Count
    format_changes = []
    min_count = min(count_a, count_b)
    for i in range(1, min_count + 1):
        try:
            fa = doc_a.Paragraphs.Item(i).Range.Font
            fb = doc_b.Paragraphs.Item(i).Range.Font
            diffs = []
            if com_property(fa, "Name", "") != com_property(fb, "Name", ""):
                diffs.append(f"font: {com_property(fa,'Name','')} -> {com_property(fb,'Name','')}")
            if abs(com_property(fa, "Size", 0) - com_property(fb, "Size", 0)) > 0.1:
                diffs.append(f"size: {com_property(fa,'Size',0)} -> {com_property(fb,'Size',0)}")
            if com_property(fa, "Bold", 0) != com_property(fb, "Bold", 0):
                diffs.append(f"bold: {com_property(fa,'Bold',0)} -> {com_property(fb,'Bold',0)}")
            if diffs:
                format_changes.append({"para_index": i, "text": com_property(doc_a.Paragraphs.Item(i).Range, "Text", "")[:50], "diffs": diffs})
        except Exception:
            continue
    extra_info = ""
    if count_a > count_b:
        extra_info = f"doc_a has {count_a - count_b} more paragraphs"
    elif count_b > count_a:
        extra_info = f"doc_b has {count_b - count_a} more paragraphs"
    return {
        "doc_a": {"paragraphs": count_a},
        "doc_b": {"paragraphs": count_b},
        "format_changes": format_changes[:50],
        "changed_count": len(format_changes),
        "extra": extra_info,
    }
