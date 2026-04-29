# -*- coding: utf-8 -*-
"""
Cross-document content transfer: copy paragraphs, tables, ranges between Word docs.
Uses doc_id format (e.g. 'word:1', 'word:2') via docspace module.
"""
from typing import Any, Dict, Optional
from .app import get_app, get_doc
from .docspace import get_word_doc_by_id, _resolve
from .utils import com_property


def _doc_by_id(doc_id: str):
    doc_id = doc_id.strip()
    if ":" not in doc_id:
        # Fallback: bare index treated as word doc index
        return get_doc(int(doc_id))
    return get_word_doc_by_id(doc_id)


def copy_paragraphs(source_doc_id: str, from_start: int, from_end: int,
                    target_doc_id: str, target_position: str = "end") -> Dict:
    src_doc = _doc_by_id(source_doc_id)
    tgt_doc = _doc_by_id(target_doc_id)
    total = src_doc.Paragraphs.Count
    from_start = max(1, from_start)
    from_end = min(from_end, total)
    # Select and copy source range (preserves formatting)
    src_rng = src_doc.Range(
        src_doc.Paragraphs.Item(from_start).Range.Start,
        src_doc.Paragraphs.Item(from_end).Range.End
    )
    src_rng.Copy()
    # Paste at target position
    if target_position == "end":
        tgt_rng = tgt_doc.Range(tgt_doc.Content.End - 1, tgt_doc.Content.End - 1)
    elif isinstance(target_position, int):
        tgt_rng = tgt_doc.Paragraphs.Item(target_position).Range
    else:
        tgt_rng = tgt_doc.Range(tgt_doc.Content.End - 1, tgt_doc.Content.End - 1)
    tgt_rng.Paste()
    return {"action": "copy_paragraphs", "source_range": f"{from_start}-{from_end}", "copied": from_end - from_start + 1}


def copy_table(source_doc_id: str, table_index: int,
               target_doc_id: str, target_position: str = "end") -> Dict:
    src_doc = _doc_by_id(source_doc_id)
    tgt_doc = _doc_by_id(target_doc_id)
    src_table = src_doc.Tables.Item(table_index)
    rows = src_table.Rows.Count
    cols = src_table.Columns.Count
    # Read source table
    data = []
    for r in range(1, rows + 1):
        row_data = []
        for c in range(1, cols + 1):
            try:
                cell_text = src_table.Cell(r, c).Range.Text.replace("\r\x07", "").strip()
                row_data.append(cell_text)
            except Exception:
                row_data.append("")
        data.append(row_data)
    # Determine insert position
    if target_position == "end":
        rng = tgt_doc.Range(tgt_doc.Content.End - 1, tgt_doc.Content.End - 1)
    elif isinstance(target_position, int):
        rng = tgt_doc.Paragraphs.Item(target_position).Range
    else:
        rng = tgt_doc.Range(tgt_doc.Content.End - 1, tgt_doc.Content.End - 1)
    # Create table in target doc
    new_table = tgt_doc.Tables.Add(rng, rows, cols)
    new_table.AutoFitBehavior(2)
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            try:
                new_table.Cell(r, c).Range.Text = data[r - 1][c - 1]
            except Exception:
                continue
    return {"action": "copy_table", "source_table": table_index, "rows": rows, "cols": cols, "target_table_index": tgt_doc.Tables.Count}


def copy_range(source_doc_id: str, start_pos: int, end_pos: int,
               target_doc_id: str, target_position: str = "end") -> Dict:
    src_doc = _doc_by_id(source_doc_id)
    tgt_doc = _doc_by_id(target_doc_id)
    src_rng = src_doc.Range(start_pos, end_pos)
    text = com_property(src_rng, "Text", "")
    if target_position == "end":
        tgt_rng = tgt_doc.Range(tgt_doc.Content.End - 1, tgt_doc.Content.End - 1)
    elif isinstance(target_position, int):
        tgt_rng = tgt_doc.Paragraphs.Item(target_position).Range
    else:
        tgt_rng = tgt_doc.Range(tgt_doc.Content.End - 1, tgt_doc.Content.End - 1)
    tgt_rng.InsertAfter(text)
    return {"action": "copy_range", "source_range": f"{start_pos}-{end_pos}", "chars": len(text)}
