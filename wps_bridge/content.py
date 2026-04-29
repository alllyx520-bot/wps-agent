# -*- coding: utf-8 -*-
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set, WDALIGNMENT, WDLINESPACING


def full_text(doc_index: Optional[int] = None) -> str:
    return com_property(get_doc(doc_index).Content, "Text", "")


def paragraph(para_index: int, doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    p = doc.Paragraphs.Item(para_index)
    r = p.Range
    f = r.Font
    pf = p.Format
    return {
        "index": para_index,
        "text": com_property(r, "Text", "").strip(),
        "style_name": com_property(p.Range.Style, "NameLocal", ""),
        "outline_level": com_property(pf, "OutlineLevel", 10),
        "font": {
            "name": com_property(f, "Name", ""),
            "name_far_east": com_property(f, "NameFarEast", ""),
            "size": com_property(f, "Size", 0),
            "bold": bool(com_property(f, "Bold", 0)),
            "italic": bool(com_property(f, "Italic", 0)),
            "underline": com_property(f, "Underline", 0),
            "color_index": com_property(f, "ColorIndex", 0),
        },
        "paragraph_format": {
            "alignment": WDALIGNMENT.get(com_property(pf, "Alignment", 0), "unknown"),
            "first_line_indent": com_property(pf, "FirstLineIndent", 0),
            "left_indent": com_property(pf, "LeftIndent", 0),
            "right_indent": com_property(pf, "RightIndent", 0),
            "line_spacing_rule": WDLINESPACING.get(com_property(pf, "LineSpacingRule", 0), "unknown"),
            "line_spacing": com_property(pf, "LineSpacing", 0),
            "space_before": com_property(pf, "SpaceBefore", 0),
            "space_after": com_property(pf, "SpaceAfter", 0),
        },
    }


def paragraphs(start: int, count: int = 10, doc_index: Optional[int] = None) -> List[Dict]:
    result = []
    doc = get_doc(doc_index)
    total = doc.Paragraphs.Count
    end = min(start + count - 1, total)
    for i in range(start, end + 1):
        try:
            result.append(paragraph(i, doc_index))
        except Exception:
            continue
    return result


def selection_info(doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    sel = get_app().Selection
    r = sel.Range
    f = r.Font
    result = {
        "text": com_property(r, "Text", "").strip(),
        "start": com_property(r, "Start", 0),
        "end": com_property(r, "End", 0),
        "font": {"name": com_property(f, "Name", ""), "size": com_property(f, "Size", 0), "bold": bool(com_property(f, "Bold", 0)), "italic": bool(com_property(f, "Italic", 0))},
    }
    try:
        result["alignment"] = WDALIGNMENT.get(com_property(r.ParagraphFormat, "Alignment", 0), "unknown")
    except Exception:
        pass
    return result


def range_text(start_pos: int, end_pos: int, doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    r = doc.Range(start_pos, end_pos)
    f = r.Font
    result = {"text": com_property(r, "Text", ""), "start": start_pos, "end": end_pos, "font": {"name": com_property(f, "Name", ""), "size": com_property(f, "Size", 0), "bold": bool(com_property(f, "Bold", 0))}}
    try:
        result["alignment"] = WDALIGNMENT.get(com_property(r.ParagraphFormat, "Alignment", 0), "unknown")
    except Exception:
        pass
    return result


def batch(items: list, doc_index: Optional[int] = None) -> List[Dict]:
    results = []
    for item in items:
        item_type = item.get("type", item.get("action", ""))
        try:
            if item_type == "paragraph":
                data = paragraph(item["para_index"], doc_index)
            elif item_type == "paragraphs":
                data = {"items": paragraphs(item.get("start", 1), item.get("count", 10), doc_index)}
            elif item_type == "outline":
                data = outline(doc_index)
            elif item_type == "full_text":
                data = full_text(doc_index)
            elif item_type == "selection":
                data = selection_info(doc_index)
            else:
                results.append({"ok": False, "type": item_type, "error": f"Unknown read type: {item_type}"})
                continue
            results.append({"ok": True, "type": item_type, "data": data})
        except Exception as e:
            results.append({"ok": False, "type": item_type, "error": str(e)})
    return results


def outline(doc_index: Optional[int] = None) -> List[Dict]:
    result = []
    doc = get_doc(doc_index)
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs.Item(i)
        level = com_property(p.Format, "OutlineLevel", 10)
        if 1 <= level <= 9:
            result.append({"index": i, "text": com_property(p.Range, "Text", "").strip(), "outline_level": level, "style": com_property(p.Range.Style, "NameLocal", "")})
    return result


def _insert_at_position(doc, text_lines: List[str], insert_before: bool = False):
    """Insert multiple paragraphs at a location using reliable paragraph creation."""
    app = get_app()
    sel = app.Selection
    # Remember original cursor
    try:
        orig_start = sel.Range.Start
    except Exception:
        orig_start = None

    # Filter out empty lines to avoid redundant empty paragraphs
    lines = [l for l in text_lines if l.strip()]
    if not lines:
        return

    # Navigate to correct position
    # For "before" mode, if we need to insert before a specific paragraph,
    # we go to that paragraph then go to its start
    # For "end" mode, we just type at the end

    if insert_before:
        # Insert before: we type paragraphs backwards? No - we go to start of target,
        # type, then the existing content moves down. But we need to handle this differently.
        # Actually for "before", we go to start of paragraph and type.
        pass

    # Build all text with paragraph breaks, then insert at once
    # Use the Content range approach which reliably creates separate paragraphs
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    # Insert paragraph breaks before each line except the first
    for i, line in enumerate(lines):
        if i > 0:
            rng.InsertParagraphAfter()
            # Move range to after the newly created paragraph
            rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        rng.InsertAfter(line)

    # Restore cursor if needed
    if orig_start is not None:
        try:
            sel.SetRange(orig_start, orig_start)
        except Exception:
            pass


def insert_text(text: str, position: str = "end", para_index: Optional[int] = None, doc_index: Optional[int] = None) -> Dict:
    """Insert text creating real paragraphs for each newline."""
    doc = get_doc(doc_index)
    lines = text.split("\n")

    # Determine target range
    if position == "end":
        if doc.Tables.Count > 0:
            first_tbl_start = com_property(doc.Tables.Item(1).Range, "Start", 0)
            rng = doc.Range(first_tbl_start - 1, first_tbl_start - 1)
        else:
            rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    elif para_index is not None and position == "before":
        rng = doc.Paragraphs.Item(para_index).Range.Duplicate
        rng.Collapse(0)  # Collapse to start
    elif para_index is not None and position == "after":
        rng = doc.Paragraphs.Item(para_index).Range.Duplicate
        rng.Collapse(1)  # Collapse to end
    else:
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)

    # Insert each line as a separate paragraph
    para_count = 0
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped and i == 0:
            continue  # skip leading blank line
        if i > 0 or stripped:
            if para_count > 0 or (i > 0 and stripped):
                rng.InsertParagraphAfter()
            if stripped:
                rng.InsertAfter(stripped)
                para_count += 1
            # Advance range past what we just inserted
            rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)

    return {"inserted": True, "paragraphs_created": para_count, "position": position}


def delete_range(start_pos: int, end_pos: int, doc_index: Optional[int] = None) -> Dict:
    r = get_doc(doc_index).Range(start_pos, end_pos)
    r.Delete()
    return {"deleted": True}


def replace_range(start_pos: int, end_pos: int, new_text: str, doc_index: Optional[int] = None) -> Dict:
    r = get_doc(doc_index).Range(start_pos, end_pos)
    r.Text = new_text
    return {"replaced": True}
