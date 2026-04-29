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


def insert_text(text: str, position: str = "end", para_index: Optional[int] = None, doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    if position == "end":
        r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        r.InsertAfter(text)
        return {"inserted": True, "position": "end"}
    elif para_index is not None:
        p = doc.Paragraphs.Item(para_index)
        if position == "before":
            p.Range.InsertBefore(text)
        elif position == "after":
            p.Range.InsertAfter(text)
        return {"inserted": True, "position": f"{position} paragraph {para_index}"}
    return {"inserted": False, "error": "Invalid position"}


def delete_range(start_pos: int, end_pos: int, doc_index: Optional[int] = None) -> Dict:
    r = get_doc(doc_index).Range(start_pos, end_pos)
    r.Delete()
    return {"deleted": True}


def replace_range(start_pos: int, end_pos: int, new_text: str, doc_index: Optional[int] = None) -> Dict:
    r = get_doc(doc_index).Range(start_pos, end_pos)
    r.Text = new_text
    return {"replaced": True}
