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
            elif item_type == "shapes":
                data = shapes(item.get("include_inlines", True), doc_index)
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


def shapes(include_inlines: bool = True, doc_index: Optional[int] = None) -> Dict:
    """Read all shapes/drawings in the document body. Returns text content, type, position, size of each shape."""
    WDSHAPES = {
        1: "auto_shape", 2: "callout", 3: "chart", 5: "freeform",
        6: "group", 9: "line", 11: "linked_picture", 13: "picture",
        15: "text_effect", 16: "media", 17: "text_box", 19: "table",
        20: "canvas", 21: "diagram", 22: "ink", 23: "ink_comment",
        24: "smart_art", 25: "web_video",
    }
    doc = get_doc(doc_index)
    result = []

    # --- Floating shapes (doc.Shapes) ---
    try:
        for i in range(1, doc.Shapes.Count + 1):
            try:
                shp = doc.Shapes.Item(i)
                item = _describe_shape(shp, i, "floating", WDSHAPES, doc)
                if item:
                    result.append(item)
            except Exception:
                continue
    except Exception:
        pass

    # --- Inline shapes (doc.InlineShapes) ---
    if include_inlines:
        try:
            for i in range(1, doc.InlineShapes.Count + 1):
                try:
                    ishp = doc.InlineShapes.Item(i)
                    item = _describe_inline_shape(ishp, i, doc)
                    if item:
                        result.append(item)
                except Exception:
                    continue
        except Exception:
            pass

    return {
        "total": len(result),
        "shapes": result,
        "floating_count": _safe_count(doc.Shapes),
        "inline_count": _safe_count(doc.InlineShapes) if include_inlines else 0,
    }


def _describe_shape(shp, idx, scope, WDSHAPES, doc) -> Optional[Dict]:
    item = {"index": idx, "scope": scope}
    shape_type = 0
    try:
        shape_type = com_property(shp, "Type", 0)
        item["type"] = WDSHAPES.get(shape_type, f"unknown_{shape_type}")
        item["type_id"] = shape_type
    except Exception:
        item["type"] = "unknown"
    try:
        item["name"] = com_property(shp, "Name", "")
    except Exception:
        item["name"] = ""
    try:
        item["visible"] = bool(com_property(shp, "Visible", 1))
    except Exception:
        pass
    try:
        item["alternative_text"] = com_property(shp, "AlternativeText", "")
    except Exception:
        pass
    # Position and size
    try:
        item["left"] = com_property(shp, "Left", 0)
        item["top"] = com_property(shp, "Top", 0)
        item["width"] = com_property(shp, "Width", 0)
        item["height"] = com_property(shp, "Height", 0)
    except Exception:
        pass
    # Text content
    try:
        if bool(com_property(shp, "HasTextFrame", 0)):
            tf = shp.TextFrame
            item["text"] = com_property(tf.TextRange, "Text", "").strip()
            item["has_text"] = len(item.get("text", "")) > 0
    except Exception:
        pass
    # Group items
    if shape_type == 6:  # msoGroup
        try:
            group_items = []
            for gi in range(1, shp.GroupItems.Count + 1):
                try:
                    gshp = shp.GroupItems.Item(gi)
                    git = _describe_shape(gshp, gi, "group", WDSHAPES, doc)
                    if git:
                        group_items.append(git)
                except Exception:
                    continue
            if group_items:
                item["group_items"] = group_items
                item["group_total"] = len(group_items)
        except Exception:
            pass
    # Anchor paragraph
    try:
        anchor = shp.Anchor
        item["anchor_paragraph"] = anchor.Paragraphs.Item(1).Range.Start if anchor.Paragraphs.Count > 0 else "unknown"
    except Exception:
        pass
    # Rotation
    try:
        rot = com_property(shp, "Rotation", 0)
        if rot:
            item["rotation"] = rot
    except Exception:
        pass
    return item


def _describe_inline_shape(ishp, idx, doc) -> Optional[Dict]:
    item = {"index": idx, "scope": "inline"}
    try:
        shape_type = com_property(ishp, "Type", 0)
        type_names = {1: "picture", 2: "linked_picture", 3: "ole_object", 4: "linked_ole_object", 5: "horizontal_line", 6: "chart", 9: "smart_art", 10: "3d_model", 12: "web_video"}
        item["type"] = type_names.get(shape_type, f"unknown_{shape_type}")
        item["type_id"] = shape_type
    except Exception:
        item["type"] = "unknown"
    try:
        item["width"] = com_property(ishp, "Width", 0)
        item["height"] = com_property(ishp, "Height", 0)
    except Exception:
        pass
    try:
        item["alternative_text"] = com_property(ishp, "AlternativeText", "")
    except Exception:
        pass
    return item


def _safe_count(collection) -> int:
    try:
        return collection.Count
    except Exception:
        return 0


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


def delete_range(start_pos: int, end_pos: int = None, doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    if end_pos is None or end_pos <= 0:
        end_pos = doc.Content.End
    r = doc.Range(start_pos, end_pos)
    r.Delete()
    return {"deleted": True}


def _apply_line_format(doc, para_index: int, line: Dict):
    """Apply font + paragraph formatting to a single paragraph from a line spec."""
    from .utils import com_set, WDALIGNMENT, WDLINESPACING
    p = doc.Paragraphs.Item(para_index)
    r = p.Range
    f = r.Font
    pf = p.Format

    com_set(f, "ColorIndex", 1)
    if "font_name" in line:
        com_set(f, "Name", line["font_name"])
        com_set(f, "NameFarEast", line["font_name"])
    if "font_size" in line:
        com_set(f, "Size", line["font_size"])
    if "bold" in line:
        com_set(f, "Bold", line["bold"])
    if "italic" in line:
        com_set(f, "Italic", line["italic"])
    if "underline" in line:
        com_set(f, "Underline", line["underline"])
    if "strike_through" in line:
        com_set(f, "StrikeThrough", line["strike_through"])

    if "alignment" in line:
        align_str = line["alignment"]
        if isinstance(align_str, str):
            align_val = WDALIGNMENT.get(align_str)
            if align_val is not None:
                com_set(pf, "Alignment", align_val)

    # Explicitly set spacing to prevent inherited style defaults from adding gaps
    com_set(pf, "SpaceBefore", line.get("space_before", 0))
    com_set(pf, "SpaceAfter", line.get("space_after", 0))

    if "line_spacing_rule" in line:
        lsr = line["line_spacing_rule"]
        if isinstance(lsr, str):
            lsr_val = WDLINESPACING.get(lsr)
            if lsr_val is not None:
                com_set(pf, "LineSpacingRule", lsr_val)
    if "line_spacing" in line:
        com_set(pf, "LineSpacing", line["line_spacing"])
    if "first_line_indent" in line:
        com_set(pf, "FirstLineIndent", line["first_line_indent"])
    if "left_indent" in line:
        com_set(pf, "LeftIndent", line["left_indent"])
    if "right_indent" in line:
        com_set(pf, "RightIndent", line["right_indent"])
    if "outline_level" in line:
        com_set(pf, "OutlineLevel", line["outline_level"])


def _clear_document(doc):
    """Reliably clear all content. Word/WPS always keeps the final ¶, so Content.Text="" is safe."""
    try:
        doc.Content.Text = ""
    except Exception:
        # Fallback: delete range except the final paragraph mark
        end = doc.Content.End
        if end > 1:
            doc.Range(0, end - 1).Delete()


def create_cover(lines: List[Dict], clear_existing: bool = True, doc_index: Optional[int] = None) -> Dict:
    """Single-call cover page creation. Each line: {text, font_name, font_size, bold, alignment, space_before, space_after, ...}"""
    doc = get_doc(doc_index)

    if clear_existing:
        _clear_document(doc)

    created = []

    for i, line in enumerate(lines):
        if not isinstance(line, dict):
            continue
        text = line.get("text", "").strip()
        if not text:
            continue

        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        if i > 0:
            rng.InsertParagraphAfter()
            rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)

        rng.InsertAfter(text)
        para_idx = doc.Paragraphs.Count
        _apply_line_format(doc, para_idx, line)
        created.append({"para_index": para_idx, "text": text[:50]})

    return {"created": True, "paragraphs": created, "total": len(created)}


def replace_range(start_pos: int, end_pos: int, new_text: str, doc_index: Optional[int] = None) -> Dict:
    r = get_doc(doc_index).Range(start_pos, end_pos)
    r.Text = new_text
    return {"replaced": True}
