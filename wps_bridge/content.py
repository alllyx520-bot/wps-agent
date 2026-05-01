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
            elif item_type == "runs_detail":
                data = runs_detail(item["para_index"], doc_index)
            elif item_type == "document_structure":
                data = document_structure(doc_index)
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
    import re
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs.Item(i)
        level = com_property(p.Format, "OutlineLevel", 10)
        text = com_property(p.Range, "Text", "").strip()
        style = com_property(p.Range.Style, "NameLocal", "")
        # Detect heading by outline level
        if 1 <= level <= 9:
            result.append({"index": i, "text": text, "outline_level": level, "style": style})
        # Also detect headings by style name even if outline level is body text (10)
        elif level == 10 and ("标题" in style or "Heading" in style.lower()):
            # Infer level from style name
            level_match = re.search(r'(\d+)', style)
            inferred = int(level_match.group(1)) if level_match else 1
            result.append({"index": i, "text": text, "outline_level": inferred, "style": style})
        # Detect headings by text pattern (Chapter/Section numbering)
        elif level == 10 and text and (re.match(r'^第[一二三四五六七八九十百千]+[章节篇部]', text) or re.match(r'^(\d+[\.\、])+\s*\S', text)):
            result.append({"index": i, "text": text, "outline_level": 1, "style": style, "inferred": True})
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


def insert_text(text: str, position: str = "end", para_index: Optional[int] = None, doc_index: Optional[int] = None) -> Dict:
    """Insert text creating real paragraphs for each newline (\r or \n)."""
    import re
    doc = get_doc(doc_index)
    lines = [l for l in re.split(r'[\r\n]+', text) if l]
    if lines and not lines[0].strip() and position == "end":
        lines = lines[1:]

    if position == "end":
        para_count = 0
        for line in lines:
            if not line.strip():
                continue
            rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            rng.InsertParagraphAfter()
            new_para = doc.Paragraphs.Item(doc.Paragraphs.Count)
            new_para.Range.Text = line
            para_count += 1
        return {"inserted": True, "paragraphs_created": para_count, "position": position}
    elif para_index is not None and position == "before":
        pi = max(para_index, 1)
        if pi > doc.Paragraphs.Count:
            return {"error": f"Paragraph {para_index} out of range (document has {doc.Paragraphs.Count} paragraphs)", "error_code": "PARAGRAPH_OUT_OF_RANGE"}
        para_count = 0
        for line in reversed(lines):
            if not line.strip():
                continue
            rng = doc.Paragraphs.Item(pi).Range.Duplicate
            rng.Collapse(0)
            rng.InsertParagraph()
            new_para = doc.Paragraphs.Item(pi)
            new_para.Range.Text = line
            para_count += 1
        return {"inserted": True, "paragraphs_created": para_count, "position": position}
    elif para_index is not None and position == "after":
        pi = max(para_index, 1)
        if pi > doc.Paragraphs.Count:
            return {"error": f"Paragraph {para_index} out of range (document has {doc.Paragraphs.Count} paragraphs)", "error_code": "PARAGRAPH_OUT_OF_RANGE"}
        para_count = 0
        for line in lines:
            if not line.strip():
                continue
            rng = doc.Paragraphs.Item(pi).Range.Duplicate
            rng.Collapse(1)
            rng.InsertParagraphAfter()
            new_index = min(pi + 1, doc.Paragraphs.Count)
            new_para = doc.Paragraphs.Item(new_index)
            new_para.Range.Text = line
            para_count += 1
            pi = new_index
        return {"inserted": True, "paragraphs_created": para_count, "position": position}
    else:
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        para_count = 0
        for i, line in enumerate(lines):
            if i == 0 and not line.strip():
                continue
            if i > 0:
                rng.InsertParagraphAfter()
                rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            if line.strip():
                rng.InsertAfter(line.strip())
            para_count += 1
        return {"inserted": True, "paragraphs_created": para_count, "position": position}


def delete_range(start_pos: int, end_pos: int = None, doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    if end_pos is None:
        end_pos = doc.Content.End
    if end_pos <= 0 or end_pos <= start_pos:
        return {"deleted": False, "warning": "delete_range: zero-width or invalid range, nothing deleted"}
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
    if not lines:
        return {"error": "create_cover requires 'lines' array. Example: [{'text':'Title','font_name':'黑体','font_size':26,'bold':true,'alignment':'center'}]",
                "error_code": "MISSING_PARAM"}
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
    """Replace text at character range while preserving formatting of the surrounding paragraph.
    
    Uses Find+Replace strategy for reliable text substitution.
    Falls back to Range.Text assignment only if the caller explicitly provides
    character positions and wants raw replacement (format may not be preserved).
    """
    doc = get_doc(doc_index)
    # Strategy: use Range.Text replacement with format preservation
    # First get the paragraph that contains this range to preserve its format
    try:
        r = doc.Range(start_pos, end_pos)
        # Get formatting from the range before clearing
        fmt = r.Font.Duplicate
        para_fmt = r.ParagraphFormat.Duplicate
        # Perform the replacement
        r.Text = new_text
        # Re-select the new text and restore formatting
        new_range = doc.Range(start_pos, start_pos + len(new_text))
        try:
            new_range.Font = fmt
            new_range.ParagraphFormat = para_fmt
        except Exception:
            pass  # Format restoration best-effort
        return {"replaced": True, "start": start_pos, "end": start_pos + len(new_text)}
    except Exception as e:
        return {"replaced": False, "error": str(e)}


def runs_detail(para_index: int, doc_index: Optional[int] = None) -> Dict:
    doc = get_doc(doc_index)
    p = doc.Paragraphs.Item(para_index)
    pf = p.Format
    rng = p.Range

    runs = []
    word_count = 0
    try:
        word_count = rng.Words.Count
        for wi in range(1, min(word_count + 1, 200)):
            try:
                w = rng.Words.Item(wi)
                wf = w.Font
                wtext = com_property(w, "Text", "")
                runs.append({
                    "index": wi,
                    "text": wtext,
                    "font": {
                        "name": com_property(wf, "Name", ""),
                        "name_far_east": com_property(wf, "NameFarEast", ""),
                        "size": com_property(wf, "Size", 0),
                        "bold": bool(com_property(wf, "Bold", 0)),
                        "italic": bool(com_property(wf, "Italic", 0)),
                        "underline": com_property(wf, "Underline", 0),
                        "color_index": com_property(wf, "ColorIndex", 0),
                        "superscript": bool(com_property(wf, "Superscript", 0)),
                        "subscript": bool(com_property(wf, "Subscript", 0)),
                        "strike_through": bool(com_property(wf, "StrikeThrough", 0)),
                        "caps": bool(com_property(wf, "AllCaps", 0)),
                        "small_caps": bool(com_property(wf, "SmallCaps", 0)),
                        "shadow": bool(com_property(wf, "Shadow", 0)),
                        "outline": bool(com_property(wf, "Outline", 0)),
                        "emboss": bool(com_property(wf, "Emboss", 0)),
                    },
                })
            except Exception:
                break
    except Exception:
        pass

    char_count = 0
    try:
        char_count = rng.Characters.Count
    except Exception:
        pass

    return {
        "para_index": para_index,
        "text": com_property(rng, "Text", "").strip(),
        "style_name": com_property(rng.Style, "NameLocal", ""),
        "outline_level": com_property(pf, "OutlineLevel", 10),
        "alignment": WDALIGNMENT.get(com_property(pf, "Alignment", 0), "unknown"),
        "first_line_indent": com_property(pf, "FirstLineIndent", 0),
        "left_indent": com_property(pf, "LeftIndent", 0),
        "right_indent": com_property(pf, "RightIndent", 0),
        "line_spacing_rule": WDLINESPACING.get(com_property(pf, "LineSpacingRule", 0), "unknown"),
        "line_spacing": com_property(pf, "LineSpacing", 0),
        "space_before": com_property(pf, "SpaceBefore", 0),
        "space_after": com_property(pf, "SpaceAfter", 0),
        "word_count": word_count,
        "character_count": char_count,
        "runs": runs,
    }


def document_structure(doc_index: Optional[int] = None) -> Dict:
    import re
    doc = get_doc(doc_index)
    total = doc.Paragraphs.Count
    paras = []
    headings = []

    for i in range(1, total + 1):
        try:
            p = doc.Paragraphs.Item(i)
            rng = p.Range
            pf = p.Format
            text = com_property(rng, "Text", "").strip()
            text_short = text[:100]
            style_id = com_property(rng.Style, "NameLocal", "")
            outline_level = com_property(pf, "OutlineLevel", 10)

            heading_level = outline_level if 1 <= outline_level <= 9 else 0

            text_lower = text.lower()

            cross_refs = list(set(re.findall(r'(?:图|表|公式)\s*\d+(?:[.-]\d+)*', text)))

            semantic_role = ""
            if outline_level == 1:
                semantic_role = "一级标题"
            elif outline_level == 2:
                semantic_role = "二级标题"
            elif outline_level == 3:
                semantic_role = "三级标题"
            elif text_lower.startswith("摘要") or text_lower.startswith("abstract"):
                semantic_role = "摘要"
            elif text_lower.startswith("关键词") or text_lower.startswith("keywords"):
                semantic_role = "关键词"
            elif "目录" in text[:4] or "目 录" in text[:4]:
                semantic_role = "目录"
            elif text_lower.startswith("参考文献") or text_lower.startswith("references"):
                semantic_role = "参考文献"
            elif text.startswith("致谢"):
                semantic_role = "致谢"
            elif text.startswith("附录"):
                semantic_role = "附录"
            elif text.startswith("图") and re.match(r'^图\s*\d', text):
                semantic_role = "图标题"
            elif text.startswith("表") and re.match(r'^表\s*\d', text):
                semantic_role = "表标题"
            elif "header" in style_id.lower():
                semantic_role = "页眉"
            elif "footer" in style_id.lower():
                semantic_role = "页脚"
            elif text and text == text.upper() and len(text) < 50:
                semantic_role = "章节标题"
            elif text and re.match(r'^\d+$', text):
                semantic_role = "页码"
            elif not text:
                semantic_role = "空段落"
            else:
                semantic_role = "正文"

            para_info = {
                "index": i,
                "text_short": text_short,
                "style_id": style_id,
                "outline_level": outline_level,
                "heading_level": heading_level,
                "semantic_role": semantic_role,
                "cross_refs": cross_refs,
            }
            paras.append(para_info)

            if 1 <= outline_level <= 9:
                headings.append({
                    "index": i,
                    "level": outline_level,
                    "text": text_short,
                    "style_id": style_id,
                })
        except Exception:
            continue

    return {
        "total_paragraphs": len(paras),
        "total_headings": len(headings),
        "paragraphs": paras,
        "headings": headings,
    }


# ─── Intuitive Selection Operations ───

def find_text(text: str, occurrence: int = 1, match_case: bool = True,
              doc_index: Optional[int] = None) -> Dict:
    """Find the N-th occurrence of text and return its character range."""
    doc = get_doc(doc_index)
    find_obj = doc.Content.Find
    find_obj.ClearFormatting()
    find_obj.Text = text
    find_obj.MatchCase = match_case
    find_obj.Forward = True
    find_obj.Wrap = 0  # wdFindStop

    for _ in range(occurrence):
        found = find_obj.Execute(FindText=text, MatchCase=match_case, Forward=True, Wrap=0)
        if not found:
            return {
                "found": False,
                "occurrences_before": _,
                "total": _,
            }

    sel = get_app().Selection
    return {
        "found": True,
        "start_pos": com_property(sel, "Start", 0),
        "end_pos": com_property(sel, "End", 0),
        "text": com_property(sel.Range, "Text", ""),
        "occurrence": occurrence,
    }


def find_all(text: str, match_case: bool = True,
             doc_index: Optional[int] = None) -> Dict:
    """Find all occurrences of text and return their positions."""
    doc = get_doc(doc_index)
    results = []
    find_obj = doc.Content.Find
    find_obj.ClearFormatting()
    find_obj.Text = text
    find_obj.MatchCase = match_case
    find_obj.Forward = True
    find_obj.Wrap = 0

    while find_obj.Execute(FindText=text, MatchCase=match_case, Forward=True, Wrap=0):
        sel = get_app().Selection
        results.append({
            "start_pos": com_property(sel, "Start", 0),
            "end_pos": com_property(sel, "End", 0),
            "occurrence": len(results) + 1,
        })

    return {"found": len(results), "positions": results}


def replace_by_pattern(pattern: str, replacement: str,
                       doc_index: Optional[int] = None) -> Dict:
    """Regex-based find and replace across the document."""
    doc = get_doc(doc_index)
    full = com_property(doc.Content, "Text", "")
    replaced, count = __import__('re').subn(pattern, replacement, full)
    if count > 0:
        doc.Content.Text = replaced
    return {"replaced_count": count, "pattern": pattern}


def select_by_role(role: str, filepath: str = None,
                   doc_index: Optional[int] = None) -> Dict:
    """Select paragraphs by semantic role using offline engine.
    Returns paragraph indices that match the given role."""
    if not filepath:
        return {"error": "filepath required for semantic role selection"}
    from docx_engine import parse_docx, build_document_model
    from docx_engine.semantic_model import SemanticParser
    parsed = parse_docx(filepath)
    doc = build_document_model(parsed)
    parser = SemanticParser(doc)
    graph = parser.parse()
    matches = graph.get_by_role(role)
    return {
        "role": role,
        "match_count": len(matches),
        "paragraphs": [{"index": e.index, "text_preview": e.text_preview,
                         "confidence": e.confidence} for e in matches],
    }


def insert_paragraph(text: str, style: Optional[str] = None,
                     position: str = "end", para_index: Optional[int] = None,
                     doc_index: Optional[int] = None) -> Dict:
    """One-step paragraph insertion with automatic \n splitting."""
    import re
    doc = get_doc(doc_index)
    lines = [l.strip() for l in re.split(r'[\r\n]+|\\n', text) if l.strip()]
    if not lines:
        lines = [text]
    count = len(lines)

    if position == "end":
        for line in lines:
            rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            rng.InsertParagraphAfter()
            new_para = doc.Paragraphs.Item(doc.Paragraphs.Count)
            new_para.Range.Text = line
            if style:
                try:
                    new_para.Range.Style = doc.Styles.Item(style)
                except Exception:
                    pass
    elif position == "before" and para_index is not None:
        pi = max(para_index, 1)
        for line in reversed(lines):
            rng = doc.Paragraphs.Item(pi).Range
            rng.InsertParagraph()
            new_para = doc.Paragraphs.Item(pi)
            new_para.Range.Text = line
            if style:
                try:
                    new_para.Range.Style = doc.Styles.Item(style)
                except Exception:
                    pass
    elif position == "after" and para_index is not None:
        pi = max(para_index, 1)
        for line in reversed(lines):
            rng = doc.Paragraphs.Item(pi).Range
            rng.InsertParagraphAfter()
            new_index = min(pi + 1, doc.Paragraphs.Count)
            new_para = doc.Paragraphs.Item(new_index)
            new_para.Range.Text = line
            if style:
                try:
                    new_para.Range.Style = doc.Styles.Item(style)
                except Exception:
                    pass
    else:
        return {"error": "Invalid position or missing para_index"}
    return {"inserted": True, "paragraphs_created": count, "position": position, "text_preview": lines[0][:80]}


# ─── Surgical Operations (paragraph-aware, no character position fragility) ───

def delete_paragraphs(from_para: int, to_para: int, doc_index: Optional[int] = None) -> Dict:
    """Delete paragraphs in range [from_para, to_para] (inclusive, 1-based)."""
    doc = get_doc(doc_index)
    if from_para < 1:
        from_para = 1
    to_para = max(from_para, to_para)
    total = doc.Paragraphs.Count
    to_para = min(to_para, total)
    removed = 0
    for i in range(to_para, from_para - 1, -1):
        try:
            doc.Paragraphs.Item(i).Range.Delete()
            removed += 1
        except Exception:
            pass
    return {"deleted": removed, "from_para": from_para, "to_para": to_para}


def delete_runs(para_index: int, from_run: int, to_run: int, doc_index: Optional[int] = None) -> Dict:
    """Delete Runs within a single paragraph. Uses Word Range to find and delete run text."""
    doc = get_doc(doc_index)
    if para_index < 1:
        para_index = 1
    p = doc.Paragraphs.Item(para_index)
    rng = p.Range
    words = rng.Words
    total_words = words.Count
    to_run = min(to_run, total_words)
    if from_run < 1:
        from_run = 1
    if from_run > total_words:
        return {"error": f"from_run {from_run} exceeds word count {total_words}", "error_code": "RUN_OUT_OF_RANGE"}
    start_pos = com_property(words.Item(from_run), "Start", 0)
    end_pos = com_property(words.Item(to_run), "End", start_pos)
    if start_pos > 0 and end_pos > start_pos:
        doc.Range(start_pos, end_pos).Delete()
    return {"deleted_runs": to_run - from_run + 1, "para_index": para_index, "run_range": [from_run, to_run]}


def replace_paragraph_text(para_index: int, new_text: str, preserve_format: bool = False,
                           doc_index: Optional[int] = None) -> Dict:
    """Replace the entire text of a paragraph, optionally preserving its format."""
    doc = get_doc(doc_index)
    if para_index < 1:
        para_index = 1
    p = doc.Paragraphs.Item(para_index)
    rng = p.Range
    if preserve_format:
        fmt = rng.Font.Duplicate
        pf = rng.ParagraphFormat.Duplicate
        style = rng.Style
    rng.Text = new_text
    if preserve_format:
        try:
            rng.Font = fmt
            rng.ParagraphFormat = pf
        except Exception:
            pass
    return {"replaced": True, "para_index": para_index, "text_preview": new_text[:80]}


def replace_runs(para_index: int, run_indices: List[int], new_text: str,
                 doc_index: Optional[int] = None) -> Dict:
    """Replace specific Runs in a paragraph with new_text. Preserves other Runs."""
    doc = get_doc(doc_index)
    if para_index < 1:
        para_index = 1
    p = doc.Paragraphs.Item(para_index)
    rng = p.Range
    words = rng.Words
    total = words.Count
    run_indices = sorted(set(run_indices))
    if not run_indices or run_indices[0] < 1 or run_indices[-1] > total:
        return {"error": "run_indices out of range", "error_code": "RUN_OUT_OF_RANGE"}
    start_pos = com_property(words.Item(run_indices[0]), "Start", 0)
    end_pos = com_property(words.Item(run_indices[-1]), "End", start_pos)
    if start_pos > 0 and end_pos > start_pos:
        doc.Range(start_pos, end_pos).Text = new_text
    return {"replaced_runs": len(run_indices), "para_index": para_index, "run_indices": run_indices, "text_preview": new_text[:80]}


# ─── Snapshot / Rollback (paragraph-level) ───

_snapshot_store: Dict[int, List[Dict]] = {}


def snapshot(doc_index: Optional[int] = None) -> Dict:
    """Save full paragraph-level snapshot (text, font, format) for all paragraphs."""
    doc = get_doc(doc_index)
    total = doc.Paragraphs.Count
    snaps = []
    for i in range(1, total + 1):
        try:
            p = doc.Paragraphs.Item(i)
            rng = p.Range
            f = rng.Font
            pf = p.Format
            snaps.append({
                "index": i,
                "text": com_property(rng, "Text", ""),
                "style": com_property(rng.Style, "NameLocal", ""),
                "font": {"name": com_property(f, "Name", ""), "size": com_property(f, "Size", 0),
                         "bold": bool(com_property(f, "Bold", 0))},
                "format": {"alignment_raw": com_property(pf, "Alignment", 0),
                           "line_spacing_rule_raw": com_property(pf, "LineSpacingRule", 0),
                           "space_before": com_property(pf, "SpaceBefore", 0),
                           "space_after": com_property(pf, "SpaceAfter", 0),
                           "outline_level": com_property(pf, "OutlineLevel", 10)},
            })
        except Exception:
            continue
    sid = doc_index if doc_index else 0
    _snapshot_store[sid] = snaps
    return {"snapshot": True, "paragraphs_saved": len(snaps)}


def rollback(doc_index: Optional[int] = None) -> Dict:
    """Restore document from last snapshot. Only restores paragraphs that were captured."""
    doc = get_doc(doc_index)
    sid = doc_index if doc_index else 0
    snaps = _snapshot_store.pop(sid, [])
    if not snaps:
        return {"error": "No snapshot available", "error_code": "NO_SNAPSHOT"}
    restored = 0
    for snap in snaps:
        idx = snap["index"]
        try:
            p = doc.Paragraphs.Item(idx)
            p.Range.Text = snap["text"]
            rng = p.Range
            rng.Style = snap["style"]
            f = rng.Font
            fmt_data = snap["font"]
            com_set(f, "Name", fmt_data["name"])
            com_set(f, "Size", fmt_data["size"])
            com_set(f, "Bold", fmt_data["bold"])
            pf = p.Format
            pf_data = snap["format"]
            com_set(pf, "Alignment", pf_data["alignment_raw"])
            com_set(pf, "SpaceBefore", pf_data["space_before"])
            com_set(pf, "SpaceAfter", pf_data["space_after"])
            com_set(pf, "OutlineLevel", pf_data["outline_level"])
            restored += 1
        except Exception:
            continue
    return {"rolled_back": True, "restored_paragraphs": restored, "total_snapshot": len(snaps)}


def doc_build(structure: Dict, output_path: str, doc_index: Optional[int] = None) -> Dict:
    """High-level document builder. One call to build a complete document.

    structure = {
        "cover": {"lines": [{text, font_name?, font_size?, bold?, alignment?, space_before?, space_after?}]},
        "sections": [
            {"heading": "一、...", "paragraphs": ["body1", "body2"],
             "table": {"headers": ["col1","col2"], "rows": [["a","b"],["c","d"]]}}
        ],
        "defaults": {body_font, body_size, heading_font, heading_size, first_line_indent},
        "page_setup": {page_width, page_height, top_margin, bottom_margin, left_margin, right_margin}
    }
    Returns: {built, paragraphs_created, tables_created, sections_created, cover_lines, output_path, errors}
    """
    from . import formatting as _fmt
    from . import table as _tbl

    doc = get_doc(doc_index)
    errors = []
    para_count = 0
    table_count = 0
    defaults = structure.get("defaults", {
        "body_font": "宋体", "body_size": 12,
        "heading_font": "黑体", "heading_size": 16,
    })

    # 1. Clear + Cover
    _clear_document(doc)
    cover_lines_data = structure.get("cover", {}).get("lines", [])
    if cover_lines_data:
        create_cover(cover_lines_data, clear_existing=False, doc_index=doc_index)
        para_count += len(cover_lines_data)

        # Page break after cover
        try:
            last_cover_para = doc.Paragraphs.Count
            from . import layout as _lay
            _lay.insert_page_break(last_cover_para, doc_index)
        except Exception as e:
            errors.append(f"page_break after cover: {e}")

    # 2. Page setup
    ps = structure.get("page_setup", {})
    if ps:
        try:
            from . import layout as _lay
            _lay.page_setup(
                doc_index, None,
                page_width=ps.get("page_width"), page_height=ps.get("page_height"),
                top_margin=ps.get("top_margin"), bottom_margin=ps.get("bottom_margin"),
                left_margin=ps.get("left_margin"), right_margin=ps.get("right_margin"),
                orientation=ps.get("orientation"), gutter=ps.get("gutter"),
            )
        except Exception as e:
            errors.append(f"page_setup: {e}")

    # 3. Build sections
    sections_data = structure.get("sections", [])
    heading_font = defaults.get("heading_font", "黑体")
    heading_size = defaults.get("heading_size", 16)
    body_font = defaults.get("body_font", "宋体")
    body_size = defaults.get("body_size", 12)
    first_line_indent = defaults.get("first_line_indent")

    for sec_idx, section in enumerate(sections_data):
        heading_text = section.get("heading", "").strip()
        if not heading_text:
            continue

        # Insert heading
        insert_paragraph(heading_text, position="end", doc_index=doc_index)
        para_count += 1
        heading_para_idx = doc.Paragraphs.Count

        # Format heading
        try:
            _fmt.set_font(heading_para_idx, name=heading_font, size=heading_size, bold=True, doc_index=doc_index)
            _fmt.set_paragraph_format(heading_para_idx, space_before=12, space_after=6, doc_index=doc_index)
        except Exception as e:
            errors.append(f"section {sec_idx} heading format: {e}")

        # Insert body paragraphs
        body_paragraphs = section.get("paragraphs", [])
        for bp_text in body_paragraphs:
            if not bp_text or not bp_text.strip():
                continue
            insert_paragraph(bp_text, position="end", doc_index=doc_index)
            para_count += 1
            body_para_idx = doc.Paragraphs.Count
            try:
                _fmt.set_font(body_para_idx, name=body_font, size=body_size, doc_index=doc_index)
                indent_kwargs = {}
                if first_line_indent:
                    indent_kwargs["first_line_indent"] = first_line_indent
                _fmt.set_paragraph_format(body_para_idx, line_spacing_rule="multiple", line_spacing=1.5, **indent_kwargs, doc_index=doc_index)
            except Exception as e:
                errors.append(f"section {sec_idx} body format para {body_para_idx}: {e}")

        # Insert table if present
        table_data = section.get("table")
        if table_data:
            headers = table_data.get("headers", [])
            rows = table_data.get("rows", [])
            num_cols = len(headers) if headers else (len(rows[0]) if rows else 3)
            num_rows = len(rows) + (1 if headers else 0)
            if num_rows > 0:
                try:
                    tbl_result = _tbl.table_create(num_rows, num_cols, position="end", doc_index=doc_index)
                    if "error" not in tbl_result:
                        table_count += 1
                        tbl_idx = tbl_result["table_index"]
                        header_bold = table_data.get("header_bold", True)
                        # Fill headers
                        for c_idx, h_text in enumerate(headers, 1):
                            if h_text:
                                _tbl.set_cell_text(tbl_idx, 1, c_idx, h_text, doc_index=doc_index)
                                if header_bold:
                                    _tbl.format_cell(tbl_idx, 1, c_idx, bold=True, font_name=body_font, font_size=body_size, doc_index=doc_index)
                        # Fill rows
                        for r_idx, row_data in enumerate(rows, 2 if headers else 1):
                            for c_idx, cell_text in enumerate(row_data, 1):
                                if c_idx <= num_cols:
                                    _tbl.set_cell_text(tbl_idx, r_idx, c_idx, str(cell_text), doc_index=doc_index)
                                    _tbl.format_cell(tbl_idx, r_idx, c_idx, font_name=body_font, font_size=body_size, doc_index=doc_index)
                        # Set header shading
                        header_shading = table_data.get("header_shading")
                        if header_shading and headers:
                            for c_idx in range(1, num_cols + 1):
                                try:
                                    _tbl.set_cell_shading(tbl_idx, 1, c_idx, header_shading, doc_index=doc_index)
                                except Exception:
                                    pass
                    else:
                        errors.append(f"table create: {tbl_result.get('error')}")
                except Exception as e:
                    errors.append(f"section {sec_idx} table: {e}")

    # 4. Save
    if output_path:
        try:
            doc.SaveAs(output_path)
        except Exception as e:
            try:
                doc.SaveAs2(output_path)
            except Exception as e2:
                errors.append(f"save: {e2}")

    return {
        "built": True,
        "paragraphs_created": para_count,
        "tables_created": table_count,
        "sections_created": len(sections_data),
        "cover_lines": len(cover_lines_data),
        "output_path": output_path,
        "errors": errors,
    }
