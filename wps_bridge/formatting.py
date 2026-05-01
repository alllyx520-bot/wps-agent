# -*- coding: utf-8 -*-
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set, com_set_batch, WDALIGNMENT, WDLINESPACING, WDBUILTINSTYLE, WDSTYLETYPE


def _resolve_range(doc, para_index, start_pos, end_pos, use_selection):
    if use_selection:
        return get_app().Selection.Range
    if start_pos is not None and end_pos is not None:
        return doc.Range(start_pos, end_pos)
    if para_index is not None:
        if para_index < 1:
            para_index = 1
        return doc.Paragraphs.Item(para_index).Range
    return doc.Content


def get_font(para_index=None, start_pos=None, end_pos=None, use_selection=False, doc_index=None):
    doc = get_doc(doc_index)
    r = _resolve_range(doc, para_index, start_pos, end_pos, use_selection)
    f = r.Font
    return {"name": com_property(f, "Name", ""), "name_far_east": com_property(f, "NameFarEast", ""), "size": com_property(f, "Size", 0), "bold": bool(com_property(f, "Bold", 0)), "italic": bool(com_property(f, "Italic", 0)), "underline": com_property(f, "Underline", 0), "color_index": com_property(f, "ColorIndex", 0), "superscript": bool(com_property(f, "Superscript", 0)), "subscript": bool(com_property(f, "Subscript", 0)), "strike_through": bool(com_property(f, "StrikeThrough", 0)), "spacing": com_property(f, "Spacing", 0), "scaling": com_property(f, "Scaling", 100), "kerning": com_property(f, "Kerning", 0)}


def set_font(para_index=None, start_pos=None, end_pos=None, use_selection=False, doc_index=None, **kwargs):
    doc = get_doc(doc_index)
    r = _resolve_range(doc, para_index, start_pos, end_pos, use_selection)
    f = r.Font
    props = {}
    for key, com_name in [
        ("name", "Name"), ("name_far_east", "NameFarEast"),
        ("size", "Size"), ("bold", "Bold"), ("underline", "Underline"),
        ("color_index", "ColorIndex"), ("superscript", "Superscript"),
        ("subscript", "Subscript"), ("strike_through", "StrikeThrough"),
        ("spacing", "Spacing"), ("scaling", "Scaling"), ("kerning", "Kerning"),
        ("caps", "AllCaps"), ("small_caps", "SmallCaps"),
        ("shadow", "Shadow"), ("outline", "Outline"),
        ("emboss", "Emboss"), ("vanish", "Hidden"),
    ]:
        if kwargs.get(key) is not None:
            props[com_name] = kwargs[key]
    if kwargs.get("italic") is not None:
        props["Italic"] = kwargs["italic"]
    failed = com_set_batch(f, props)
    hilite = kwargs.get("highlight")
    if hilite is not None:
        try:
            f.HighlightColorIndex = hilite
        except Exception:
            failed.append("HighlightColorIndex")
    color_rgb = kwargs.get("color_rgb")
    if color_rgb is not None:
        try:
            f.TextColor.RGB = color_rgb
        except Exception:
            try:
                f.Color = color_rgb
            except Exception:
                failed.append("color_rgb")
    return {"updated": True, "failed": failed, "text_sample": com_property(r, "Text", "")[:50]}


def get_paragraph_format(para_index, doc_index=None):
    doc = get_doc(doc_index)
    valid_idx = max(para_index, 1)
    if valid_idx > doc.Paragraphs.Count:
        return {"error": f"Paragraph {para_index} out of range", "error_code": "PARAGRAPH_OUT_OF_RANGE"}
    p = doc.Paragraphs.Item(valid_idx)
    pf = p.Format
    return {"alignment": WDALIGNMENT.get(com_property(pf, "Alignment", 0), "unknown"), "alignment_raw": com_property(pf, "Alignment", 0), "first_line_indent": com_property(pf, "FirstLineIndent", 0), "left_indent": com_property(pf, "LeftIndent", 0), "right_indent": com_property(pf, "RightIndent", 0), "line_spacing_rule": WDLINESPACING.get(com_property(pf, "LineSpacingRule", 0), "unknown"), "line_spacing_rule_raw": com_property(pf, "LineSpacingRule", 0), "line_spacing": com_property(pf, "LineSpacing", 0), "space_before": com_property(pf, "SpaceBefore", 0), "space_after": com_property(pf, "SpaceAfter", 0), "outline_level": com_property(pf, "OutlineLevel", 10), "widow_control": bool(com_property(pf, "WidowControl", 0)), "keep_with_next": bool(com_property(pf, "KeepWithNext", 0)), "keep_together": bool(com_property(pf, "KeepTogether", 0)), "page_break_before": bool(com_property(pf, "PageBreakBefore", 0))}


def set_paragraph_format(para_index=None, use_selection=False, doc_index=None, **kwargs):
    doc = get_doc(doc_index)
    if para_index is not None:
        if para_index < 1:
            para_index = 1
        pc = doc.Paragraphs.Count
        if para_index > pc:
            return {"error": f"Paragraph {para_index} out of range (document has {pc} paragraphs)", "error_code": "PARAGRAPH_OUT_OF_RANGE", "para_count": pc}
    if use_selection:
        pf = get_app().Selection.ParagraphFormat
    elif para_index is not None:
        pf = doc.Paragraphs.Item(para_index).Format
    else:
        pf = doc.Content.ParagraphFormat
    align = kwargs.get("alignment")
    if isinstance(align, str):
        align = WDALIGNMENT.get(align, align)
    lsr = kwargs.get("line_spacing_rule")
    if isinstance(lsr, str):
        lsr = WDLINESPACING.get(lsr, lsr)
    ls = kwargs.get("line_spacing")
    outline = kwargs.get("outline_level")
    props = {}
    if align is not None:
        try:
            props["Alignment"] = int(align)
        except (ValueError, TypeError):
            return {"error": f"Invalid alignment value: {align}", "error_code": "INVALID_PARAM"}
    if (ks := kwargs.get("first_line_indent")) is not None:
        props["FirstLineIndent"] = float(ks)
    if (ks := kwargs.get("left_indent")) is not None:
        props["LeftIndent"] = float(ks)
    if (ks := kwargs.get("right_indent")) is not None:
        props["RightIndent"] = float(ks)
    if lsr is not None:
        props["LineSpacingRule"] = int(lsr)
    if ls is not None:
        props["LineSpacing"] = float(ls)
        if lsr is None:
            props["LineSpacingRule"] = 5
    if (ks := kwargs.get("space_before")) is not None:
        props["SpaceBefore"] = float(ks)
    if (ks := kwargs.get("space_after")) is not None:
        props["SpaceAfter"] = float(ks)
    if outline is not None:
        props["OutlineLevel"] = int(outline)
    if (wc := kwargs.get("widow_control")) is not None:
        props["WidowControl"] = int(wc)
    if (kn := kwargs.get("keep_with_next")) is not None:
        props["KeepWithNext"] = int(kn)
    failed = com_set_batch(pf, props)
    return {"updated": True, "failed": failed, "para_index": para_index}


def apply_style(style_name, para_index=None, use_selection=False, doc_index=None):
    doc = get_doc(doc_index)
    try:
        style = doc.Styles.Item(style_name)
    except Exception:
        index = WDBUILTINSTYLE.get(style_name)
        if index is not None:
            style = doc.Styles.Item(index)
        else:
            return {"error": f"Style not found: {style_name}"}
    if use_selection:
        get_app().Selection.ParagraphFormat.Style = style
    elif para_index is not None:
        if para_index < 1:
            para_index = 1
        if para_index > doc.Paragraphs.Count:
            return {"error": f"Paragraph {para_index} out of range", "error_code": "PARAGRAPH_OUT_OF_RANGE"}
        doc.Paragraphs.Item(para_index).Range.Style = style
    else:
        return {"error": "Specify para_index or use_selection=True"}
    return {"applied": com_property(style, "NameLocal", "")}


def clear_formatting(para_index=None, use_selection=False, doc_index=None):
    doc = get_doc(doc_index)
    if use_selection:
        r = get_app().Selection.Range
    elif para_index is not None:
        r = doc.Paragraphs.Item(para_index).Range
    else:
        return {"error": "Specify para_index or use_selection=True"}
    try:
        r.Font.Reset()
        r.ParagraphFormat.Reset()
    except Exception:
        pass
    return {"cleared": True}


def copy_format(source_para_index, target_para_indices, doc_index=None):
    doc = get_doc(doc_index)
    sel = get_app().Selection
    # Remember original selection position
    try:
        orig_start = sel.Range.Start
        orig_end = sel.Range.End
    except Exception:
        orig_start = orig_end = None
    # Select source paragraph range
    doc.Paragraphs.Item(source_para_index).Range.Select()
    sel.CopyFormat()
    applied = []
    for idx in target_para_indices:
        try:
            doc.Paragraphs.Item(idx).Range.Select()
            sel.PasteFormat()
            applied.append(idx)
        except Exception:
            continue
    # Restore original selection
    if orig_start is not None:
        doc.Range(orig_start, orig_end).Select()
    return {"copied_from": source_para_index, "applied_to": applied}


def batch(operations: list, doc_index=None):
    if not operations:
        return {"executed": 0, "details": []}
    doc = get_doc(doc_index)
    results = []
    success = 0
    for op in operations:
        op_type = op.get("type", op.get("action", ""))
        try:
            if op_type == "set_font":
                res = set_font(doc_index=doc_index, **{k: v for k, v in op.items() if k not in ("type", "action")})
            elif op_type == "set_paragraph_format":
                res = set_paragraph_format(doc_index=doc_index, **{k: v for k, v in op.items() if k not in ("type", "action")})
            elif op_type == "apply_style":
                res = apply_style(op["style_name"], op.get("para_index"), False, doc_index)
            elif op_type == "clear_formatting":
                res = clear_formatting(op.get("para_index"), False, doc_index)
            elif op_type == "copy_format":
                res = copy_format(op["source_para_index"], op["target_para_indices"], doc_index)
            elif op_type == "get_font":
                res = get_font(op.get("para_index"), op.get("start_pos"), op.get("end_pos"), op.get("use_selection", False), doc_index)
            elif op_type == "get_paragraph_format":
                res = get_paragraph_format(op["para_index"], doc_index)
            elif op_type == "get_run_font":
                res = get_run_font(op["para_index"], op["run_index"], doc_index)
            elif op_type == "set_run_font":
                res = set_run_font(op["para_index"], op["run_index"], doc_index, **{k: v for k, v in op.items() if k not in ("type", "action", "para_index", "run_index")})
            elif op_type == "set_tab_stops":
                res = set_tab_stops(op["para_index"], op.get("stops", []), doc_index)
            else:
                results.append({"ok": False, "type": op_type, "error": f"Unknown op type: {op_type}"})
                continue
            results.append({"ok": True, "type": op_type, "result": res})
            success += 1
        except Exception as e:
            results.append({"ok": False, "type": op_type, "error": str(e)})
    return {"total": len(operations), "success": success, "failed": len(operations) - success, "details": results}


def list_styles(doc_index=None, builtin_only=False):
    doc = get_doc(doc_index)
    styles = doc.Styles
    result = []
    count = min(styles.Count, 200)
    for i in range(1, count + 1):
        try:
            s = styles.Item(i)
            result.append({"name": com_property(s, "NameLocal", ""), "type": WDSTYLETYPE.get(com_property(s, "Type", 0), "unknown"), "builtin": bool(com_property(s, "BuiltIn", 0))})
        except Exception:
            continue
    return result


def get_style(style_name, doc_index=None):
    doc = get_doc(doc_index)
    try:
        s = doc.Styles.Item(style_name)
    except Exception:
        index = WDBUILTINSTYLE.get(style_name)
        if index is not None:
            s = doc.Styles.Item(index)
        else:
            return {"error": f"Style not found: {style_name}"}
    return {"name": com_property(s, "NameLocal", ""), "type": WDSTYLETYPE.get(com_property(s, "Type", 0), "unknown"), "builtin": bool(com_property(s, "BuiltIn", 0)), "base_style": com_property(s.BaseStyle, "NameLocal", "") if com_property(s, "BaseStyle", None) else "", "font": {"name": com_property(s.Font, "Name", ""), "size": com_property(s.Font, "Size", 0), "bold": bool(com_property(s.Font, "Bold", 0))}, "paragraph_format": {"alignment": WDALIGNMENT.get(com_property(s.ParagraphFormat, "Alignment", 0), "unknown"), "first_line_indent": com_property(s.ParagraphFormat, "FirstLineIndent", 0), "line_spacing_rule": WDLINESPACING.get(com_property(s.ParagraphFormat, "LineSpacingRule", 0), "unknown"), "space_before": com_property(s.ParagraphFormat, "SpaceBefore", 0), "space_after": com_property(s.ParagraphFormat, "SpaceAfter", 0)}, "description": com_property(s, "Description", "")}


def create_style(name, base_style=None, doc_index=None, **kwargs):
    doc = get_doc(doc_index)
    s = doc.Styles.Add(name, 1)
    if base_style:
        try:
            bs = doc.Styles.Item(base_style)
        except Exception:
            bs_idx = WDBUILTINSTYLE.get(base_style)
            bs = doc.Styles.Item(bs_idx) if bs_idx else None
        if bs:
            com_set(s, "BaseStyle", bs)
    com_set_batch(s.Font, {"Name": kwargs.get("font_name"), "Size": kwargs.get("font_size"), "Bold": kwargs.get("bold"), "Italic": kwargs.get("italic")})
    return {"created": name}


def modify_style(style_name, doc_index=None, **kwargs):
    doc = get_doc(doc_index)
    try:
        s = doc.Styles.Item(style_name)
    except Exception:
        index = WDBUILTINSTYLE.get(style_name)
        if index is not None:
            s = doc.Styles.Item(index)
        else:
            return {"error": f"Style not found: {style_name}"}
    com_set_batch(s.Font, {"Name": kwargs.get("font_name"), "Size": kwargs.get("font_size"), "Bold": kwargs.get("bold"), "Italic": kwargs.get("italic")})
    com_set_batch(s.ParagraphFormat, {"Alignment": kwargs.get("alignment"), "FirstLineIndent": kwargs.get("first_line_indent"), "LineSpacingRule": kwargs.get("line_spacing_rule"), "LineSpacing": kwargs.get("line_spacing"), "SpaceBefore": kwargs.get("space_before"), "SpaceAfter": kwargs.get("space_after")})
    return {"modified": com_property(s, "NameLocal", "")}


def add_hyperlink(text: str, url: str, para_index: Optional[int] = None, doc_index: Optional[int] = None) -> Dict:
    """Insert a hyperlink at paragraph. If para_index is given, add to that paragraph; otherwise append."""
    doc = get_doc(doc_index)
    if para_index is not None:
        rng = doc.Paragraphs.Item(para_index).Range
        rng.Collapse(0)  # Collapse to start
    else:
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    doc.Hyperlinks.Add(rng, url, "", "", text)
    return {"hyperlink": True, "text": text, "url": url}


def set_tab_stops(para_index: int, stops: List[Dict], doc_index: Optional[int] = None) -> Dict:
    """Set tab stops for a paragraph.
    stops = [{"position": 400, "alignment": "right", "leader": "dot"}, ...]
    alignment: left/center/right/decimal
    leader: none/dot/dash/heavy/heavy_dash
    """
    ALIGN = {"left": 0, "center": 1, "right": 2, "decimal": 3}
    LEADER = {"none": 0, "dot": 1, "dash": 2, "heavy": 3, "heavy_dash": 4}
    doc = get_doc(doc_index)
    pf = doc.Paragraphs.Item(para_index).Format
    try:
        pf.TabStops.ClearAll()
    except Exception:
        pass
    applied = []
    for s in stops:
        try:
            ts = pf.TabStops.Add(
                s.get("position", 400),
                ALIGN.get(s.get("alignment", "left"), 0),
                LEADER.get(s.get("leader", "none"), 0)
            )
            applied.append(s)
        except Exception:
            continue
    return {"tab_stops_set": len(applied), "tab_stops": applied}


def set_bullet_list(para_indices: List[int], bullet_char: Optional[str] = None, doc_index: Optional[int] = None) -> Dict:
    """Apply bullet list formatting to paragraphs. If bullet_char not given, uses WPS default bullet."""
    doc = get_doc(doc_index)
    applied = []
    for idx in para_indices:
        try:
            pf = doc.Paragraphs.Item(idx).Format
            rng = doc.Paragraphs.Item(idx).Range
            list_format = rng.ListFormat
            list_format.ApplyBulletDefault()
            if bullet_char:
                try:
                    list_format.ListTemplate.ListLevels.Item(1).NumberFormat = bullet_char
                except Exception:
                    pass
            applied.append(idx)
        except Exception:
            continue
    return {"bullet_applied": len(applied), "para_indices": applied}


def get_run_font(para_index, run_index, doc_index=None):
    doc = get_doc(doc_index)
    valid_idx = max(para_index, 1)
    if valid_idx > doc.Paragraphs.Count:
        return {"error": f"Paragraph {para_index} out of range", "error_code": "PARAGRAPH_OUT_OF_RANGE"}
    r = doc.Paragraphs.Item(valid_idx).Range
    w = r.Words.Item(run_index)
    f = w.Font
    color_rgb = 0
    try:
        tc = com_property(f, "TextColor", None)
        if tc is not None:
            color_rgb = com_property(tc, "RGB", 0)
        else:
            color_rgb = com_property(f, "Color", 0)
    except Exception:
        try:
            color_rgb = com_property(f, "Color", 0)
        except Exception:
            color_rgb = 0
    return {
        "para_index": para_index,
        "run_index": run_index,
        "name": com_property(f, "Name", ""),
        "name_far_east": com_property(f, "NameFarEast", ""),
        "size": com_property(f, "Size", 0),
        "bold": bool(com_property(f, "Bold", 0)),
        "italic": bool(com_property(f, "Italic", 0)),
        "underline": com_property(f, "Underline", 0),
        "color_index": com_property(f, "ColorIndex", 0),
        "color_rgb": color_rgb,
        "highlight": com_property(f, "HighlightColorIndex", 0),
        "superscript": bool(com_property(f, "Superscript", 0)),
        "subscript": bool(com_property(f, "Subscript", 0)),
        "strike_through": bool(com_property(f, "StrikeThrough", 0)),
        "spacing": com_property(f, "Spacing", 0),
        "scaling": com_property(f, "Scaling", 100),
        "kerning": com_property(f, "Kerning", 0),
        "caps": bool(com_property(f, "AllCaps", 0)),
        "small_caps": bool(com_property(f, "SmallCaps", 0)),
        "emboss": bool(com_property(f, "Emboss", 0)),
        "shadow": bool(com_property(f, "Shadow", 0)),
        "outline": bool(com_property(f, "Outline", 0)),
        "vanish": bool(com_property(f, "Hidden", 0)),
    }


def set_run_font(para_index, run_index, doc_index=None, **kwargs):
    doc = get_doc(doc_index)
    if para_index < 1:
        para_index = 1
    pc = doc.Paragraphs.Count
    if para_index > pc:
        return {"error": f"Paragraph {para_index} out of range (document has {pc} paragraphs)", "error_code": "PARAGRAPH_OUT_OF_RANGE", "para_count": pc}
    try:
        r = doc.Paragraphs.Item(para_index).Range
        w = r.Words.Item(run_index)
    except Exception as e:
        return {"error": f"Failed to access Paragraph {para_index} Run {run_index}: {e}", "error_code": "COM_ACCESS_FAILED"}
    f = w.Font
    color_rgb = kwargs.get("color_rgb")
    if color_rgb is not None:
        try:
            f.TextColor.RGB = color_rgb
        except Exception:
            try:
                f.Color = color_rgb
            except Exception:
                pass
    highlight = kwargs.get("highlight")
    if highlight is not None:
        try:
            f.HighlightColorIndex = highlight
        except Exception:
            pass
    props = {}
    for key, com_name in [
        ("name", "Name"), ("name_far_east", "NameFarEast"),
        ("size", "Size"), ("bold", "Bold"), ("underline", "Underline"),
        ("color_index", "ColorIndex"), ("superscript", "Superscript"),
        ("subscript", "Subscript"), ("strike_through", "StrikeThrough"),
        ("spacing", "Spacing"), ("scaling", "Scaling"),
        ("kerning", "Kerning"), ("caps", "AllCaps"),
        ("small_caps", "SmallCaps"), ("emboss", "Emboss"),
        ("shadow", "Shadow"), ("outline", "Outline"), ("vanish", "Hidden"),
    ]:
        if kwargs.get(key) is not None:
            props[com_name] = kwargs[key]
    if kwargs.get("italic") is not None:
        props["Italic"] = kwargs["italic"]
    failed = com_set_batch(f, props)
    return {"updated": True, "para_index": para_index, "run_index": run_index, "failed": failed}


def set_text_effect(para_index: int, effect: str, color_rgb: int = 0, offset: float = 2.0, doc_index: Optional[int] = None) -> Dict:
    """Apply text effect (shadow/outline/glow/reflection/emboss/engrave)."""
    doc = get_doc(doc_index)
    if para_index < 1:
        para_index = 1
    pc = doc.Paragraphs.Count
    if para_index > pc:
        return {"error": f"Paragraph {para_index} out of range", "error_code": "PARAGRAPH_OUT_OF_RANGE", "para_count": pc}
    try:
        rng = doc.Paragraphs.Item(para_index).Range
        f = rng.Font
        effect = effect.lower()
        if effect in ("shadow", "shadowtext"):
            com_set(f, "Shadow", True)
            return {"effect": "shadow", "para_index": para_index, "applied": True}
        elif effect in ("outline", "outlinetext"):
            com_set(f, "Outline", True)
            return {"effect": "outline", "para_index": para_index, "applied": True}
        elif effect in ("emboss", "embosstext"):
            com_set(f, "Emboss", True)
            return {"effect": "emboss", "para_index": para_index, "applied": True}
        elif effect in ("engrave", "engravetext", "imprint"):
            com_set(f, "Engrave", True)
            return {"effect": "engrave", "para_index": para_index, "applied": True}
        elif effect in ("glow", "glowtext"):
            try:
                glow = f.Glow
                glow.Radius = offset
                if color_rgb:
                    glow.Color.RGB = color_rgb
                return {"effect": "glow", "para_index": para_index, "offset": offset, "applied": True}
            except Exception:
                return {"error": "Glow effect not supported by this WPS version", "error_code": "EFFECT_UNSUPPORTED"}
        elif effect in ("reflection", "reflectiontext"):
            try:
                refl = f.Reflection
                refl.Offset = offset
                if color_rgb:
                    refl.Color.RGB = color_rgb
                return {"effect": "reflection", "para_index": para_index, "offset": offset, "applied": True}
            except Exception:
                return {"error": "Reflection effect not supported by this WPS version", "error_code": "EFFECT_UNSUPPORTED"}
        else:
            return {"error": f"Unknown effect: {effect}. Valid: shadow, outline, emboss, engrave, glow, reflection", "error_code": "INVALID_PARAM"}
    except Exception as e:
        return {"error": str(e), "error_code": "COM_ERROR"}
