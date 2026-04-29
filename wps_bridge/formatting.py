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
    failed = com_set_batch(f, {"Name": kwargs.get("name"), "NameFarEast": kwargs.get("name_far_east"), "Size": kwargs.get("size"), "Bold": kwargs.get("bold"), "Italic": kwargs.get("italic"), "Underline": kwargs.get("underline"), "ColorIndex": kwargs.get("color_index"), "Superscript": kwargs.get("superscript"), "Subscript": kwargs.get("subscript"), "StrikeThrough": kwargs.get("strike_through"), "Spacing": kwargs.get("spacing"), "Scaling": kwargs.get("scaling"), "Kerning": kwargs.get("kerning")})
    return {"updated": True, "failed": failed, "text_sample": com_property(r, "Text", "")[:50]}


def get_paragraph_format(para_index, doc_index=None):
    doc = get_doc(doc_index)
    p = doc.Paragraphs.Item(para_index)
    pf = p.Format
    return {"alignment": WDALIGNMENT.get(com_property(pf, "Alignment", 0), "unknown"), "alignment_raw": com_property(pf, "Alignment", 0), "first_line_indent": com_property(pf, "FirstLineIndent", 0), "left_indent": com_property(pf, "LeftIndent", 0), "right_indent": com_property(pf, "RightIndent", 0), "line_spacing_rule": WDLINESPACING.get(com_property(pf, "LineSpacingRule", 0), "unknown"), "line_spacing_rule_raw": com_property(pf, "LineSpacingRule", 0), "line_spacing": com_property(pf, "LineSpacing", 0), "space_before": com_property(pf, "SpaceBefore", 0), "space_after": com_property(pf, "SpaceAfter", 0), "outline_level": com_property(pf, "OutlineLevel", 10), "widow_control": bool(com_property(pf, "WidowControl", 0)), "keep_with_next": bool(com_property(pf, "KeepWithNext", 0))}


def set_paragraph_format(para_index=None, use_selection=False, doc_index=None, **kwargs):
    doc = get_doc(doc_index)
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
    failed = com_set_batch(pf, {"Alignment": align, "FirstLineIndent": kwargs.get("first_line_indent"), "LeftIndent": kwargs.get("left_indent"), "RightIndent": kwargs.get("right_indent"), "LineSpacingRule": lsr, "LineSpacing": kwargs.get("line_spacing"), "SpaceBefore": kwargs.get("space_before"), "SpaceAfter": kwargs.get("space_after"), "OutlineLevel": kwargs.get("outline_level"), "WidowControl": kwargs.get("widow_control"), "KeepWithNext": kwargs.get("keep_with_next")})
    return {"updated": True, "failed": failed}


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
