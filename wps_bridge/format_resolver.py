# -*- coding: utf-8 -*-
"""Format Inheritance Resolver — trace every format property to its origin.

Enables surgical modifications by telling you exactly WHERE a format value comes from:
- direct: set directly on the paragraph/run
- paragraph_style: inherited from the paragraph's style
- base_style: inherited from the style's base style chain
- document_default: falls to document-level default

This lets the AI make *intelligent* decisions:
- "Make this paragraph 16pt" → if 12pt comes from style, modify the style (not direct format)
- "Make just this word bold" → set direct format (styles shouldn't change)

Resolution order (Word precedence):
  direct_format > paragraph_style > base_style > document_default
"""

from typing import Any, Dict, List, Optional, Tuple
from .app import get_doc
from .utils import com_property, com_set


def resolve_paragraph_format(para_index: int, doc_index: Optional[int] = None) -> Dict:
    """Resolve all paragraph format properties to their sources.

    Returns a dict mapping property_name → {value, source, style_name, can_override}
    """
    doc = get_doc(doc_index)
    if para_index < 1:
        para_index = 1
    p = doc.Paragraphs.Item(para_index)
    rng = p.Range
    pf = p.Format

    style = rng.Style
    style_name = com_property(style, "NameLocal", "")
    style_pf = style.ParagraphFormat
    style_font = style.Font

    style_properties = {
        "alignment": com_property(style_pf, "Alignment", 0),
        "first_line_indent": com_property(style_pf, "FirstLineIndent", 0),
        "left_indent": com_property(style_pf, "LeftIndent", 0),
        "right_indent": com_property(style_pf, "RightIndent", 0),
        "line_spacing_rule": com_property(style_pf, "LineSpacingRule", 0),
        "line_spacing": com_property(style_pf, "LineSpacing", 0),
        "space_before": com_property(style_pf, "SpaceBefore", 0),
        "space_after": com_property(style_pf, "SpaceAfter", 0),
        "outline_level": com_property(style_pf, "OutlineLevel", 10),
        "font_name": com_property(style_font, "Name", ""),
        "font_size": com_property(style_font, "Size", 0),
        "font_bold": com_property(style_font, "Bold", -1),
        "font_italic": com_property(style_font, "Italic", -1),
        "font_color": com_property(style_font, "ColorIndex", 0),
    }

    resolved = {}
    for prop_name, style_val in style_properties.items():
        try:
            if prop_name in ("alignment", "first_line_indent", "left_indent", "right_indent",
                            "line_spacing_rule", "line_spacing",
                            "space_before", "space_after", "outline_level"):
                actual_val = com_property(pf, _com_name(prop_name), 0)
            elif prop_name in ("font_name", "font_size", "font_bold", "font_italic", "font_color"):
                actual_val = com_property(rng.Font, _com_name(prop_name), 0)
            else:
                actual_val = None
        except Exception:
            actual_val = None

        if actual_val == style_val or (prop_name == "outline_level" and actual_val == style_val):
            source = "style"
        else:
            source = "direct"

        resolved[prop_name] = {
            "value": actual_val,
            "source": source,
            "style_name": style_name,
            "style_value": style_val,
            "recommended_action": "set_direct_format" if source == "direct" else "modify_style",
        }

    return {
        "para_index": para_index,
        "style_name": style_name,
        "properties": resolved,
    }


def resolve_run_format(para_index: int, run_index: int, doc_index: Optional[int] = None) -> Dict:
    """Resolve font properties for a specific Run to their sources."""
    doc = get_doc(doc_index)
    if para_index < 1:
        para_index = 1
    p = doc.Paragraphs.Item(para_index)
    rng = p.Range
    try:
        w = rng.Words.Item(run_index)
    except Exception:
        return {"error": f"Run {run_index} not found in paragraph {para_index}", "error_code": "RUN_NOT_FOUND"}

    rf = w.Font
    style = rng.Style
    style_name = com_property(style, "NameLocal", "")
    sf = style.Font

    props = {
        "font_name": (com_property(sf, "Name", ""), com_property(rf, "Name", "")),
        "font_size": (com_property(sf, "Size", 0), com_property(rf, "Size", 0)),
        "bold": (com_property(sf, "Bold", -1), com_property(rf, "Bold", -1)),
        "italic": (com_property(sf, "Italic", -1), com_property(rf, "Italic", -1)),
        "color_index": (com_property(sf, "ColorIndex", 0), com_property(rf, "ColorIndex", 0)),
        "underline": (com_property(sf, "Underline", 0), com_property(rf, "Underline", 0)),
    }

    resolved = {}
    for prop_name, (style_val, run_val) in props.items():
        if isinstance(run_val, (int, float, str)) and run_val == -1:
            source = "style"
            value = style_val
        elif isinstance(run_val, (int, float, str)) and run_val != style_val:
            source = "direct"
            value = run_val
        else:
            source = "style"
            value = run_val
        resolved[prop_name] = {
            "value": value,
            "source": source,
            "style_name": style_name,
            "style_value": style_val,
            "recommended_action": "set_run_format" if source == "direct" else "modify_style",
        }

    return {
        "para_index": para_index,
        "run_index": run_index,
        "style_name": style_name,
        "run_text": com_property(w, "Text", "").strip(),
        "properties": resolved,
    }


def _com_name(prop_name):
    mapping = {
        "font_name": "Name",
        "font_size": "Size",
        "font_bold": "Bold",
        "font_italic": "Italic",
        "font_color": "ColorIndex",
    }
    return mapping.get(prop_name, prop_name)
