# -*- coding: utf-8 -*-
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set, com_set_batch, WDALIGNMENT, WDWRAPTYPE, WDNUMBERINGRULE, WDBORDERTYPE


def _sec(idx):
    return idx if idx is not None else 1


def page_setup(doc_index=None, section_index=None, **kwargs):
    doc = get_doc(doc_index)
    ps = doc.Sections.Item(_sec(section_index)).PageSetup
    failed = com_set_batch(ps, {"PageWidth": kwargs.get("page_width"), "PageHeight": kwargs.get("page_height"), "TopMargin": kwargs.get("top_margin"), "BottomMargin": kwargs.get("bottom_margin"), "LeftMargin": kwargs.get("left_margin"), "RightMargin": kwargs.get("right_margin"), "Orientation": kwargs.get("orientation"), "Gutter": kwargs.get("gutter")})
    if kwargs.get("different_first_page") is not None:
        com_set(ps, "DifferentFirstPageHeaderFooter", kwargs["different_first_page"])
    return {"page_width": com_property(ps, "PageWidth", 0), "page_height": com_property(ps, "PageHeight", 0), "top_margin": com_property(ps, "TopMargin", 0), "bottom_margin": com_property(ps, "BottomMargin", 0), "left_margin": com_property(ps, "LeftMargin", 0), "right_margin": com_property(ps, "RightMargin", 0), "orientation": com_property(ps, "Orientation", 0), "failed": failed}


def section_info(section_index=None, doc_index=None):
    doc = get_doc(doc_index)
    si = _sec(section_index)
    ps = doc.Sections.Item(si).PageSetup
    return {"index": si, "page_width": com_property(ps, "PageWidth", 0), "page_height": com_property(ps, "PageHeight", 0), "top_margin": com_property(ps, "TopMargin", 0), "bottom_margin": com_property(ps, "BottomMargin", 0), "left_margin": com_property(ps, "LeftMargin", 0), "right_margin": com_property(ps, "RightMargin", 0), "orientation": "portrait" if com_property(ps, "Orientation", 0) == 0 else "landscape", "different_first_page": bool(com_property(ps, "DifferentFirstPageHeaderFooter", 0)), "total_sections": com_property(doc.Sections, "Count", 1)}


def add_section_break(para_index, break_type="next_page", doc_index=None):
    types = {"next_page": 2, "continuous": 3, "even_page": 4, "odd_page": 5}
    get_doc(doc_index).Paragraphs.Item(para_index).Range.InsertBreak(types.get(break_type, 2))
    return {"section_break": break_type, "after_paragraph": para_index}


def set_columns(count, section_index=None, doc_index=None):
    get_doc(doc_index).Sections.Item(_sec(section_index)).PageSetup.TextColumns.SetCount(count)
    return {"columns": count}


def header_footer(section_index=None, header_type="header", text=None, doc_index=None):
    sec = get_doc(doc_index).Sections.Item(_sec(section_index))
    hf = sec.Headers(1) if header_type == "header" else sec.Footers(1)
    if text is not None: hf.Range.Text = text
    return {"type": header_type, "text": com_property(hf.Range, "Text", "").strip()}


def page_numbers(alignment="center", start_at=None, section_index=None, doc_index=None):
    sec = get_doc(doc_index).Sections.Item(_sec(section_index))
    ft = sec.Footers(1)
    ft.PageNumbers.Add(WDALIGNMENT.get(alignment, 1))
    if start_at is not None: ft.PageNumbers.StartingNumber = start_at
    return {"page_numbers": alignment, "start_at": start_at}


# DXA conversion constants (1440 DXA = 1 inch, 567 DXA ≈ 1 cm)
DXA_PER_INCH = 1440
DXA_PER_CM = 567

# Common page sizes in DXA
PAGE_SIZES_DXA = {
    "A4": {"width": 11906, "height": 16838},
    "A3": {"width": 16838, "height": 23811},
    "A5": {"width": 8392, "height": 11906},
    "Letter": {"width": 12240, "height": 15840},
    "Legal": {"width": 12240, "height": 20160},
    "B5": {"width": 10126, "height": 14388},
}


def get_page_dimensions(section_index=None, doc_index=None):
    """Get page dimensions in DXA, inches, and cm."""
    doc = get_doc(doc_index)
    ps = doc.Sections.Item(_sec(section_index)).PageSetup
    w_dxa = com_property(ps, "PageWidth", 0)
    h_dxa = com_property(ps, "PageHeight", 0)
    return {
        "width_dxa": w_dxa, "height_dxa": h_dxa,
        "width_inches": round(w_dxa / DXA_PER_INCH, 2), "height_inches": round(h_dxa / DXA_PER_INCH, 2),
        "width_cm": round(w_dxa / DXA_PER_CM, 1), "height_cm": round(h_dxa / DXA_PER_CM, 1),
        "top_margin": com_property(ps, "TopMargin", 0), "bottom_margin": com_property(ps, "BottomMargin", 0),
        "left_margin": com_property(ps, "LeftMargin", 0), "right_margin": com_property(ps, "RightMargin", 0),
        "content_width_dxa": w_dxa - com_property(ps, "LeftMargin", 0) - com_property(ps, "RightMargin", 0),
        "orientation": "portrait" if com_property(ps, "Orientation", 0) == 0 else "landscape",
    }


def insert_page_break(para_index, doc_index=None):
    doc = get_doc(doc_index)
    doc.Paragraphs.Item(para_index).Range.InsertBreak(Type=7)
    return {"page_break_before": para_index}


def set_image_wrap(shape_index, wrap_type="square", doc_index=None):
    doc = get_doc(doc_index)
    wrap_int = WDWRAPTYPE.get(wrap_type)
    result = {"shape_index": shape_index, "wrap_type": wrap_type, "found_in": None, "errors": []}
    try:
        shp = doc.Shapes.Item(shape_index)
        shp.WrapFormat.Type = wrap_int
        result["found_in"] = "shapes"
        result["wrap_set"] = wrap_type
        return result
    except Exception as e:
        result["errors"].append(f"Shapes error: {e}")
    try:
        ishp = doc.InlineShapes.Item(shape_index)
        shp = ishp.ConvertToShape()
        shp.WrapFormat.Type = wrap_int
        result["found_in"] = "inlineshapes"
        result["wrap_set"] = wrap_type
        return result
    except Exception as e:
        result["errors"].append(f"InlineShapes error: {e}")
    result["wrap_set"] = None
    return result


def set_page_border(section_index=None, doc_index=None, **kwargs):
    doc = get_doc(doc_index)
    sec = doc.Sections.Item(_sec(section_index))
    borders = sec.Borders
    current = {
        "outside_style": com_property(borders, "OutsideLineStyle", 0),
        "outside_width": com_property(borders, "OutsideLineWidth", 0),
        "outside_color_index": com_property(borders, "OutsideColorIndex", 0),
        "outside_color": com_property(borders, "OutsideColor", 0),
    }
    if "style" in kwargs:
        com_set(borders, "OutsideLineStyle", kwargs["style"])
    if "width" in kwargs:
        com_set(borders, "OutsideLineWidth", kwargs["width"])
    if "color_index" in kwargs:
        com_set(borders, "OutsideColorIndex", kwargs["color_index"])
    if "color_rgb" in kwargs:
        com_set(borders, "OutsideColor", kwargs["color_rgb"])
    if "distance_from" in kwargs:
        com_set(borders, "DistanceFrom", kwargs["distance_from"])
    if "art" in kwargs:
        com_set(borders, "OutsideLineStyle", kwargs["art"])
    return {"section": _sec(section_index), "current": current, "applied": {k: v for k, v in kwargs.items()}}


def set_line_numbers(enable=True, section_index=None, doc_index=None, **kwargs):
    doc = get_doc(doc_index)
    ps = doc.Sections.Item(_sec(section_index)).PageSetup
    ln = ps.LineNumbering
    com_set(ln, "Active", enable)
    config = {"active": enable}
    count_by = kwargs.get("count_by", 1)
    com_set(ln, "CountBy", count_by)
    config["count_by"] = count_by
    restart = kwargs.get("restart", 0)
    if isinstance(restart, str):
        restart = WDNUMBERINGRULE.get(restart, 0)
    com_set(ln, "RestartMode", restart)
    config["restart"] = restart
    if "distance" in kwargs:
        com_set(ln, "DistanceFromText", kwargs["distance"])
        config["distance"] = kwargs["distance"]
    return {"line_numbering": config}


# ─── Advanced Layout Operations ───

def set_columns_advanced(count: int, width: Optional[float] = None,
                         spacing: Optional[float] = None,
                         line_between: bool = False,
                         section_index: Optional[int] = None,
                         doc_index: Optional[int] = None) -> Dict:
    """Set columns with advanced options: custom width, spacing, and separator line."""
    sec = get_doc(doc_index).Sections.Item(section_index or 1)
    tc = sec.PageSetup.TextColumns
    tc.SetCount(count)
    if spacing is not None:
        try:
            tc.Spacing = spacing
        except Exception:
            pass
    if line_between:
        com_set(tc, "LineBetween", True)
    if width is not None:
        try:
            tc.Width = width
        except Exception:
            pass
    return {
        "columns": count,
        "width": width,
        "spacing": spacing,
        "line_between": line_between,
    }


def set_section_start(section_index: int, start_type: str,
                      doc_index: Optional[int] = None) -> Dict:
    """Set section start type.
    start_type: continuous/new_column/new_page/even_page/odd_page
    """
    types = {"continuous": 0, "new_column": 1, "new_page": 2, "even_page": 3, "odd_page": 4}
    sec = get_doc(doc_index).Sections.Item(section_index)
    com_set(sec.PageSetup, "SectionStart", types.get(start_type, 2))
    return {"section_index": section_index, "start_type": start_type}


def header_footer_link_to_previous(link: bool, header_type: str = "header",
                                    section_index: Optional[int] = None,
                                    doc_index: Optional[int] = None) -> Dict:
    """Link or unlink header/footer to previous section.
    header_type: 'header', 'footer', 'first_page_header', 'first_page_footer'
    """
    sec = get_doc(doc_index).Sections.Item(section_index or 1)
    if header_type == "first_page_header":
        hf = sec.Headers(2)
    elif header_type == "first_page_footer":
        hf = sec.Footers(2)
    elif header_type == "footer":
        hf = sec.Footers(1)
    else:
        hf = sec.Headers(1)
    com_set(hf, "LinkToPrevious", link)
    return {"section": section_index or 1, "header_type": header_type, "linked": link}


def get_header_footer_links(section_index: Optional[int] = None,
                            doc_index: Optional[int] = None) -> Dict:
    """Read header/footer linkage status for a section."""
    sec = get_doc(doc_index).Sections.Item(section_index or 1)
    result = {"section": section_index or 1}
    for name, idx in [("header", 1), ("first_page_header", 2),
                       ("footer", 1), ("first_page_footer", 2)]:
        try:
            if name.startswith("footer"):
                obj = sec.Footers(idx)
            else:
                obj = sec.Headers(idx)
            result[f"{name}_linked_to_previous"] = com_property(obj, "LinkToPrevious", False)
            result[f"{name}_exists"] = com_property(obj, "Exists", False)
        except Exception:
            continue
    return result


def page_dimensions(section_index: Optional[int] = None,
                    doc_index: Optional[int] = None) -> Dict:
    """Get page dimensions and margins with human-readable values."""
    sec = get_doc(doc_index).Sections.Item(section_index or 1)
    ps = sec.PageSetup
    return {
        "page_width_pt": com_property(ps, "PageWidth", 0),
        "page_height_pt": com_property(ps, "PageHeight", 0),
        "top_margin_pt": com_property(ps, "TopMargin", 0),
        "bottom_margin_pt": com_property(ps, "BottomMargin", 0),
        "left_margin_pt": com_property(ps, "LeftMargin", 0),
        "right_margin_pt": com_property(ps, "RightMargin", 0),
        "gutter_pt": com_property(ps, "Gutter", 0),
        "orientation": "portrait" if com_property(ps, "Orientation", 0) == 0 else "landscape",
        "printable_width_pt": com_property(ps, "PageWidth", 0) - com_property(ps, "LeftMargin", 0) - com_property(ps, "RightMargin", 0),
        "columns": com_property(ps.TextColumns, "Count", 1),
        "different_first_page": bool(com_property(ps, "DifferentFirstPageHeaderFooter", 0)),
    }


def page_break(para_index: Optional[int] = None, doc_index: Optional[int] = None) -> Dict:
    """Insert a page break before or after a paragraph."""
    doc = get_doc(doc_index)
    if para_index:
        doc.Paragraphs.Item(para_index).Range.InsertBreak(7)  # wdPageBreak
    else:
        doc.Content.InsertBreak(7)
    return {"page_break": True, "para_index": para_index}

