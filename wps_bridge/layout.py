# -*- coding: utf-8 -*-
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set, com_set_batch, WDALIGNMENT


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
