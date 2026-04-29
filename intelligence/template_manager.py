# -*- coding: utf-8 -*-
"""
Template Manager: extract/save/load/compare document formatting templates.
Stores templates as JSON in intelligence/templates/ directory.
"""
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from .chinese_rules import CHINESE_FORMATTING
from wps_bridge.document import doc_info
from wps_bridge.layout import section_info
from wps_bridge.table import table_info, table_read
from wps_bridge.formatting import get_font, get_paragraph_format

TEMPLATES_DIR = Path(__file__).parent / "templates"


def _ensure_dir():
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)


# ═══════════════════════════════════════════════════════════════════════════════
#  Low-level extraction helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _extract_full_paragraph(doc, i: int, doc_index: Optional[int]) -> Dict:
    """Extract every available font + paragraph property for a single paragraph."""
    from wps_bridge.utils import com_property
    from wps_bridge import formatting as fmt

    font = fmt.get_font(para_index=i, doc_index=doc_index)
    pf = fmt.get_paragraph_format(i, doc_index=doc_index)
    text = ""
    style_name = ""
    try:
        p = doc.Paragraphs.Item(i)
        r = p.Range
        text = com_property(r, "Text", "")
        style_name = com_property(r.Style, "NameLocal", "")
    except Exception:
        pass

    return {
        "index": i,
        "text": text[:120],
        "style_name": style_name,
        "font": {
            "name": font.get("name", ""),
            "name_far_east": font.get("name_far_east", ""),
            "size_pt": font.get("size", 0),
            "bold": font.get("bold", False),
            "italic": font.get("italic", False),
            "underline": font.get("underline", 0),
            "color_index": font.get("color_index", 1),
            "superscript": font.get("superscript", False),
            "subscript": font.get("subscript", False),
            "strike_through": font.get("strike_through", False),
            "spacing": font.get("spacing", 0),
            "scaling": font.get("scaling", 100),
            "kerning": font.get("kerning", 0),
        },
        "paragraph_format": {
            "alignment": pf.get("alignment", "left"),
            "alignment_raw": pf.get("alignment_raw", 0),
            "first_line_indent_pt": pf.get("first_line_indent", 0),
            "left_indent_pt": pf.get("left_indent", 0),
            "right_indent_pt": pf.get("right_indent", 0),
            "line_spacing_rule": pf.get("line_spacing_rule", "single"),
            "line_spacing_rule_raw": pf.get("line_spacing_rule_raw", 0),
            "line_spacing_pt": pf.get("line_spacing", 0),
            "space_before_pt": pf.get("space_before", 0),
            "space_after_pt": pf.get("space_after", 0),
            "outline_level": pf.get("outline_level", 10),
            "widow_control": pf.get("widow_control", False),
            "keep_with_next": pf.get("keep_with_next", False),
        },
    }


def _make_fingerprint(pdata: Dict) -> str:
    """Build a hash from font + paragraph format to identify identical formatting."""
    import hashlib
    parts = []
    f = pdata.get("font", {})
    pf = pdata.get("paragraph_format", {})
    parts.append(str(f.get("name", "")))
    parts.append(str(f.get("name_far_east", "")))
    parts.append(str(f.get("size_pt", 0)))
    parts.append(str(int(f.get("bold", False))))
    parts.append(str(int(f.get("italic", False))))
    parts.append(str(f.get("underline", 0)))
    parts.append(str(f.get("color_index", 1)))
    parts.append(str(int(f.get("superscript", False))))
    parts.append(str(int(f.get("subscript", False))))
    parts.append(str(int(f.get("strike_through", False))))
    parts.append(str(f.get("spacing", 0)))
    parts.append(str(f.get("scaling", 100)))
    parts.append(str(f.get("kerning", 0)))
    parts.append(str(pf.get("alignment_raw", 0)))
    parts.append(str(pf.get("first_line_indent_pt", 0)))
    parts.append(str(pf.get("left_indent_pt", 0)))
    parts.append(str(pf.get("right_indent_pt", 0)))
    parts.append(str(pf.get("line_spacing_rule_raw", 0)))
    parts.append(str(pf.get("line_spacing_pt", 0)))
    parts.append(str(pf.get("space_before_pt", 0)))
    parts.append(str(pf.get("space_after_pt", 0)))
    parts.append(str(pf.get("outline_level", 10)))
    parts.append(str(int(pf.get("widow_control", False))))
    parts.append(str(int(pf.get("keep_with_next", False))))
    parts.append(pdata.get("style_name", ""))
    return hashlib.md5("|".join(parts).encode()).hexdigest()[:8]


# ═══════════════════════════════════════════════════════════════════════════════
#  Section, header/footer, style gallery extraction
# ═══════════════════════════════════════════════════════════════════════════════

def _extract_all_sections(doc, doc_index: Optional[int]) -> List[Dict]:
    """Extract page setup, header, footer for EVERY section in the document."""
    from wps_bridge.utils import com_property
    sections_data = []
    total = com_property(doc.Sections, "Count", 1)
    for si in range(1, total + 1):
        try:
            sec = doc.Sections.Item(si)
            ps = sec.PageSetup
            sec_info = {
                "index": si,
                "name": f"第{si}节",
                "page_width_pt": com_property(ps, "PageWidth", 0),
                "page_height_pt": com_property(ps, "PageHeight", 0),
                "top_margin_pt": com_property(ps, "TopMargin", 0),
                "bottom_margin_pt": com_property(ps, "BottomMargin", 0),
                "left_margin_pt": com_property(ps, "LeftMargin", 0),
                "right_margin_pt": com_property(ps, "RightMargin", 0),
                "orientation": "portrait" if com_property(ps, "Orientation", 0) == 0 else "landscape",
                "different_first_page": bool(com_property(ps, "DifferentFirstPageHeaderFooter", 0)),
            }
            # Header
            try:
                hdr = sec.Headers(1)
                sec_info["header"] = {
                    "text": com_property(hdr.Range, "Text", "").strip(),
                    "font": com_property(hdr.Range.Font, "Name", ""),
                    "font_far_east": com_property(hdr.Range.Font, "NameFarEast", ""),
                    "size_pt": com_property(hdr.Range.Font, "Size", 0),
                    "bold": bool(com_property(hdr.Range.Font, "Bold", 0)),
                }
            except Exception:
                sec_info["header"] = None
            # Footer
            try:
                ftr = sec.Footers(1)
                has_page_num = False
                try:
                    has_page_num = ftr.PageNumbers.Count > 0
                except Exception:
                    pass
                sec_info["footer"] = {
                    "text": com_property(ftr.Range, "Text", "").strip(),
                    "font": com_property(ftr.Range.Font, "Name", ""),
                    "size_pt": com_property(ftr.Range.Font, "Size", 0),
                    "has_page_number": has_page_num,
                }
                if has_page_num:
                    sec_info["footer"]["page_number_alignment"] = "center"
            except Exception:
                sec_info["footer"] = None
            sections_data.append(sec_info)
        except Exception:
            continue
    return sections_data


def _extract_named_style_definitions(doc, doc_index: Optional[int]) -> Dict:
    """Extract full font + paragraph format for every named style in the document's style gallery."""
    from wps_bridge.utils import com_property, WDALIGNMENT, WDLINESPACING
    from wps_bridge import formatting as fmt

    styles_data = {}
    count = min(doc.Styles.Count, 300)
    for i in range(1, count + 1):
        try:
            s = doc.Styles.Item(i)
            name = com_property(s, "NameLocal", "")
            stype = com_property(s, "Type", 0)
            # Only paragraph styles (type 1)
            if stype != 1 or not name:
                continue
            f = s.Font
            pf = s.ParagraphFormat
            styles_data[name] = {
                "font": {
                    "name": com_property(f, "Name", ""),
                    "name_far_east": com_property(f, "NameFarEast", ""),
                    "size_pt": com_property(f, "Size", 0),
                    "bold": bool(com_property(f, "Bold", 0)),
                    "italic": bool(com_property(f, "Italic", 0)),
                    "underline": com_property(f, "Underline", 0),
                    "color_index": com_property(f, "ColorIndex", 1),
                },
                "paragraph_format": {
                    "alignment": WDALIGNMENT.get(com_property(pf, "Alignment", 0), "left"),
                    "first_line_indent_pt": com_property(pf, "FirstLineIndent", 0),
                    "left_indent_pt": com_property(pf, "LeftIndent", 0),
                    "right_indent_pt": com_property(pf, "RightIndent", 0),
                    "line_spacing_rule": WDLINESPACING.get(com_property(pf, "LineSpacingRule", 0), "single"),
                    "line_spacing_pt": com_property(pf, "LineSpacing", 0),
                    "space_before_pt": com_property(pf, "SpaceBefore", 0),
                    "space_after_pt": com_property(pf, "SpaceAfter", 0),
                    "outline_level": com_property(pf, "OutlineLevel", 10),
                },
            }
        except Exception:
            continue
    return styles_data


# ═══════════════════════════════════════════════════════════════════════════════
#  Table cell formatting extraction
# ═══════════════════════════════════════════════════════════════════════════════

def _extract_table_formatting(doc, doc_index: Optional[int]) -> List[Dict]:
    """Extract cell-level formatting for every table."""
    from wps_bridge.utils import com_property
    tables = []
    tbl_count = doc.Tables.Count
    for ti in range(1, min(tbl_count + 1, 50)):
        try:
            tbl = doc.Tables.Item(ti)
            rows = tbl.Rows.Count
            cols = tbl.Columns.Count
            tinfo = {"index": ti, "rows": rows, "cols": cols}

            # Sample up to 3 rows (first, middle, last)
            sample_rows = sorted(set([
                1,
                min(2, rows),
                max(1, rows // 2),
                max(1, rows - 1),
                rows,
            ]))
            sample_rows = [r for r in sample_rows if 1 <= r <= rows][:5]
            cells_sample = []
            for r in sample_rows:
                for c in range(1, min(cols + 1, 5)):
                    try:
                        cell = tbl.Cell(r, c)
                        cr = cell.Range
                        cf = cr.Font
                        text = com_property(cr, "Text", "").replace("\r\x07", "").strip()[:60]
                        cells_sample.append({
                            "row": r, "col": c,
                            "text": text,
                            "font": {
                                "name": com_property(cf, "Name", ""),
                                "name_far_east": com_property(cf, "NameFarEast", ""),
                                "size_pt": com_property(cf, "Size", 0),
                                "bold": bool(com_property(cf, "Bold", 0)),
                            },
                        })
                    except Exception:
                        continue
            tinfo["cells_sample"] = cells_sample

            # Detect border style by checking LineStyle of first cell outer borders
            try:
                import win32com.client
                wd = win32com.client.constants
                border_style = None
                for bn in ["wdBorderTop", "wdBorderBottom"]:
                    try:
                        b = tbl.Cell(1, 1).Borders(getattr(wd, bn))
                        border_style = com_property(b, "LineStyle", 0)
                        if border_style:
                            break
                    except Exception:
                        pass
                tinfo["border_line_style"] = border_style
            except Exception:
                pass

            tables.append(tinfo)
        except Exception:
            continue
    return tables


# ═══════════════════════════════════════════════════════════════════════════════
#  Caption detection (图/表/Figure/Table)
# ═══════════════════════════════════════════════════════════════════════════════

_CAPTION_RE = re.compile(r'^[\s　]*(图|表|Figure|Table)\s*(\d+|[一二三四五六七八九十百千]+)')

def _detect_captions(paragraphs_data: List[Dict]) -> Dict:
    """Detect figure/table captions and extract their formatting."""
    figure_captions = []
    table_captions = []
    for pd in paragraphs_data:
        text = pd.get("text", "")
        m = _CAPTION_RE.match(text)
        if not m:
            continue
        prefix = m.group(1)
        entry = {
            "para_index": pd["index"],
            "text": text[:100],
            "font": dict(pd["font"]),
            "paragraph_format": dict(pd["paragraph_format"]),
            "style_name": pd["style_name"],
        }
        if prefix in ("图", "Figure"):
            figure_captions.append(entry)
        else:
            table_captions.append(entry)

    result = {}
    if figure_captions:
        sample = figure_captions[0]
        result["figure_caption"] = {
            "count": len(figure_captions),
            "font": sample["font"],
            "paragraph_format": sample["paragraph_format"],
            "style_name": sample["style_name"],
            "examples": [c["text"] for c in figure_captions[:5]],
        }
    if table_captions:
        sample = table_captions[0]
        result["table_caption"] = {
            "count": len(table_captions),
            "font": sample["font"],
            "paragraph_format": sample["paragraph_format"],
            "style_name": sample["style_name"],
            "examples": [c["text"] for c in table_captions[:5]],
        }
    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  Semantic section detection
# ═══════════════════════════════════════════════════════════════════════════════

_SEMANTIC_PATTERNS = [
    ("abstract_cn", re.compile(r'(摘[　\s]*要|中文摘要)')),
    ("abstract_en", re.compile(r'(ABSTRACT|Abstract|英文摘要)')),
    ("toc", re.compile(r'(目[　\s]*录|Table\s*of\s*Contents|TOC)')),
    ("references", re.compile(r'(参考[文资]|REFERENCE|Reference|参考文献|Bibliography)')),
    ("acknowledgements", re.compile(r'(致[　\s]*谢|ACKNOWLEDGEMENT|Acknowledgement|谢辞)')),
    ("conclusion", re.compile(r'(结束语|结[　\s]*论|Conclusion|总[　\s]*结)')),
    ("appendix", re.compile(r'(附[　\s]*录|Appendix|Appendices)')),
]

def _detect_semantic_sections(paragraphs_data: List[Dict]) -> Dict:
    """Detect semantic document sections by heading text patterns."""
    semantic = {}
    for i, pd in enumerate(paragraphs_data):
        ol = pd["paragraph_format"].get("outline_level", 10)
        if ol > 9:
            continue
        text = pd.get("text", "").strip()
        for key, pattern in _SEMANTIC_PATTERNS:
            if pattern.match(text) and key not in semantic:
                semantic[key] = {
                    "para_index": pd["index"],
                    "heading_text": text,
                    "heading_font": dict(pd["font"]),
                    "heading_paragraph_format": dict(pd["paragraph_format"]),
                    "heading_style_name": pd["style_name"],
                }
                break

    # Add body text section headers (all non-semantic headings)
    body_headings = []
    for pd in paragraphs_data:
        ol = pd["paragraph_format"].get("outline_level", 10)
        if 1 <= ol <= 9:
            text = pd.get("text", "").strip()
            if not any(pattern.match(text) for _, pattern in _SEMANTIC_PATTERNS):
                body_headings.append(pd)

    return semantic


# ═══════════════════════════════════════════════════════════════════════════════
#  Summary generation
# ═══════════════════════════════════════════════════════════════════════════════

def _build_summary(extracted: Dict) -> Dict:
    """Build a natural-language summary of key formatting rules."""
    s = []
    body = extracted.get("body_rules", {})
    if body:
        font = body.get("font", {})
        pf = body.get("paragraph_format", {})
        cn = font.get("name_far_east", font.get("name", ""))
        en = font.get("name", "")
        lang = f"中文: {cn}" if cn else ""
        if en and en != cn:
            lang += f" / 英文/数字: {en}"
        if lang:
            s.append(lang)
        size = font.get("size_pt", 0)
        if size:
            s.append(f"正文字号: {size}pt")
        if pf.get("line_spacing_rule") and pf.get("line_spacing_pt"):
            s.append(f"行距: {pf['line_spacing_rule']} {pf['line_spacing_pt']}pt")
        indent = pf.get("first_line_indent_pt", 0)
        if indent:
            s.append(f"首行缩进: {indent}pt")

    headings = extracted.get("outline_rules", {})
    for level_key in sorted(headings):
        rule = headings[level_key]
        font = rule.get("font", {})
        hn = font.get("name_far_east", font.get("name", ""))
        hs = font.get("size_pt", 0)
        hb = "加粗" if font.get("bold") else ""
        s.append(f"{level_key}: {hn} {hs}pt {hb}".strip())

    cover = extracted.get("cover", [])
    if cover:
        s.append(f"封面: {len(cover)} 个元素")

    captions = extracted.get("captions", {})
    for key in ("figure_caption", "table_caption"):
        if key in captions:
            cap = captions[key]
            s.append(f"{key}: {cap['font'].get('name_far_east', '')} {cap['font'].get('size_pt', 0)}pt ×{cap['count']}")

    tables = extracted.get("tables", [])
    if tables:
        s.append(f"表格: {len(tables)} 个")

    return {"summary_text": "；".join(s)}


# ═══════════════════════════════════════════════════════════════════════════════
#  Main extract() — comprehensive
# ═══════════════════════════════════════════════════════════════════════════════

def extract(doc_index: Optional[int] = None) -> Dict:
    """Extract every formatting detail from a document into a comprehensive template JSON.

    Scans ALL paragraphs, ALL sections, ALL named styles, ALL tables.
    Returns an output matching the detail level of human-curated templates
    (e.g. 湖北经济学院论文排版规范.json).
    """
    from wps_bridge.app import get_doc

    info = doc_info(doc_index)
    doc = get_doc(doc_index)
    para_count = doc.Paragraphs.Count

    # ── 1. All sections ──
    all_sections = _extract_all_sections(doc, doc_index)

    # ── 2. Named style definitions from gallery ──
    named_styles = _extract_named_style_definitions(doc, doc_index)

    # ── 3. Full paragraph scan ──
    all_paragraphs = []
    for i in range(1, para_count + 1):
        try:
            pdata = _extract_full_paragraph(doc, i, doc_index)
            all_paragraphs.append(pdata)
        except Exception:
            continue

    if not all_paragraphs:
        return {"error": "No paragraphs found", "doc_name": info.get("name", "")}

    # ── 4. Distinct formatting patterns ──
    pattern_groups = {}
    for pdata in all_paragraphs:
        fp = _make_fingerprint(pdata)
        if fp not in pattern_groups:
            pattern_groups[fp] = []
        pattern_groups[fp].append(pdata)

    distinct_patterns = []
    for fp, group in sorted(pattern_groups.items(), key=lambda x: -len(x[1])):
        sample = group[0]
        distinct_patterns.append({
            "fingerprint": fp,
            "count": len(group),
            "percentage": round(len(group) / len(all_paragraphs) * 100, 1),
            "font": sample["font"],
            "paragraph_format": sample["paragraph_format"],
            "style_name": sample["style_name"],
            "para_indices": [p["index"] for p in group[:20]],
            "examples": [p["text"] for p in group[:5] if p["text"].strip()],
        })

    # ── 5. Outline-level heading rules ──
    level_groups = {}
    for pdata in all_paragraphs:
        ol = pdata["paragraph_format"]["outline_level"]
        if 1 <= ol <= 9:
            if ol not in level_groups:
                level_groups[ol] = []
            level_groups[ol].append(pdata)

    outline_rules = {}
    for level in sorted(level_groups):
        sample = level_groups[level][0]
        outline_rules[f"heading_level_{level}"] = {
            "count": len(level_groups[level]),
            "font": dict(sample["font"]),
            "paragraph_format": dict(sample["paragraph_format"]),
            "style_name": sample["style_name"],
            "examples": [p["text"] for p in level_groups[level][:5]],
        }

    # ── 6. Body text rules ──
    body_rules = {}
    body_found = None
    for pdata in all_paragraphs:
        if pdata["paragraph_format"]["outline_level"] > 9 and pdata["text"].strip():
            body_found = pdata
            break
    if body_found:
        body_rules = {
            "font": dict(body_found["font"]),
            "paragraph_format": dict(body_found["paragraph_format"]),
            "style_name": body_found["style_name"],
            "note": f"取自第{body_found['index']}段",
        }

    # ── 7. Cover page detection ──
    cover_paragraphs = []
    first_section_end = para_count
    if len(all_sections) > 1:
        try:
            first_section_end = doc.Sections.Item(2).Range.Start
            # Find last paragraph before section 2
            for pd in all_paragraphs:
                p = doc.Paragraphs.Item(pd["index"])
                if hasattr(p, "Range"):
                    try:
                        if p.Range.Start >= first_section_end:
                            break
                    except Exception:
                        pass
                cover_paragraphs.append(pd)
            # Fallback: use paragraphs before first heading if section-based fails
            if not cover_paragraphs:
                for pdata in all_paragraphs:
                    if 1 <= pdata["paragraph_format"]["outline_level"] <= 9:
                        break
                    if pdata["text"].strip():
                        cover_paragraphs.append({
                            "index": pdata["index"],
                            "font": dict(pdata["font"]),
                            "paragraph_format": dict(pdata["paragraph_format"]),
                            "style_name": pdata["style_name"],
                            "text": pdata["text"],
                        })
        except Exception:
            cover_paragraphs = []
    else:
        # Single section: paragraphs before first heading
        for pdata in all_paragraphs:
            if 1 <= pdata["paragraph_format"]["outline_level"] <= 9:
                break
            if pdata["text"].strip():
                cover_paragraphs.append({
                    "index": pdata["index"],
                    "font": dict(pdata["font"]),
                    "paragraph_format": dict(pdata["paragraph_format"]),
                    "style_name": pdata["style_name"],
                    "text": pdata["text"],
                })
                if len(cover_paragraphs) >= 15:
                    break

    is_cover_detected = bool(cover_paragraphs) and (
        cover_paragraphs[0]["paragraph_format"].get("alignment", "") == "center"
        or len(cover_paragraphs) >= 3
    )

    cover_elements = []
    if is_cover_detected and cover_paragraphs:
        for cp in cover_paragraphs:
            cover_elements.append({
                "index": cp["index"],
                "font": cp["font"],
                "paragraph_format": cp["paragraph_format"],
                "style_name": cp["style_name"],
                "text": cp["text"],
            })

    # ── 8. Named style usage ──
    style_counter = {}
    for pdata in all_paragraphs:
        sn = pdata["style_name"]
        if sn:
            style_counter[sn] = style_counter.get(sn, 0) + 1
    styles_used = [
        {"name": name, "count": count, "pct": round(count / len(all_paragraphs) * 100, 1)}
        for name, count in sorted(style_counter.items(), key=lambda x: -x[1])
    ]

    # ── 9. Table formatting ──
    tables_data = _extract_table_formatting(doc, doc_index)

    # ── 10. Figure/Table caption detection ──
    captions_data = _detect_captions(all_paragraphs)

    # ── 11. Semantic section detection ──
    semantic_sections = _detect_semantic_sections(all_paragraphs)

    # ── 12. Backward-compatible rules ──
    legacy_rules = _build_legacy_rules(all_sections, outline_rules, distinct_patterns, cover_paragraphs)

    # ── 13. Build paper section ──
    paper = {
        "size": "A4",
        "sections": all_sections,
    }
    # Extract header/footer from first section as top-level convenience
    if all_sections:
        paper["header"] = all_sections[0].get("header")
        paper["footer"] = all_sections[0].get("footer")

    # ── 14. Build result ──
    result = {
        "name": f"Extracted from {info.get('name', 'unknown')}",
        "description": f"Auto-extracted from {info.get('name', 'unknown')} — all paragraph/section/style/table formatting",
        "doc_type": "general",
        "extracted_from": info.get("full_name", info.get("name", "")),
        "extracted_at": None,

        "paper": paper,

        "styles": named_styles,

        "heading_rules": outline_rules,
        "body_rules": body_rules,

        "cover_page": {
            "has_cover": is_cover_detected,
            "elements": cover_elements,
            "section_index": 1 if is_cover_detected and len(all_sections) > 0 else None,
        },

        "semantic_sections": semantic_sections,
        "captions": captions_data,

        "tables": tables_data,

        "distinct_patterns": distinct_patterns,
        "styles_used": styles_used,
        "summary": _build_summary({
            "cover": cover_elements,
            "outline_rules": outline_rules,
            "body_rules": body_rules,
            "tables": tables_data,
            "captions": captions_data,
        }),

        # backward compatible
        "rules": legacy_rules,
        "_meta": {
            "total_paragraphs": para_count,
            "total_paragraphs_scanned": len(all_paragraphs),
            "distinct_patterns_count": len(distinct_patterns),
            "outline_levels": sorted(level_groups.keys()),
            "styles_defined_in_gallery": len(named_styles),
            "table_count": doc.Tables.Count,
        },
    }

    return result


def _build_legacy_rules(all_sections, outline_rules, distinct_patterns, cover_paragraphs) -> Dict:
    """Build the old-style 'rules' dict for backward compatibility with apply_template."""
    rules = {}
    if all_sections:
        sec0 = all_sections[0]
        rules["page"] = {
            "paper": "A4",
            "top_margin": sec0.get("top_margin_pt", 25.4),
            "bottom_margin": sec0.get("bottom_margin_pt", 25.4),
            "left_margin": sec0.get("left_margin_pt", 31.7),
            "right_margin": sec0.get("right_margin_pt", 31.7),
        }
    else:
        rules["page"] = {"paper": "A4"}

    for label, rule in outline_rules.items():
        font = rule.get("font", {})
        pf = rule.get("paragraph_format", {})
        rules[label] = {
            "font_name": font.get("name_far_east", font.get("name", "宋体")),
            "font_size": font.get("size_pt", 12),
            "bold": font.get("bold", False),
            "italic": font.get("italic", False),
            "underline": font.get("underline", 0),
            "alignment": pf.get("alignment", "left"),
            "outline_level": pf.get("outline_level", 10),
            "space_before": pf.get("space_before_pt", 0),
            "space_after": pf.get("space_after_pt", 0),
            "line_spacing_rule": pf.get("line_spacing_rule", "single"),
            "line_spacing": pf.get("line_spacing_pt", 1),
            "first_line_indent": pf.get("first_line_indent_pt", 0),
            "left_indent": pf.get("left_indent_pt", 0),
            "right_indent": pf.get("right_indent_pt", 0),
        }

    body_pattern = None
    for pat in distinct_patterns:
        if pat["paragraph_format"]["outline_level"] > 9 and pat["font"].get("size_pt", 0) > 0:
            body_pattern = pat
            break
    if body_pattern:
        font = body_pattern.get("font", {})
        pf = body_pattern.get("paragraph_format", {})
        rules["正文"] = {
            "font_name": font.get("name_far_east", font.get("name", "宋体")),
            "font_size": font.get("size_pt", 12),
            "bold": font.get("bold", False),
            "italic": font.get("italic", False),
            "underline": font.get("underline", 0),
            "alignment": pf.get("alignment", "justify"),
            "first_line_indent": pf.get("first_line_indent_pt", 28),
            "line_spacing_rule": pf.get("line_spacing_rule", "multiple"),
            "line_spacing": pf.get("line_spacing_pt", 1.5),
            "space_before": pf.get("space_before_pt", 0),
            "space_after": pf.get("space_after_pt", 0),
        }

    if cover_paragraphs:
        cp = cover_paragraphs[0]
        font = cp.get("font", {})
        pf = cp.get("paragraph_format", {})
        rules["封面标题"] = {
            "font_name": font.get("name_far_east", font.get("name", "黑体")),
            "font_size": font.get("size_pt", 22),
            "bold": font.get("bold", True),
            "alignment": pf.get("alignment", "center"),
            "space_before": pf.get("space_before_pt", 0),
            "space_after": pf.get("space_after_pt", 12),
            "is_cover": True,
        }

    return rules


# ═══════════════════════════════════════════════════════════════════════════════
#  Save / Load / List / Delete / Export / Import / Compare
# ═══════════════════════════════════════════════════════════════════════════════

def save(template_name: str, template_data: Dict) -> Dict:
    """Save a template as JSON file."""
    _ensure_dir()
    from datetime import datetime
    template_data["extracted_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    filepath = TEMPLATES_DIR / f"{template_name}.json"
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(template_data, f, ensure_ascii=False, indent=2)
    return {"saved": template_name, "path": str(filepath)}


def load(template_name: str) -> Optional[Dict]:
    """Load a saved template by name."""
    _ensure_dir()
    filepath = TEMPLATES_DIR / f"{template_name}.json"
    if filepath.exists():
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def list_all() -> List[Dict]:
    """List all saved custom templates."""
    _ensure_dir()
    result = []
    for f in sorted(TEMPLATES_DIR.glob("*.json")):
        try:
            with open(f, "r", encoding="utf-8") as fp:
                data = json.load(fp)
                result.append({
                    "name": f.stem,
                    "description": data.get("description", ""),
                    "extracted_from": data.get("extracted_from", ""),
                    "extracted_at": data.get("extracted_at", ""),
                })
        except Exception:
            continue
    for name in CHINESE_FORMATTING:
        result.append({"name": name, "type": "builtin", "description": f"Built-in: {name}"})
    return result


def delete(template_name: str) -> Dict:
    """Delete a saved template."""
    filepath = TEMPLATES_DIR / f"{template_name}.json"
    if filepath.exists():
        filepath.unlink()
        return {"deleted": template_name}
    return {"error": f"Template not found: {template_name}"}


def export_template(template_name: str, filepath: str) -> Dict:
    """Export a template to an external JSON file."""
    data = load(template_name)
    if not data:
        if template_name in CHINESE_FORMATTING:
            data = {
                "name": template_name,
                "doc_type": template_name,
                "rules": CHINESE_FORMATTING[template_name],
            }
        else:
            return {"error": f"Template not found: {template_name}"}
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return {"exported": template_name, "path": filepath}


def import_template(filepath: str) -> Dict:
    """Import a template from a JSON file."""
    _ensure_dir()
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    name = data.get("name", Path(filepath).stem)
    filepath_dest = TEMPLATES_DIR / f"{name}.json"
    with open(filepath_dest, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return {"imported": name, "path": str(filepath_dest)}


def compare_with_template(doc_index: Optional[int] = None, template_name: str = "") -> Dict:
    """Compare current document formatting against a template. Uses detailed extraction."""
    extracted = extract(doc_index)
    doc_rules = extracted.get("rules", {})

    cmp_rules = None
    saved = load(template_name)
    if saved:
        cmp_rules = saved.get("rules", {})
    elif template_name in CHINESE_FORMATTING:
        cmp_rules = CHINESE_FORMATTING[template_name]
    else:
        return {"error": f"Template not found: {template_name}", "available_builtins": list(CHINESE_FORMATTING.keys())}

    mismatches = []
    compare_props = [
        ("font_name", "字体"), ("font_size", "字号"), ("bold", "加粗"), ("italic", "斜体"),
        ("underline", "下划线"), ("alignment", "对齐"),
        ("first_line_indent", "首行缩进"), ("left_indent", "左缩进"), ("right_indent", "右缩进"),
        ("line_spacing_rule", "行距规则"), ("line_spacing", "行距值"),
        ("space_before", "段前间距"), ("space_after", "段后间距"),
    ]

    for key, expected in cmp_rules.items():
        if key == "page":
            continue
        actual = doc_rules.get(key, {})
        if not actual:
            mismatches.append({"rule": key, "issue": "文档中未找到该规则"})
            continue
        for prop_en, prop_cn in compare_props:
            exp_val = expected.get(prop_en)
            act_val = actual.get(prop_en)
            if exp_val is not None and act_val is not None and exp_val != act_val:
                mismatches.append({
                    "rule": key, "property": prop_en, "property_cn": prop_cn,
                    "expected": exp_val, "actual": act_val,
                })

    return {
        "template": template_name,
        "doc_rules_count": len(doc_rules),
        "template_rules_count": len(cmp_rules),
        "mismatches": len(mismatches),
        "details": mismatches[:100],
    }
