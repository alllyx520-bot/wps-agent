# -*- coding: utf-8 -*-
"""
Incremental Serializer — applies Document model changes to existing OOXML XML trees.
Uses _xml_element references on Paragraph/Run/Cell/Table/Section to modify XML
in-place, preserving ALL unmapped OOXML attributes for lossless round-trip.
"""
import os
import shutil
from pathlib import Path
from typing import Dict, Any, Optional, Set, List
from lxml import etree

from .document_model import Document, Paragraph, Run, Table, Cell, Section
from .xml_parser import (
    unpack_docx, pack_docx, parse_xml, parse_docx,
    _q, W,
    get_tables, get_table_rows, get_table_cells,
)
from .errors import SerializeError, ErrorCode

R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


# ═══════════════════════════════════════════════════════════════════════
# Construction (from XML → model): uses + extends old build_document_model
# ═══════════════════════════════════════════════════════════════════════

def build_document_model(parsed: Dict[str, Any]) -> Document:
    """Build a Document DOM from parsed XML structures (full property coverage)."""
    doc_root = parsed["document"]
    styles_root = parsed.get("styles")
    numbering_root = parsed.get("numbering")
    rels_root = parsed.get("rels")

    doc = Document()
    doc._unpacked_dir = parsed.get("unpacked_dir")
    doc.media_files = parsed.get("media_files", [])

    if rels_root is not None:
        for rel in rels_root.findall(f"{{{R[1:-1]}}}Relationship"):
            rid = rel.get("Id", "")
            rel_type = rel.get("Type", "")
            target = rel.get("Target", "")
            if "hyperlink" in rel_type:
                doc.hyperlinks[rid] = target

    from .style_resolver import StyleResolver
    if styles_root is not None:
        resolver = StyleResolver(styles_root)
        doc.styles = resolver.styles

    if numbering_root is not None:
        doc.numbering = {"root": numbering_root}

    # Parse footnotes
    if parsed.get("footnotes") is not None:
        fn_root = parsed["footnotes"]
        doc.footnotes = _parse_footnotes(fn_root)

    # Parse endnotes
    if parsed.get("endnotes") is not None:
        en_root = parsed["endnotes"]
        doc.endnotes = _parse_footnotes(en_root)

    # Parse headers/footers
    doc.headers = parsed.get("headers", {})
    doc.footers = parsed.get("footers", {})

    # Parse settings
    if parsed.get("settings") is not None:
        doc.settings = {"root": parsed["settings"]}

    # Parse font table
    if parsed.get("fontTable") is not None:
        doc.font_table = {"root": parsed["fontTable"]}

    # Parse theme
    doc.theme = parsed.get("theme")

    # Parse body content
    body = doc_root.find(_q("body"))
    if body is None:
        return doc

    for child in body:
        tag = etree.QName(child).localname
        if tag == "p":
            para = _parse_paragraph_element_full(child, doc.hyperlinks)
            doc.paragraphs.append(para)
        elif tag == "tbl":
            tbl = _parse_table_element_full(child, doc.hyperlinks)
            doc.tables.append(tbl)
        elif tag == "sectPr":
            sec = _parse_section_properties_full(child)
            sec._xml_element = child
            doc.sections.append(sec)

    return doc


def _parse_footnotes(root: etree._Element) -> List[Dict]:
    """Parse footnote or endnote XML into a list of dicts."""
    notes = []
    for fn in root.findall(_q("footnote")) or root.findall(_q("endnote")):
        note_id = fn.get(_q("id"), "")
        text_parts = []
        for t in fn.findall(f".//{_q('t')}"):
            text_parts.append(t.text or "")
        notes.append({"id": note_id, "text": "".join(text_parts)})
    return notes


# ═══════════════════════════════════════════════════════════════════════
# Full property parsing (Run, Paragraph, Table, Section)
# ═══════════════════════════════════════════════════════════════════════

def _parse_paragraph_element_full(p_elem: etree._Element, hyperlinks: Dict[str, str]) -> Paragraph:
    """Parse a <w:p> element into a Paragraph with ALL properties."""
    para = Paragraph()
    para._xml_element = p_elem

    ppr = p_elem.find(_q("pPr"))
    if ppr is not None:
        para.style_id = _attr_val(ppr, "pStyle", "val")

        jc = ppr.find(_q("jc"))
        if jc is not None:
            para.alignment = jc.get(_q("val"), "left")

        ind = ppr.find(_q("ind"))
        if ind is not None:
            para.first_line_indent = _twips_to_pt(ind.get(_q("firstLine")))
            para.left_indent = _twips_to_pt(ind.get(_q("left")))
            para.right_indent = _twips_to_pt(ind.get(_q("right")))
            para.hanging = _twips_to_pt(ind.get(_q("hanging")))

        spacing = ppr.find(_q("spacing"))
        if spacing is not None:
            para.space_before = _twips_to_pt(spacing.get(_q("before")))
            para.space_after = _twips_to_pt(spacing.get(_q("after")))
            line = spacing.get(_q("line"))
            line_rule = spacing.get(_q("lineRule"))
            if line:
                if line_rule == "auto":
                    para.line_spacing = int(line) / 240
                else:
                    para.line_spacing = _twips_to_pt(line)
            para.line_rule = line_rule

        para.keep_with_next = ppr.find(_q("keepNext")) is not None
        para.keep_lines = ppr.find(_q("keepLines")) is not None
        para.page_break_before = ppr.find(_q("pageBreakBefore")) is not None

        wc = ppr.find(_q("widowControl"))
        if wc is not None:
            para.widow_control = wc.get(_q("val"), "true") != "false"

        outline = ppr.find(_q("outlineLvl"))
        if outline is not None:
            para.outline_level = int(outline.get(_q("val"), "9"))

        numpr = ppr.find(_q("numPr"))
        if numpr is not None:
            ilvl = numpr.find(_q("ilvl"))
            numId_elem = numpr.find(_q("numId"))
            para.numPr = {
                "ilvl": int(ilvl.get(_q("val"), "0")) if ilvl is not None else 0,
                "numId": int(numId_elem.get(_q("val"), "0")) if numId_elem is not None else 0,
            }

        pBdr = ppr.find(_q("pBdr"))
        if pBdr is not None:
            for side in ("top", "bottom", "left", "right"):
                side_elem = pBdr.find(_q(side))
                if side_elem is not None:
                    border = {
                        "val": side_elem.get(_q("val"), "single"),
                        "sz": side_elem.get(_q("sz"), "4"),
                        "color": side_elem.get(_q("color"), "auto"),
                    }
                    setattr(para, f"paragraph_border_{side}", border)

        shd = ppr.find(_q("shd"))
        if shd is not None:
            para.paragraph_shading = shd.get(_q("fill"))
            para.paragraph_shading_color = shd.get(_q("color"))

        tabs = ppr.find(_q("tabs"))
        if tabs is not None:
            for tab in tabs.findall(_q("tab")):
                para.tab_stops.append({
                    "pos": _twips_to_pt(tab.get(_q("val"))),
                    "val": tab.get(_q("val"), "left"),
                    "leader": tab.get(_q("leader")),
                })

        td = ppr.find(_q("textDirection"))
        if td is not None:
            para.text_direction = td.get(_q("val"))

    for child in p_elem:
        tag = etree.QName(child).localname
        if tag == "r":
            run = _parse_run_element_full(child)
            para.runs.append(run)
        elif tag == "hyperlink":
            rid = child.get(f"{R}id", "")
            url = hyperlinks.get(rid, "")
            for r_elem in child.findall(_q("r")):
                run = _parse_run_element_full(r_elem)
                run.hyperlink_url = url
                para.runs.append(run)

    return para


def _parse_run_element_full(r_elem: etree._Element) -> Run:
    """Parse a <w:r> element into a Run with ALL rPr properties."""
    run = Run()
    run._xml_element = r_elem

    texts = []
    for t in r_elem.findall(f".//{_q('t')}"):
        texts.append(t.text or "")
    run.text = "".join(texts)

    rpr = r_elem.find(_q("rPr"))
    if rpr is None:
        return run

    rfonts = rpr.find(_q("rFonts"))
    if rfonts is not None:
        run.font = rfonts.get(_q("ascii"))
        run.font_east_asia = rfonts.get(_q("eastAsia"))
        run.font_cs = rfonts.get(_q("cs"))
        run.font_hansi = rfonts.get(_q("hAnsi"))

    sz = rpr.find(_q("sz"))
    if sz is not None:
        run.size = int(sz.get(_q("val"), "0")) / 2

    szCs = rpr.find(_q("szCs"))
    if szCs is not None:
        run.size_cs = int(szCs.get(_q("val"), "0")) / 2

    run.bold = _bool_attr(rpr, "b")
    run.bold_cs = _bool_attr(rpr, "bCs")
    run.italic = _bool_attr(rpr, "i")
    run.italic_cs = _bool_attr(rpr, "iCs")
    run.caps = _bool_attr(rpr, "caps")
    run.small_caps = _bool_attr(rpr, "smallCaps")

    u = rpr.find(_q("u"))
    if u is not None:
        run.underline = u.get(_q("val"), "single")

    color = rpr.find(_q("color"))
    if color is not None:
        run.color = color.get(_q("val"), "auto")

    highlight = rpr.find(_q("highlight"))
    if highlight is not None:
        run.highlight = highlight.get(_q("val"), "yellow")

    run.strike = _bool_attr(rpr, "strike")
    run.double_strike = _bool_attr(rpr, "dstrike")
    run.emboss = _bool_attr(rpr, "emboss")
    run.imprint = _bool_attr(rpr, "imprint")
    run.shadow = _bool_attr(rpr, "shadow")
    run.outline = _bool_attr(rpr, "outline")

    vert_align = rpr.find(_q("vertAlign"))
    if vert_align is not None:
        val = vert_align.get(_q("val"), "baseline")
        if val == "superscript":
            run.superscript = True
        elif val == "subscript":
            run.subscript = True

    position = rpr.find(_q("position"))
    if position is not None:
        try:
            run.baseline_offset = int(position.get(_q("val"), "0")) / 2
        except ValueError:
            pass

    spacing = rpr.find(_q("spacing"))
    if spacing is not None:
        try:
            run.char_spacing = _twips_to_pt(spacing.get(_q("val")))
        except (ValueError, TypeError):
            pass

    kern = rpr.find(_q("kern"))
    if kern is not None:
        try:
            run.kerning = int(kern.get(_q("val"), "0")) / 2
        except (ValueError, TypeError):
            pass

    w = rpr.find(_q("w"))
    if w is not None:
        try:
            run.scaling = int(w.get(_q("val"), "100"))
        except ValueError:
            pass

    lang = rpr.find(_q("lang"))
    if lang is not None:
        run.lang = lang.get(_q("val"))
        run.lang_east_asia = lang.get(_q("eastAsia"))
        run.lang_bidi = lang.get(_q("bidi"))

    run.rtl = rpr.find(_q("rtl")) is not None
    run.vanish = _bool_attr(rpr, "vanish")

    em = rpr.find(_q("em"))
    if em is not None:
        run.emphasis_mark = em.get(_q("val"))

    return run


def _parse_table_element_full(tbl_elem: etree._Element, hyperlinks: Dict[str, str]) -> Table:
    """Parse a <w:tbl> element into a Table with extended properties."""
    tbl = Table()
    tbl._xml_element = tbl_elem

    tblPr = tbl_elem.find(_q("tblPr"))
    if tblPr is not None:
        jc = tblPr.find(_q("jc"))
        if jc is not None:
            tbl.alignment = jc.get(_q("val"), "left")
        tblStyle = tblPr.find(_q("tblStyle"))
        if tblStyle is not None:
            tbl.style_id = tblStyle.get(_q("val"), "")
        tblW = tblPr.find(_q("tblW"))
        if tblW is not None:
            w = tblW.get(_q("w"))
            tbl.table_width = int(w) / 20 if w else None
        tblInd = tblPr.find(_q("tblInd"))
        if tblInd is not None:
            w = tblInd.get(_q("w"))
            tbl.table_indent = int(w) / 20 if w else None

    tblGrid = tbl_elem.find(_q("tblGrid"))
    if tblGrid is not None:
        tbl.column_widths = []
        for gc in tblGrid.findall(_q("gridCol")):
            w = gc.get(_q("w"))
            if w:
                tbl.column_widths.append(int(w) / 20)

    for tr_elem in tbl_elem.findall(_q("tr")):
        row_cells = []
        for tc_elem in tr_elem.findall(_q("tc")):
            cell = Cell()
            cell._xml_element = tc_elem

            tcPr = tc_elem.find(_q("tcPr"))
            if tcPr is not None:
                tcW = tcPr.find(_q("tcW"))
                if tcW is not None:
                    w = tcW.get(_q("w"))
                    if w:
                        cell.width = int(w) / 20

                shading = tcPr.find(_q("shd"))
                if shading is not None:
                    cell.shading = shading.get(_q("fill"), "auto")

                vmerge = tcPr.find(_q("vMerge"))
                if vmerge is not None:
                    val = vmerge.get(_q("val"), "continue")
                    cell.merge_down = val if val else "continue"

                gridspan = tcPr.find(_q("gridSpan"))
                if gridspan is not None:
                    val = gridspan.get(_q("val"), "1")
                    if int(val) > 1:
                        cell.merge_across = True

                vAlign = tcPr.find(_q("vAlign"))
                if vAlign is not None:
                    cell.v_align = vAlign.get(_q("val"))

            for p_elem in tc_elem.findall(_q("p")):
                para = _parse_paragraph_element_full(p_elem, hyperlinks)
                para.is_table_cell = True
                cell.paragraphs.append(para)

            row_cells.append(cell)
        tbl.rows.append(row_cells)

    return tbl


def _parse_section_properties_full(sectPr: etree._Element) -> Section:
    """Parse <w:sectPr> into Section (full coverage)."""
    sec = Section()

    pgSz = sectPr.find(_q("pgSz"))
    if pgSz is not None:
        w = pgSz.get(_q("w"))
        h = pgSz.get(_q("h"))
        orient = pgSz.get(_q("orient"), "portrait")
        if w:
            sec.page_width = int(w) / 20
        if h:
            sec.page_height = int(h) / 20
        sec.orientation = orient

    pgMar = sectPr.find(_q("pgMar"))
    if pgMar is not None:
        sec.top_margin = _twips_to_pt(pgMar.get(_q("top"))) or sec.top_margin
        sec.bottom_margin = _twips_to_pt(pgMar.get(_q("bottom"))) or sec.bottom_margin
        sec.left_margin = _twips_to_pt(pgMar.get(_q("left"))) or sec.left_margin
        sec.right_margin = _twips_to_pt(pgMar.get(_q("right"))) or sec.right_margin
        sec.header_distance = _twips_to_pt(pgMar.get(_q("header")))
        sec.footer_distance = _twips_to_pt(pgMar.get(_q("footer")))
        sec.gutter = _twips_to_pt(pgMar.get(_q("gutter"))) or 0.0

    cols = sectPr.find(_q("cols"))
    if cols is not None:
        sec.cols = int(cols.get(_q("num"), "1"))
        space = cols.get(_q("space"))
        if space:
            sec.col_space = int(space) / 20

    pgNumType = sectPr.find(_q("pgNumType"))
    if pgNumType is not None:
        start = pgNumType.get(_q("start"))
        if start is not None:
            sec.page_number_start = int(start)
        fmt = pgNumType.get(_q("fmt"))
        if fmt is not None:
            sec.page_number_fmt = fmt

    titlePg = sectPr.find(_q("titlePg"))
    if titlePg is not None:
        sec.title_page = titlePg.get(_q("val"), "false") != "false"

    return sec


def _attr_val(parent: etree._Element, tag: str, attr: str) -> Optional[str]:
    elem = parent.find(_q(tag))
    if elem is not None:
        return elem.get(_q(attr))
    return None


def _bool_attr(parent: etree._Element, tag: str) -> bool:
    elem = parent.find(_q(tag))
    if elem is not None:
        return elem.get(_q("val"), "true") != "false"
    return False


def _twips_to_pt(val: Optional[str]) -> Optional[float]:
    if val is None:
        return None
    try:
        return int(val) / 20
    except ValueError:
        return None


# ═══════════════════════════════════════════════════════════════════════
# Incremental Serialization (lossless round-trip)
# ═══════════════════════════════════════════════════════════════════════

def serialize_document_model(doc: Document, output_path: str, original_docx: Optional[str] = None) -> str:
    """
    Incrementally serialize a Document model back to .docx.
    Modifies the original XML tree in-place based on model changes,
    preserving all unmapped OOXML attributes for lossless round-trip.
    """
    if original_docx is None and doc._unpacked_dir is None:
        raise SerializeError("No original file or unpacked directory", ErrorCode.SERIALIZE_FAILED)

    if original_docx:
        work_dir = unpack_docx(original_docx)
        # Reparse to get fresh XML tree
        parsed = parse_docx(original_docx)
        doc_root = parsed["document"]
        work_dir = parsed["unpacked_dir"]
    else:
        work_dir = doc._unpacked_dir
        if not work_dir or not Path(work_dir).exists():
            raise SerializeError("Unpacked directory not found", ErrorCode.SERIALIZE_FAILED)
        document_xml = Path(work_dir) / "word" / "document.xml"
        doc_root = parse_xml(str(document_xml))

    body = doc_root.find(_q("body"))
    if body is None:
        raise SerializeError("Invalid document.xml: no body element", ErrorCode.XML_MALFORMED)

    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Collect original body-only elements (exclude sectPr)
    original_paras = []
    original_tables = []
    other_elements = []
    sectPr_elem = None

    for child in list(body):
        tag = etree.QName(child).localname
        if tag == "p":
            original_paras.append(child)
        elif tag == "tbl":
            original_tables.append(child)
        elif tag == "sectPr":
            sectPr_elem = child
        else:
            other_elements.append(child)

    # ── Apply paragraph order: modify in-place, add new, remove deleted ──
    body.clear()

    for i, para in enumerate(doc.paragraphs):
        if para._xml_element is not None and para._xml_element in original_paras:
            p_elem = para._xml_element
            _apply_paragraph_to_element(para, p_elem, ns)
            body.append(p_elem)
        elif i < len(original_paras):
            p_elem = original_paras[i]
            _apply_paragraph_to_element(para, p_elem, ns)
            body.append(p_elem)
        else:
            p_elem = _build_paragraph_element_full(para, ns)
            body.append(p_elem)

    # Append tables
    for tbl in doc.tables:
        if tbl._xml_element is not None:
            tbl_elem = tbl._xml_element
            body.append(tbl_elem)
        else:
            tbl_elem = _build_table_element_full(tbl, ns)
            body.append(tbl_elem)

    # Append preserved non-paragraph/non-table elements
    for elem in other_elements:
        body.append(elem)

    # ── Serialize section properties from model ──
    if doc.sections:
        sec_model = doc.sections[0]
        if sectPr_elem is not None and sec_model._xml_element is not None:
            _apply_section_to_element(sec_model, sectPr_elem, ns)
        elif sectPr_elem is not None:
            _apply_section_to_element(sec_model, sectPr_elem, ns)
        else:
            sectPr_elem = _build_section_element(sec_model, ns)
        body.append(sectPr_elem)
    elif sectPr_elem is not None:
        body.append(sectPr_elem)

    # Write back
    document_xml = Path(work_dir) / "word" / "document.xml"
    tree = doc_root.getroottree()
    tree.write(str(document_xml), xml_declaration=True, encoding="UTF-8", standalone=True)

    result = pack_docx(work_dir, output_path)

    doc.mark_clean()
    return result


# ═══════════════════════════════════════════════════════════════════════
# In-place element modifiers (preserve unmapped attributes)
# ═══════════════════════════════════════════════════════════════════════

def _apply_paragraph_to_element(para: Paragraph, p_elem: etree._Element, ns: str):
    """Modify a <w:p> element in-place based on Paragraph model changes."""
    pPr = p_elem.find(_q("pPr"))
    if pPr is None:
        pPr = etree.Element(f"{{{ns}}}pPr")
        p_elem.insert(0, pPr)

    if para.has_paragraph_formatting():
        _set_or_remove(pPr, f"{{{ns}}}pStyle", para.style_id, "val" if para.style_id else None)
        _set_or_remove(pPr, f"{{{ns}}}jc", para.alignment, "val" if para.alignment else None)
        _set_or_remove(pPr, f"{{{ns}}}outlineLvl", str(para.outline_level) if para.outline_level is not None else None, "val")

        # Indentation
        if any(v is not None for v in (para.first_line_indent, para.left_indent, para.right_indent, para.hanging)):
            ind = _ensure_child(pPr, f"{{{ns}}}ind")
            _set_indent_attr(ind, ns, "firstLine", para.first_line_indent)
            _set_indent_attr(ind, ns, "left", para.left_indent)
            _set_indent_attr(ind, ns, "right", para.right_indent)
            _set_indent_attr(ind, ns, "hanging", para.hanging)

        # Spacing
        if any(v is not None for v in (para.space_before, para.space_after, para.line_spacing)):
            spacing = _ensure_child(pPr, f"{{{ns}}}spacing")
            _set_twips_attr(spacing, ns, "before", para.space_before)
            _set_twips_attr(spacing, ns, "after", para.space_after)
            if para.line_spacing is not None:
                if para.line_rule == "auto":
                    spacing.set(f"{{{ns}}}line", str(int(para.line_spacing * 240)))
                    spacing.set(f"{{{ns}}}lineRule", "auto")
                elif para.line_rule == "exact":
                    spacing.set(f"{{{ns}}}line", str(int(para.line_spacing * 20)))
                    spacing.set(f"{{{ns}}}lineRule", "exact")
                else:
                    spacing.set(f"{{{ns}}}line", str(int(para.line_spacing * 20)))

        # Pagination
        _bool_element(pPr, ns, "keepNext", para.keep_with_next)
        _bool_element(pPr, ns, "keepLines", para.keep_lines)
        _bool_element(pPr, ns, "pageBreakBefore", para.page_break_before)
        if not para.widow_control:
            wc = _ensure_child(pPr, f"{{{ns}}}widowControl")
            wc.set(f"{{{ns}}}val", "false")

        # Numbering
        if para.numPr:
            numPr_elem = _ensure_child(pPr, f"{{{ns}}}numPr")
            ilvl = _ensure_child(numPr_elem, f"{{{ns}}}ilvl")
            ilvl.set(f"{{{ns}}}val", str(para.numPr.get("ilvl", 0)))
            numId = _ensure_child(numPr_elem, f"{{{ns}}}numId")
            numId.set(f"{{{ns}}}val", str(para.numPr.get("numId", 0)))

        # Shading
        if para.paragraph_shading:
            shd = _ensure_child(pPr, f"{{{ns}}}shd")
            shd.set(f"{{{ns}}}fill", para.paragraph_shading)
            if para.paragraph_shading_color:
                shd.set(f"{{{ns}}}color", para.paragraph_shading_color)

        # Tabs
        if para.tab_stops:
            tabs = _ensure_child(pPr, f"{{{ns}}}tabs")
            for ts in para.tab_stops:
                tab = etree.SubElement(tabs, f"{{{ns}}}tab")
                tab.set(f"{{{ns}}}val", ts.get("val", "left"))
                if ts.get("pos"):
                    tab.set(f"{{{ns}}}pos", str(int(ts["pos"] * 20)))
                if ts.get("leader"):
                    tab.set(f"{{{ns}}}leader", ts["leader"])

        # Text direction
        if para.text_direction:
            td = _ensure_child(pPr, f"{{{ns}}}textDirection")
            td.set(f"{{{ns}}}val", para.text_direction)

        # Borders
        for side in ("top", "bottom", "left", "right"):
            border = getattr(para, f"paragraph_border_{side}", None)
            if border:
                pBdr = _ensure_child(pPr, f"{{{ns}}}pBdr")
                b_elem = etree.SubElement(pBdr, f"{{{ns}}}{side}")
                b_elem.set(f"{{{ns}}}val", border.get("val", "single"))
                if border.get("sz"):
                    b_elem.set(f"{{{ns}}}sz", str(border["sz"]))
                if border.get("color"):
                    b_elem.set(f"{{{ns}}}color", str(border["color"]))

    # ── Apply run modifications (always run, even without paragraph formatting) ──
    orig_runs = list(p_elem.findall(_q("r")))
    orig_hyperlinks = list(p_elem.findall(f"{{{ns}}}hyperlink"))

    # Clear existing runs and hyperlinks
    for r in orig_runs:
        p_elem.remove(r)
    for h in orig_hyperlinks:
        p_elem.remove(h)

    model_run_idx = 0
    for orig_run in orig_runs:
        if model_run_idx < len(para.runs):
            run = para.runs[model_run_idx]
            _apply_run_to_element(run, orig_run, ns)
            p_elem.append(orig_run)
            model_run_idx += 1

    # Append any extra model runs that don't have corresponding original runs
    for i in range(model_run_idx, len(para.runs)):
        run = para.runs[i]
        new_r = _build_run_element_full(run, ns)
        p_elem.append(new_r)


def _apply_run_to_element(run: Run, r_elem: etree._Element, ns: str):
    """Modify a <w:r> element in-place based on Run model changes."""
    # Update text — collect all text content, set on first <w:t>, remove rest
    t_elems = r_elem.findall(f".//{_q('t')}")
    if run.text:
        if not t_elems:
            t_elem = etree.SubElement(r_elem, f"{{{ns}}}t")
        else:
            t_elem = t_elems[0]
            # Remove extra <w:t> elements
            for extra in t_elems[1:]:
                extra.getparent().remove(extra)
        t_elem.text = run.text
        if run.text.startswith(" ") or run.text.endswith(" "):
            t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    elif t_elems:
        for te in t_elems:
            te.getparent().remove(te)

    # Update or create rPr
    rPr = r_elem.find(_q("rPr"))
    if not run.has_formatting():
        return
    if rPr is None:
        rPr = etree.Element(f"{{{ns}}}rPr")
        r_elem.insert(0, rPr)

    # Font
    if run.font or run.font_east_asia or run.font_cs or run.font_hansi:
        rfonts = _ensure_child(rPr, f"{{{ns}}}rFonts")
        if run.font:
            rfonts.set(f"{{{ns}}}ascii", run.font)
            rfonts.set(f"{{{ns}}}hAnsi", run.font)
        if run.font_east_asia:
            rfonts.set(f"{{{ns}}}eastAsia", run.font_east_asia)
        if run.font_cs:
            rfonts.set(f"{{{ns}}}cs", run.font_cs)

    # Size
    if run.size:
        sz = _ensure_child(rPr, f"{{{ns}}}sz")
        sz.set(f"{{{ns}}}val", str(int(run.size * 2)))
        szCs = _ensure_child(rPr, f"{{{ns}}}szCs")
        szCs.set(f"{{{ns}}}val", str(int(run.size * 2)))

    # Emphasis
    _bool_element(rPr, ns, "b", run.bold)
    _bool_element(rPr, ns, "bCs", run.bold_cs)
    _bool_element(rPr, ns, "i", run.italic)
    _bool_element(rPr, ns, "iCs", run.italic_cs)
    _bool_element(rPr, ns, "caps", run.caps)
    _bool_element(rPr, ns, "smallCaps", run.small_caps)
    _bool_element(rPr, ns, "strike", run.strike)
    _bool_element(rPr, ns, "dstrike", run.double_strike)
    _bool_element(rPr, ns, "emboss", run.emboss)
    _bool_element(rPr, ns, "imprint", run.imprint)
    _bool_element(rPr, ns, "shadow", run.shadow)
    _bool_element(rPr, ns, "outline", run.outline)
    _bool_element(rPr, ns, "vanish", run.vanish)
    _bool_element(rPr, ns, "rtl", run.rtl)

    # Underline
    if run.underline:
        u = _ensure_child(rPr, f"{{{ns}}}u")
        u.set(f"{{{ns}}}val", run.underline)

    # Color
    if run.color:
        color = _ensure_child(rPr, f"{{{ns}}}color")
        color.set(f"{{{ns}}}val", run.color)

    # Highlight
    if run.highlight:
        hl = _ensure_child(rPr, f"{{{ns}}}highlight")
        hl.set(f"{{{ns}}}val", run.highlight)

    # Vertical align
    if run.superscript:
        va = _ensure_child(rPr, f"{{{ns}}}vertAlign")
        va.set(f"{{{ns}}}val", "superscript")
    elif run.subscript:
        va = _ensure_child(rPr, f"{{{ns}}}vertAlign")
        va.set(f"{{{ns}}}val", "subscript")

    # Position
    if run.baseline_offset is not None:
        pos = _ensure_child(rPr, f"{{{ns}}}position")
        pos.set(f"{{{ns}}}val", str(int(run.baseline_offset * 2)))

    # Character spacing
    if run.char_spacing is not None:
        sp = _ensure_child(rPr, f"{{{ns}}}spacing")
        sp.set(f"{{{ns}}}val", str(int(run.char_spacing * 20)))

    # Kerning
    if run.kerning is not None:
        kern = _ensure_child(rPr, f"{{{ns}}}kern")
        kern.set(f"{{{ns}}}val", str(int(run.kerning * 2)))

    # Scaling
    if run.scaling is not None:
        w = _ensure_child(rPr, f"{{{ns}}}w")
        w.set(f"{{{ns}}}val", str(run.scaling))

    # Language
    if run.lang or run.lang_east_asia or run.lang_bidi:
        lang = _ensure_child(rPr, f"{{{ns}}}lang")
        if run.lang:
            lang.set(f"{{{ns}}}val", run.lang)
        if run.lang_east_asia:
            lang.set(f"{{{ns}}}eastAsia", run.lang_east_asia)
        if run.lang_bidi:
            lang.set(f"{{{ns}}}bidi", run.lang_bidi)

    # Emphasis mark
    if run.emphasis_mark:
        em = _ensure_child(rPr, f"{{{ns}}}em")
        em.set(f"{{{ns}}}val", run.emphasis_mark)


def _apply_section_to_element(sec: Section, sectPr: etree._Element, ns: str):
    """Modify a <w:sectPr> element in-place based on Section model."""
    pgSz = _ensure_child(sectPr, f"{{{ns}}}pgSz")
    pgSz.set(f"{{{ns}}}w", str(int(sec.page_width * 20)))
    pgSz.set(f"{{{ns}}}h", str(int(sec.page_height * 20)))
    pgSz.set(f"{{{ns}}}orient", sec.orientation)

    pgMar = _ensure_child(sectPr, f"{{{ns}}}pgMar")
    pgMar.set(f"{{{ns}}}top", str(int(sec.top_margin * 20)))
    pgMar.set(f"{{{ns}}}bottom", str(int(sec.bottom_margin * 20)))
    pgMar.set(f"{{{ns}}}left", str(int(sec.left_margin * 20)))
    pgMar.set(f"{{{ns}}}right", str(int(sec.right_margin * 20)))
    if sec.header_distance is not None:
        pgMar.set(f"{{{ns}}}header", str(int(sec.header_distance * 20)))
    if sec.footer_distance is not None:
        pgMar.set(f"{{{ns}}}footer", str(int(sec.footer_distance * 20)))
    if sec.gutter:
        pgMar.set(f"{{{ns}}}gutter", str(int(sec.gutter * 20)))

    cols = _ensure_child(sectPr, f"{{{ns}}}cols")
    cols.set(f"{{{ns}}}num", str(sec.cols))
    if sec.col_space is not None:
        cols.set(f"{{{ns}}}space", str(int(sec.col_space * 20)))

    if sec.page_number_start is not None or sec.page_number_fmt is not None:
        pgNumType = _ensure_child(sectPr, f"{{{ns}}}pgNumType")
        if sec.page_number_start is not None:
            pgNumType.set(f"{{{ns}}}start", str(sec.page_number_start))
        if sec.page_number_fmt is not None:
            pgNumType.set(f"{{{ns}}}fmt", sec.page_number_fmt)

    if sec.title_page:
        titlePg = _ensure_child(sectPr, f"{{{ns}}}titlePg")


# ═══════════════════════════════════════════════════════════════════════
# Full element builders (for new paragraphs/runs without _xml_element)
# ═══════════════════════════════════════════════════════════════════════

def _build_paragraph_element_full(para: Paragraph, ns: str) -> etree._Element:
    """Build a new <w:p> element from scratch (full property coverage)."""
    p = etree.Element(f"{{{ns}}}p")
    pPr = etree.SubElement(p, f"{{{ns}}}pPr")

    if para.style_id:
        pStyle = etree.SubElement(pPr, f"{{{ns}}}pStyle")
        pStyle.set(f"{{{ns}}}val", para.style_id)

    if para.alignment:
        jc = etree.SubElement(pPr, f"{{{ns}}}jc")
        jc.set(f"{{{ns}}}val", para.alignment)

    if any(v is not None for v in (para.first_line_indent, para.left_indent, para.right_indent, para.hanging)):
        ind = etree.SubElement(pPr, f"{{{ns}}}ind")
        _set_indent_attr(ind, ns, "firstLine", para.first_line_indent)
        _set_indent_attr(ind, ns, "left", para.left_indent)
        _set_indent_attr(ind, ns, "right", para.right_indent)
        _set_indent_attr(ind, ns, "hanging", para.hanging)

    if any(v is not None for v in (para.space_before, para.space_after, para.line_spacing)):
        spacing = etree.SubElement(pPr, f"{{{ns}}}spacing")
        _set_twips_attr(spacing, ns, "before", para.space_before)
        _set_twips_attr(spacing, ns, "after", para.space_after)
        if para.line_spacing is not None:
            if para.line_rule == "auto":
                spacing.set(f"{{{ns}}}line", str(int(para.line_spacing * 240)))
                spacing.set(f"{{{ns}}}lineRule", "auto")
            elif para.line_rule == "exact":
                spacing.set(f"{{{ns}}}line", str(int(para.line_spacing * 20)))
                spacing.set(f"{{{ns}}}lineRule", "exact")
            else:
                spacing.set(f"{{{ns}}}line", str(int(para.line_spacing * 20)))

    _bool_element(pPr, ns, "keepNext", para.keep_with_next)
    _bool_element(pPr, ns, "keepLines", para.keep_lines)
    _bool_element(pPr, ns, "pageBreakBefore", para.page_break_before)
    if not para.widow_control:
        wc = etree.SubElement(pPr, f"{{{ns}}}widowControl")
        wc.set(f"{{{ns}}}val", "false")

    if para.outline_level is not None:
        ol = etree.SubElement(pPr, f"{{{ns}}}outlineLvl")
        ol.set(f"{{{ns}}}val", str(para.outline_level))

    if para.numPr:
        numPr = etree.SubElement(pPr, f"{{{ns}}}numPr")
        ilvl = etree.SubElement(numPr, f"{{{ns}}}ilvl")
        ilvl.set(f"{{{ns}}}val", str(para.numPr.get("ilvl", 0)))
        numId = etree.SubElement(numPr, f"{{{ns}}}numId")
        numId.set(f"{{{ns}}}val", str(para.numPr.get("numId", 0)))

    for run in para.runs:
        r_elem = _build_run_element_full(run, ns)
        p.append(r_elem)

    return p


def _build_run_element_full(run: Run, ns: str) -> etree._Element:
    """Build a new <w:r> element from scratch (full rPr coverage)."""
    r = etree.Element(f"{{{ns}}}r")
    rPr = etree.SubElement(r, f"{{{ns}}}rPr")

    if run.font or run.font_east_asia or run.font_hansi:
        rFonts = etree.SubElement(rPr, f"{{{ns}}}rFonts")
        if run.font:
            rFonts.set(f"{{{ns}}}ascii", run.font)
            rFonts.set(f"{{{ns}}}hAnsi", run.font)
        if run.font_east_asia:
            rFonts.set(f"{{{ns}}}eastAsia", run.font_east_asia)
        if run.font_cs:
            rFonts.set(f"{{{ns}}}cs", run.font_cs)

    if run.size:
        sz = etree.SubElement(rPr, f"{{{ns}}}sz")
        sz.set(f"{{{ns}}}val", str(int(run.size * 2)))
        szCs = etree.SubElement(rPr, f"{{{ns}}}szCs")
        szCs.set(f"{{{ns}}}val", str(int(run.size * 2)))

    _bool_element(rPr, ns, "b", run.bold)
    _bool_element(rPr, ns, "bCs", run.bold_cs)
    _bool_element(rPr, ns, "i", run.italic)
    _bool_element(rPr, ns, "iCs", run.italic_cs)
    _bool_element(rPr, ns, "caps", run.caps)
    _bool_element(rPr, ns, "smallCaps", run.small_caps)
    _bool_element(rPr, ns, "strike", run.strike)
    _bool_element(rPr, ns, "dstrike", run.double_strike)
    _bool_element(rPr, ns, "emboss", run.emboss)
    _bool_element(rPr, ns, "imprint", run.imprint)
    _bool_element(rPr, ns, "shadow", run.shadow)
    _bool_element(rPr, ns, "outline", run.outline)
    _bool_element(rPr, ns, "vanish", run.vanish)
    _bool_element(rPr, ns, "rtl", run.rtl)

    if run.underline:
        u = etree.SubElement(rPr, f"{{{ns}}}u")
        u.set(f"{{{ns}}}val", run.underline)
    if run.color:
        color = etree.SubElement(rPr, f"{{{ns}}}color")
        color.set(f"{{{ns}}}val", run.color)
    if run.highlight:
        hl = etree.SubElement(rPr, f"{{{ns}}}highlight")
        hl.set(f"{{{ns}}}val", run.highlight)
    if run.superscript:
        va = etree.SubElement(rPr, f"{{{ns}}}vertAlign")
        va.set(f"{{{ns}}}val", "superscript")
    elif run.subscript:
        va = etree.SubElement(rPr, f"{{{ns}}}vertAlign")
        va.set(f"{{{ns}}}val", "subscript")
    if run.baseline_offset is not None:
        pos = etree.SubElement(rPr, f"{{{ns}}}position")
        pos.set(f"{{{ns}}}val", str(int(run.baseline_offset * 2)))
    if run.char_spacing is not None:
        sp = etree.SubElement(rPr, f"{{{ns}}}spacing")
        sp.set(f"{{{ns}}}val", str(int(run.char_spacing * 20)))
    if run.kerning is not None:
        kern = etree.SubElement(rPr, f"{{{ns}}}kern")
        kern.set(f"{{{ns}}}val", str(int(run.kerning * 2)))
    if run.scaling is not None:
        w = etree.SubElement(rPr, f"{{{ns}}}w")
        w.set(f"{{{ns}}}val", str(run.scaling))
    if run.lang or run.lang_east_asia or run.lang_bidi:
        lang = etree.SubElement(rPr, f"{{{ns}}}lang")
        if run.lang:
            lang.set(f"{{{ns}}}val", run.lang)
        if run.lang_east_asia:
            lang.set(f"{{{ns}}}eastAsia", run.lang_east_asia)
        if run.lang_bidi:
            lang.set(f"{{{ns}}}bidi", run.lang_bidi)
    if run.emphasis_mark:
        em = etree.SubElement(rPr, f"{{{ns}}}em")
        em.set(f"{{{ns}}}val", run.emphasis_mark)

    if run.text:
        t = etree.SubElement(r, f"{{{ns}}}t")
        if run.text.startswith(" ") or run.text.endswith(" "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = run.text

    return r


def _build_table_element_full(tbl: Table, ns: str) -> etree._Element:
    """Build a new <w:tbl> element."""
    tbl_elem = etree.Element(f"{{{ns}}}tbl")
    tblPr = etree.SubElement(tbl_elem, f"{{{ns}}}tblPr")
    if tbl.style_id:
        tblStyle = etree.SubElement(tblPr, f"{{{ns}}}tblStyle")
        tblStyle.set(f"{{{ns}}}val", tbl.style_id)
    if tbl.alignment:
        jc = etree.SubElement(tblPr, f"{{{ns}}}jc")
        jc.set(f"{{{ns}}}val", tbl.alignment)

    tblGrid = etree.SubElement(tbl_elem, f"{{{ns}}}tblGrid")
    for cw in tbl.column_widths:
        gridCol = etree.SubElement(tblGrid, f"{{{ns}}}gridCol")
        gridCol.set(f"{{{ns}}}w", str(int(cw * 20)))

    for row_cells in tbl.rows:
        tr = etree.SubElement(tbl_elem, f"{{{ns}}}tr")
        for cell in row_cells:
            tc = etree.SubElement(tr, f"{{{ns}}}tc")
            tcPr = etree.SubElement(tc, f"{{{ns}}}tcPr")
            if cell.width:
                tcW = etree.SubElement(tcPr, f"{{{ns}}}tcW")
                tcW.set(f"{{{ns}}}w", str(int(cell.width * 20)))
                tcW.set(f"{{{ns}}}type", "dxa")
            if cell.shading:
                shd = etree.SubElement(tcPr, f"{{{ns}}}shd")
                shd.set(f"{{{ns}}}val", "clear")
                shd.set(f"{{{ns}}}fill", cell.shading)
            if cell.merge_down:
                vMerge = etree.SubElement(tcPr, f"{{{ns}}}vMerge")
                if cell.merge_down == "continue":
                    vMerge.set(f"{{{ns}}}val", "continue")
                else:
                    vMerge.set(f"{{{ns}}}val", "restart")
            if cell.merge_across:
                gs = etree.SubElement(tcPr, f"{{{ns}}}gridSpan")
                gs.set(f"{{{ns}}}val", "2")
            if cell.v_align:
                va = etree.SubElement(tcPr, f"{{{ns}}}vAlign")
                va.set(f"{{{ns}}}val", cell.v_align)
            for para in cell.paragraphs:
                p_elem = _build_paragraph_element_full(para, ns)
                tc.append(p_elem)

    return tbl_elem


def _build_section_element(sec: Section, ns: str) -> etree._Element:
    """Build a new <w:sectPr> element."""
    sectPr = etree.Element(f"{{{ns}}}sectPr")
    pgSz = etree.SubElement(sectPr, f"{{{ns}}}pgSz")
    pgSz.set(f"{{{ns}}}w", str(int(sec.page_width * 20)))
    pgSz.set(f"{{{ns}}}h", str(int(sec.page_height * 20)))
    pgSz.set(f"{{{ns}}}orient", sec.orientation)

    pgMar = etree.SubElement(sectPr, f"{{{ns}}}pgMar")
    pgMar.set(f"{{{ns}}}top", str(int(sec.top_margin * 20)))
    pgMar.set(f"{{{ns}}}bottom", str(int(sec.bottom_margin * 20)))
    pgMar.set(f"{{{ns}}}left", str(int(sec.left_margin * 20)))
    pgMar.set(f"{{{ns}}}right", str(int(sec.right_margin * 20)))
    if sec.header_distance is not None:
        pgMar.set(f"{{{ns}}}header", str(int(sec.header_distance * 20)))
    if sec.footer_distance is not None:
        pgMar.set(f"{{{ns}}}footer", str(int(sec.footer_distance * 20)))
    if sec.gutter:
        pgMar.set(f"{{{ns}}}gutter", str(int(sec.gutter * 20)))

    cols = etree.SubElement(sectPr, f"{{{ns}}}cols")
    cols.set(f"{{{ns}}}num", str(sec.cols))
    if sec.col_space is not None:
        cols.set(f"{{{ns}}}space", str(int(sec.col_space * 20)))

    if sec.page_number_start is not None or sec.page_number_fmt is not None:
        pgNumType = etree.SubElement(sectPr, f"{{{ns}}}pgNumType")
        if sec.page_number_start is not None:
            pgNumType.set(f"{{{ns}}}start", str(sec.page_number_start))
        if sec.page_number_fmt is not None:
            pgNumType.set(f"{{{ns}}}fmt", sec.page_number_fmt)

    if sec.title_page:
        etree.SubElement(sectPr, f"{{{ns}}}titlePg")

    return sectPr


# ═══════════════════════════════════════════════════════════════════════
# XML helpers
# ═══════════════════════════════════════════════════════════════════════

def _ensure_child(parent: etree._Element, tag: str) -> etree._Element:
    """Get or create a child element."""
    child = parent.find(tag)
    if child is None:
        child = etree.SubElement(parent, tag)
    return child


def _set_or_remove(parent: etree._Element, tag: str, value: Optional[str], attr_name: Optional[str]):
    """Set an attribute on a child element, or remove the child if value is None."""
    if value is None:
        child = parent.find(tag)
        if child is not None:
            parent.remove(child)
    else:
        child = _ensure_child(parent, tag)
        if attr_name:
            child.set(attr_name, value)


def _set_indent_attr(ind: etree._Element, ns: str, name: str, value: Optional[float]):
    if value is not None:
        ind.set(f"{{{ns}}}{name}", str(int(value * 20)))


def _set_twips_attr(elem: etree._Element, ns: str, name: str, value: Optional[float]):
    if value is not None:
        elem.set(f"{{{ns}}}{name}", str(int(value * 20)))


def _bool_element(parent: etree._Element, ns: str, name: str, value: bool):
    """Add or remove a boolean child element like <w:b/>."""
    elem = parent.find(f"{{{ns}}}{name}")
    if value:
        if elem is None:
            etree.SubElement(parent, f"{{{ns}}}{name}")
    else:
        if elem is not None:
            parent.remove(elem)


# Convenience
def load_and_build(docx_path: str) -> Document:
    """Load a .docx file and build the full Document model."""
    parsed = parse_docx(docx_path)
    return build_document_model(parsed)
