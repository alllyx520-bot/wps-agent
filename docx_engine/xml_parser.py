# -*- coding: utf-8 -*-
"""
XML Native Parser — unpack .docx ZIP and parse XML into lxml trees.
Reads word/document.xml at Run-level precision, preserving all formatting.
"""
import zipfile
import os
import shutil
from pathlib import Path
from typing import Dict, Optional, Any
from lxml import etree

from .errors import ParseError, ErrorCode

# OOXML namespaces
NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _q(tag: str) -> str:
    """Qualified name with w: namespace."""
    return f"{W}{tag}"


def unpack_docx(docx_path: str, output_dir: Optional[str] = None) -> str:
    """Extract .docx ZIP to a directory. Returns the output directory path."""
    src = Path(docx_path)
    if not src.exists():
        raise ParseError(f"File not found: {docx_path}", ErrorCode.MISSING_PART)
    if output_dir is None:
        output_dir = str(src.with_suffix("")) + "_unpacked"
    out = Path(output_dir)
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(src, "r") as zf:
            zf.extractall(out)
    except zipfile.BadZipFile as e:
        raise ParseError(f"Invalid ZIP file: {e}", ErrorCode.ZIP_INVALID, {"path": docx_path}) from e
    return str(out)


def pack_docx(source_dir: str, output_path: str) -> str:
    """Repack a directory into a .docx ZIP file."""
    src = Path(source_dir)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_path in src.rglob("*"):
            if file_path.is_file():
                arcname = file_path.relative_to(src).as_posix()
                zf.write(file_path, arcname)
    return str(out)


def parse_xml(xml_path: str) -> etree._Element:
    """Parse an XML file into an lxml ElementTree."""
    try:
        tree = etree.parse(xml_path)
        return tree.getroot()
    except etree.XMLSyntaxError as e:
        raise ParseError(f"Malformed XML: {e}", ErrorCode.XML_MALFORMED, {"path": xml_path}) from e


def parse_docx(docx_path: str) -> Dict[str, Any]:
    """
    Parse a .docx file and return a dictionary of all relevant XML trees and metadata.
    Returns: {
        "document": <root of word/document.xml>,
        "styles": <root of word/styles.xml> or None,
        "numbering": <root of word/numbering.xml> or None,
        "rels": <root of word/_rels/document.xml.rels>,
        "content_types": <root of [Content_Types].xml>,
        "media_files": [list of media file paths],
        "unpacked_dir": path to unpacked directory,
    }
    """
    unpacked = unpack_docx(docx_path)
    word_dir = Path(unpacked) / "word"

    document_xml = word_dir / "document.xml"
    if not document_xml.exists():
        raise ParseError("word/document.xml not found", ErrorCode.MISSING_PART, {"path": docx_path})

    result = {
        "document": parse_xml(str(document_xml)),
        "styles": None,
        "numbering": None,
        "rels": None,
        "content_types": None,
        "footnotes": None,
        "endnotes": None,
        "settings": None,
        "fontTable": None,
        "theme": None,
        "headers": {},
        "footers": {},
        "media_files": [],
        "unpacked_dir": unpacked,
    }

    styles_xml = word_dir / "styles.xml"
    if styles_xml.exists():
        result["styles"] = parse_xml(str(styles_xml))

    numbering_xml = word_dir / "numbering.xml"
    if numbering_xml.exists():
        result["numbering"] = parse_xml(str(numbering_xml))

    rels_xml = word_dir / "_rels" / "document.xml.rels"
    if rels_xml.exists():
        result["rels"] = parse_xml(str(rels_xml))

    ct_xml = Path(unpacked) / "[Content_Types].xml"
    if ct_xml.exists():
        result["content_types"] = parse_xml(str(ct_xml))

    # Footnotes / endnotes
    footnotes_xml = word_dir / "footnotes.xml"
    if footnotes_xml.exists():
        result["footnotes"] = parse_xml(str(footnotes_xml))

    endnotes_xml = word_dir / "endnotes.xml"
    if endnotes_xml.exists():
        result["endnotes"] = parse_xml(str(endnotes_xml))

    # Settings
    settings_xml = word_dir / "settings.xml"
    if settings_xml.exists():
        result["settings"] = parse_xml(str(settings_xml))

    # Font table
    fontTable_xml = word_dir / "fontTable.xml"
    if fontTable_xml.exists():
        result["fontTable"] = parse_xml(str(fontTable_xml))

    # Theme
    theme_dir = word_dir / "theme"
    if theme_dir.exists():
        theme_xml = theme_dir / "theme1.xml"
        if theme_xml.exists():
            result["theme"] = parse_xml(str(theme_xml))

    # Headers
    for hdr_file in word_dir.glob("header*.xml"):
        try:
            result["headers"][hdr_file.name] = parse_xml(str(hdr_file))
        except Exception:
            pass

    # Footers
    for ftr_file in word_dir.glob("footer*.xml"):
        try:
            result["footers"][ftr_file.name] = parse_xml(str(ftr_file))
        except Exception:
            pass

    media_dir = word_dir / "media"
    if media_dir.exists():
        result["media_files"] = [str(f) for f in media_dir.iterdir() if f.is_file()]

    return result


def get_paragraphs(root: etree._Element) -> list:
    """Return all <w:p> elements from the document body."""
    body = root.find(_q("body"))
    if body is None:
        return []
    return body.findall(_q("p"))


def get_runs(paragraph: etree._Element) -> list:
    """Return all <w:r> elements within a paragraph."""
    return paragraph.findall(_q("r"))


def get_run_text(run: etree._Element) -> str:
    """Extract text content from a <w:r> element."""
    texts = []
    for t in run.findall(f".//{_q('t')}"):
        texts.append(t.text or "")
    return "".join(texts)


def get_paragraph_text(paragraph: etree._Element) -> str:
    """Extract full text from a paragraph."""
    return "".join(get_run_text(r) for r in get_runs(paragraph))


def get_run_properties(run: etree._Element) -> Dict[str, Any]:
    """Extract formatting properties from a <w:rPr> element."""
    rpr = run.find(_q("rPr"))
    props = {}
    if rpr is None:
        return props

    # Font name
    rfonts = rpr.find(_q("rFonts"))
    if rfonts is not None:
        props["font"] = rfonts.get(_q("ascii")) or rfonts.get(_q("eastAsia")) or rfonts.get(_q("hAnsi"))

    # Size
    sz = rpr.find(_q("sz"))
    if sz is not None:
        props["size"] = int(sz.get(_q("val"), "0")) / 2  # half-points to points

    # Bold
    b = rpr.find(_q("b"))
    props["bold"] = b is not None and b.get(_q("val"), "true") != "false"

    # Italic
    i = rpr.find(_q("i"))
    props["italic"] = i is not None and i.get(_q("val"), "true") != "false"

    # Underline
    u = rpr.find(_q("u"))
    if u is not None:
        props["underline"] = u.get(_q("val"), "single")

    # Color
    color = rpr.find(_q("color"))
    if color is not None:
        props["color"] = color.get(_q("val"), "auto")

    # Highlight
    highlight = rpr.find(_q("highlight"))
    if highlight is not None:
        props["highlight"] = highlight.get(_q("val"), "yellow")

    # Superscript / Subscript
    vert_align = rpr.find(_q("vertAlign"))
    if vert_align is not None:
        props["vertical_align"] = vert_align.get(_q("val"), "baseline")

    # StrikeThrough
    strike = rpr.find(_q("strike"))
    props["strike"] = strike is not None and strike.get(_q("val"), "true") != "false"

    return props


def get_paragraph_properties(paragraph: etree._Element) -> Dict[str, Any]:
    """Extract formatting properties from a <w:pPr> element."""
    ppr = paragraph.find(_q("pPr"))
    props = {}
    if ppr is None:
        return props

    # Style reference
    pstyle = ppr.find(_q("pStyle"))
    if pstyle is not None:
        props["style_id"] = pstyle.get(_q("val"), "")

    # Alignment
    jc = ppr.find(_q("jc"))
    if jc is not None:
        props["alignment"] = jc.get(_q("val"), "left")

    # Indentation
    ind = ppr.find(_q("ind"))
    if ind is not None:
        props["first_line_indent"] = _twips_to_pt(ind.get(_q("firstLine")))
        props["left_indent"] = _twips_to_pt(ind.get(_q("left")))
        props["right_indent"] = _twips_to_pt(ind.get(_q("right")))
        props["hanging"] = _twips_to_pt(ind.get(_q("hanging")))

    # Spacing
    spacing = ppr.find(_q("spacing"))
    if spacing is not None:
        props["space_before"] = _twips_to_pt(spacing.get(_q("before")))
        props["space_after"] = _twips_to_pt(spacing.get(_q("after")))
        line = spacing.get(_q("line"))
        line_rule = spacing.get(_q("lineRule"))
        if line:
            if line_rule == "auto":
                props["line_spacing"] = int(line) / 240  # 240 = single line
            else:
                props["line_spacing"] = _twips_to_pt(line)
        props["line_rule"] = line_rule

    # Outline level
    outline = ppr.find(_q("outlineLvl"))
    if outline is not None:
        props["outline_level"] = int(outline.get(_q("val"), "9"))

    # Numbering
    numpr = ppr.find(_q("numPr"))
    if numpr is not None:
        ilvl = numpr.find(_q("ilvl"))
        numId = numpr.find(_q("numId"))
        props["numPr"] = {
            "ilvl": int(ilvl.get(_q("val"), "0")) if ilvl is not None else 0,
            "numId": int(numId.get(_q("val"), "0")) if numId is not None else 0,
        }

    return props


def _twips_to_pt(val: Optional[str]) -> Optional[float]:
    """Convert twips (1/20 of a point) to points."""
    if val is None:
        return None
    try:
        return int(val) / 20
    except ValueError:
        return None


def get_tables(root: etree._Element) -> list:
    """Return all <w:tbl> elements from the document body."""
    body = root.find(_q("body"))
    if body is None:
        return []
    return body.findall(_q("tbl"))


def get_table_rows(table: etree._Element) -> list:
    """Return all <w:tr> elements within a table."""
    return table.findall(_q("tr"))


def get_table_cells(row: etree._Element) -> list:
    """Return all <w:tc> elements within a table row."""
    return row.findall(_q("tc"))


def get_cell_text(cell: etree._Element) -> str:
    """Extract text from a table cell."""
    paragraphs = cell.findall(_q("p"))
    return "\n".join(get_paragraph_text(p) for p in paragraphs)


def get_hyperlinks(root: etree._Element) -> list:
    """Return all hyperlinks in the document body."""
    body = root.find(_q("body"))
    if body is None:
        return []
    links = []
    for hyperlink in body.findall(f".//{_q('hyperlink')}"):
        rid = hyperlink.get(f"{R}id", "")
        text = "".join(t.text or "" for t in hyperlink.findall(f".//{_q('t')}"))
        links.append({"r_id": rid, "text": text})
    return links


def get_bookmarks(root: etree._Element) -> list:
    """Return all bookmarks in the document body."""
    body = root.find(_q("body"))
    if body is None:
        return []
    bms = []
    for bm_start in body.findall(f".//{_q('bookmarkStart')}"):
        bms.append({
            "id": bm_start.get(_q("id"), ""),
            "name": bm_start.get(_q("name"), ""),
        })
    return bms


def get_comments(root: etree._Element) -> list:
    """Return comment range markers from the document body."""
    body = root.find(_q("body"))
    if body is None:
        return []
    comments = []
    for cstart in body.findall(f".//{_q('commentRangeStart')}"):
        comments.append({
            "id": cstart.get(_q("id"), ""),
            "type": "start",
        })
    return comments
