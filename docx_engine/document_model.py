# -*- coding: utf-8 -*-
"""
Document Object Model (DOM) — in-memory representation of a .docx document.
Provides run-level precision for reading and writing, with full OOXML property coverage.
"""
from typing import List, Dict, Optional, Any, Iterator
from dataclasses import dataclass, field
from copy import deepcopy


@dataclass
class Run:
    """A single run of text with uniform formatting. Models core OOXML rPr properties."""
    text: str = ""

    # ── Font identity ──
    font: Optional[str] = None              # rFonts/@ascii
    font_east_asia: Optional[str] = None    # rFonts/@eastAsia
    font_cs: Optional[str] = None           # rFonts/@cs
    font_hansi: Optional[str] = None        # rFonts/@hAnsi

    # ── Font emphasis ──
    size: Optional[float] = None            # sz/@val (half-pts → pts)
    size_cs: Optional[float] = None         # szCs/@val
    bold: bool = False                      # b
    bold_cs: bool = False                   # bCs
    italic: bool = False                    # i
    italic_cs: bool = False                 # iCs
    caps: bool = False                      # caps (all caps)
    small_caps: bool = False                # smallCaps

    # ── Decoration ──
    underline: Optional[str] = None         # u/@val (single/double/wave/...)
    color: Optional[str] = None             # color/@val (hex or "auto")
    highlight: Optional[str] = None         # highlight/@val
    strike: bool = False                    # strike
    double_strike: bool = False             # dstrike
    emboss: bool = False                    # emboss
    imprint: bool = False                   # imprint
    shadow: bool = False                    # shadow
    outline: bool = False                   # outline

    # ── Vertical position ──
    superscript: bool = False               # vertAlign=superscript
    subscript: bool = False                 # vertAlign=subscript
    baseline_offset: Optional[float] = None # position/@val (pts)

    # ── Spacing / kerning / scale ──
    char_spacing: Optional[float] = None    # spacing/@val (twips → pts)
    kerning: Optional[float] = None         # kern/@val (half-pts → pts)
    scaling: Optional[int] = None           # w/@val (percent)

    # ── Language / layout ──
    lang: Optional[str] = None              # lang/@val
    lang_east_asia: Optional[str] = None    # lang/@eastAsia
    lang_bidi: Optional[str] = None         # lang/@bidi
    rtl: bool = False                       # rtl
    vanish: bool = False                    # vanish (hidden)

    # ── Emphasis mark ──
    emphasis_mark: Optional[str] = None     # em/@val

    # ── Hyperlink ──
    hyperlink_url: Optional[str] = None
    hyperlink_anchor: Optional[str] = None

    # ── Original XML element (for round-trip fidelity) ──
    _xml_element: Any = field(default=None, repr=False)

    def is_empty(self) -> bool:
        return not self.text

    def clone(self) -> "Run":
        return deepcopy(self)

    def has_formatting(self) -> bool:
        """Return True if this run has any explicit formatting."""
        return any([
            self.font, self.size, self.bold, self.italic, self.underline,
            self.color, self.highlight, self.strike, self.double_strike,
            self.emboss, self.imprint, self.shadow, self.outline,
            self.superscript, self.subscript, self.caps, self.small_caps,
            self.char_spacing, self.kerning, self.scaling,
            self.baseline_offset, self.vanish, self.emphasis_mark,
            self.bold_cs, self.italic_cs, self.size_cs,
        ])


@dataclass
class Paragraph:
    """A paragraph containing one or more runs."""
    runs: List[Run] = field(default_factory=list)

    # ── Identity ──
    style_id: Optional[str] = None          # pStyle/@val

    # ── Alignment ──
    alignment: Optional[str] = None         # jc/@val (left/center/right/justify/both)

    # ── Indentation ──
    first_line_indent: Optional[float] = None  # ind/@firstLine (pts)
    left_indent: Optional[float] = None        # ind/@left (pts)
    right_indent: Optional[float] = None       # ind/@right (pts)
    hanging: Optional[float] = None            # ind/@hanging (pts)

    # ── Spacing ──
    space_before: Optional[float] = None    # spacing/@before (pts)
    space_after: Optional[float] = None     # spacing/@after (pts)
    line_spacing: Optional[float] = None    # spacing/@line
    line_rule: Optional[str] = None         # spacing/@lineRule (auto/exact/atLeast)

    # ── Pagination control ──
    keep_with_next: bool = False            # keepNext
    keep_lines: bool = False                # keepLines
    page_break_before: bool = False         # pageBreakBefore
    widow_control: bool = True              # widowControl

    # ── Outline / numbering ──
    outline_level: Optional[int] = None     # outlineLvl/@val (0-8)
    numPr: Optional[Dict] = None            # numbering properties

    # ── Borders & shading ──
    paragraph_border_top: Optional[Dict] = None     # pBdr/top
    paragraph_border_bottom: Optional[Dict] = None  # pBdr/bottom
    paragraph_border_left: Optional[Dict] = None    # pBdr/left
    paragraph_border_right: Optional[Dict] = None   # pBdr/right
    paragraph_shading: Optional[str] = None         # shd/@fill
    paragraph_shading_color: Optional[str] = None   # shd/@color

    # ── Tabs ──
    tab_stops: List[Dict] = field(default_factory=list)  # [{pos, val, leader}]

    # ── Text direction ──
    text_direction: Optional[str] = None    # textDirection

    # ── Original XML element ──
    _xml_element: Any = field(default=None, repr=False)

    # ── Metadata ──
    is_table_cell: bool = False
    cell_ref: Optional[str] = None

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def is_heading(self) -> bool:
        if self.outline_level is not None and 0 <= self.outline_level <= 8:
            return True
        if self.style_id and self.style_id.lower().startswith("heading"):
            return True
        return False

    def heading_level(self) -> Optional[int]:
        if self.outline_level is not None:
            return self.outline_level + 1
        if self.style_id:
            s = self.style_id.lower()
            if s.startswith("heading"):
                try:
                    return int(s.replace("heading", "").strip())
                except ValueError:
                    pass
        return None

    def is_empty(self) -> bool:
        return not self.text.strip()

    def add_run(self, text: str = "", **kwargs) -> Run:
        run = Run(text=text, **kwargs)
        self.runs.append(run)
        return run

    def clear_runs(self):
        self.runs.clear()

    def clone(self) -> "Paragraph":
        return deepcopy(self)

    def has_paragraph_formatting(self) -> bool:
        """Return True if paragraph-level formatting is explicitly set."""
        return any([
            self.alignment, self.first_line_indent, self.left_indent,
            self.right_indent, self.hanging, self.space_before, self.space_after,
            self.line_spacing, self.line_rule, self.keep_with_next, self.keep_lines,
            self.page_break_before, not self.widow_control,
            self.paragraph_shading, self.tab_stops, self.text_direction,
            self.style_id, self.outline_level is not None, self.numPr,
        ])


@dataclass
class Cell:
    """A table cell containing paragraphs."""
    paragraphs: List[Paragraph] = field(default_factory=list)
    width: Optional[float] = None
    shading: Optional[str] = None
    merge_across: bool = False
    merge_down: bool = False
    v_align: Optional[str] = None           # vAlign (top/center/bottom)
    _xml_element: Any = field(default=None, repr=False)

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)


@dataclass
class Table:
    """A table containing rows of cells."""
    rows: List[List[Cell]] = field(default_factory=list)
    column_widths: List[float] = field(default_factory=list)
    alignment: Optional[str] = None
    style_id: Optional[str] = None
    table_width: Optional[float] = None
    table_indent: Optional[float] = None
    _xml_element: Any = field(default=None, repr=False)

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def col_count(self) -> int:
        return len(self.rows[0]) if self.rows else 0

    def cell(self, row: int, col: int) -> Optional[Cell]:
        if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
            return self.rows[row][col]
        return None


@dataclass
class Section:
    """Document section properties — full OOXML coverage."""
    page_width: float = 595.3
    page_height: float = 841.9
    orientation: str = "portrait"
    top_margin: float = 72.0
    bottom_margin: float = 72.0
    left_margin: float = 72.0
    right_margin: float = 72.0
    header_distance: Optional[float] = None
    footer_distance: Optional[float] = None
    gutter: float = 0.0
    cols: int = 1
    col_space: Optional[float] = None
    page_number_start: Optional[int] = None     # pgNumType/@start
    page_number_fmt: Optional[str] = None       # pgNumType/@fmt
    title_page: bool = False                     # titlePg/@val
    _xml_element: Any = field(default=None, repr=False)

    def has_changes(self) -> bool:
        """Return True if any property differs from A4 defaults."""
        return (
            abs(self.page_width - 595.3) > 0.01 or
            abs(self.page_height - 841.9) > 0.01 or
            self.orientation != "portrait" or
            abs(self.top_margin - 72.0) > 0.01 or
            abs(self.bottom_margin - 72.0) > 0.01 or
            abs(self.left_margin - 72.0) > 0.01 or
            abs(self.right_margin - 72.0) > 0.01 or
            abs(self.gutter) > 0.01 or
            self.cols != 1 or
            self.header_distance is not None or
            self.footer_distance is not None or
            self.page_number_start is not None or
            self.page_number_fmt is not None or
            self.title_page
        )


@dataclass
class Document:
    """Complete document model."""
    paragraphs: List[Paragraph] = field(default_factory=list)
    tables: List[Table] = field(default_factory=list)
    sections: List[Section] = field(default_factory=list)
    styles: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    numbering: Dict[str, Any] = field(default_factory=dict)
    hyperlinks: Dict[str, str] = field(default_factory=dict)
    bookmarks: Dict[str, Any] = field(default_factory=dict)
    comments: List[Dict] = field(default_factory=list)
    footnotes: List[Dict] = field(default_factory=list)
    endnotes: List[Dict] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)
    media_files: List[str] = field(default_factory=list)
    headers: Dict[str, Any] = field(default_factory=dict)
    footers: Dict[str, Any] = field(default_factory=dict)
    settings: Dict[str, Any] = field(default_factory=dict)
    font_table: Dict[str, Any] = field(default_factory=dict)
    theme: Optional[Any] = None

    # Dirty tracking
    _dirty_paragraphs: set = field(default_factory=set, repr=False)
    _dirty_tables: set = field(default_factory=set, repr=False)
    _deleted_paragraph_indices: set = field(default_factory=set, repr=False)
    _unpacked_dir: Optional[str] = None

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)

    def get_paragraph(self, index: int) -> Optional[Paragraph]:
        if 1 <= index <= len(self.paragraphs):
            return self.paragraphs[index - 1]
        return None

    def iter_paragraphs(self) -> Iterator[Paragraph]:
        yield from self.paragraphs

    def iter_all_paragraphs(self) -> Iterator[Paragraph]:
        for p in self.paragraphs:
            yield p
        for tbl in self.tables:
            for row in tbl.rows:
                for cell in row:
                    for p in cell.paragraphs:
                        yield p

    def insert_paragraph(self, index: int, paragraph: Paragraph) -> Paragraph:
        self.paragraphs.insert(index - 1, paragraph)
        self._dirty_paragraphs.add(index)
        return paragraph

    def append_paragraph(self, paragraph: Paragraph) -> Paragraph:
        self.paragraphs.append(paragraph)
        self._dirty_paragraphs.add(len(self.paragraphs))
        return paragraph

    def delete_paragraph(self, index: int) -> Optional[Paragraph]:
        if 1 <= index <= len(self.paragraphs):
            self._dirty_paragraphs.discard(index)
            self._deleted_paragraph_indices.add(index)
            return self.paragraphs.pop(index - 1)
        return None

    def get_heading_structure(self) -> List[Dict]:
        result = []
        for i, p in enumerate(self.paragraphs, 1):
            level = p.heading_level()
            if level is not None:
                result.append({
                    "index": i,
                    "level": level,
                    "text": p.text,
                    "style_id": p.style_id,
                })
        return result

    def find_paragraphs_by_text(self, text: str, case_sensitive: bool = True) -> List[int]:
        indices = []
        for i, p in enumerate(self.paragraphs, 1):
            ptext = p.text
            ttext = text
            if not case_sensitive:
                ptext = ptext.lower()
                ttext = ttext.lower()
            if ttext in ptext:
                indices.append(i)
        return indices

    def find_paragraphs_by_style(self, style_id: str) -> List[int]:
        return [i for i, p in enumerate(self.paragraphs, 1) if p.style_id == style_id]

    def replace_text(self, old: str, new: str, case_sensitive: bool = True) -> int:
        count = 0
        for para in self.iter_all_paragraphs():
            for run in para.runs:
                if case_sensitive:
                    if old in run.text:
                        count += 1
                        run.text = run.text.replace(old, new)
                else:
                    rt_lower = run.text.lower()
                    old_lower = old.lower()
                    if old_lower in rt_lower:
                        count += 1
                        idx = rt_lower.index(old_lower)
                        actual_old = run.text[idx:idx + len(old)]
                        run.text = run.text.replace(actual_old, new)
        return count

    def get_statistics(self) -> Dict[str, Any]:
        total_chars = sum(len(p.text) for p in self.iter_all_paragraphs())
        total_words = sum(len(p.text.split()) for p in self.iter_all_paragraphs())
        heading_count = sum(1 for p in self.paragraphs if p.is_heading())
        return {
            "paragraph_count": len(self.paragraphs),
            "table_count": len(self.tables),
            "section_count": len(self.sections),
            "total_characters": total_chars,
            "total_words": total_words,
            "heading_count": heading_count,
            "media_files": len(self.media_files),
        }

    def is_dirty(self) -> bool:
        return bool(self._dirty_paragraphs or self._dirty_tables or self._deleted_paragraph_indices)

    def mark_clean(self):
        self._dirty_paragraphs.clear()
        self._dirty_tables.clear()
        self._deleted_paragraph_indices.clear()

    def detect_semantic_structure(self) -> List[Dict]:
        import re

        result = []
        for i, p in enumerate(self.paragraphs, 1):
            text = p.text.strip()
            text_preview = text[:80]
            outline_level = p.outline_level
            heading_level = p.heading_level()
            style_id = p.style_id

            semantic_role = "body"

            if heading_level is not None:
                semantic_role = f"heading_{heading_level}"
            elif style_id and "toc" in (style_id or "").lower():
                semantic_role = "toc"
            elif re.match(r'^[（(]?\s*(摘要|abstract)\s*[）)]?\s*$', text, re.IGNORECASE):
                semantic_role = "abstract_label"
            elif re.match(r'^[（(]?\s*(关键词|关键字|keywords|key\s*words)\s*[）)]?\s*[:：]?', text, re.IGNORECASE):
                semantic_role = "keywords_label"
            elif re.match(r'^(图|Figure|Fig\.?)\s*\d+', text, re.IGNORECASE):
                semantic_role = "figure_caption"
            elif re.match(r'^(表|Table)\s*\d+', text, re.IGNORECASE):
                semantic_role = "table_caption"
            elif re.match(r'^[（(]?\s*(方程|公式|Equation|Formula)\s*\(?\d*\)?\s*[）)]?\s*[:：]?', text, re.IGNORECASE):
                semantic_role = "equation"
            elif re.match(r'^\s*(参考文献|References|Bibliography)\s*$', text, re.IGNORECASE):
                semantic_role = "reference_section_header"
            elif re.match(r'^\[\d+\]', text):
                semantic_role = "reference_item"
            elif re.match(r'^\s*(目录|Table\s*of\s*Contents|Contents)\s*$', text, re.IGNORECASE):
                semantic_role = "toc_heading"
            elif re.match(r'^\s*(附录|Appendix)\s*[A-Za-z]*\s*$', text, re.IGNORECASE):
                semantic_role = "appendix"
            elif re.match(r'^\s*(致谢|Acknowledgements?)\s*$', text, re.IGNORECASE):
                semantic_role = "acknowledgements"
            elif re.match(r'^\s*(作者简介|作者信息|Author\s*Information)\s*$', text, re.IGNORECASE):
                semantic_role = "author_info"
            elif re.match(r'^\s*\d+[\.\)、]\s', text):
                semantic_role = "list_item"
            elif re.match(r'^[一二三四五六七八九十]+[、．.]', text):
                semantic_role = "list_item"
            elif re.match(r'^[（(]\d+[）)]', text):
                semantic_role = "list_item"
            elif text == "":
                semantic_role = "empty"

            has_table_before = False
            has_image_before = False

            result.append({
                "index": i,
                "text_preview": text_preview,
                "outline_level": outline_level,
                "heading_level": heading_level,
                "semantic_role": semantic_role,
                "style_id": style_id,
                "has_table_before": has_table_before,
                "has_image_before": has_image_before,
            })
        return result

    def detect_cross_references(self) -> List[Dict]:
        import re

        references = []

        patterns = [
            ("figure", re.compile(r'图\s*\d+')),
            ("table", re.compile(r'表\s*\d+')),
            ("chapter", re.compile(r'第[一二三四五六七八九十\d]+章')),
            ("see_also", re.compile(r'(?:参见|见)\s*[^\s，。,\.]+')),
            ("equation", re.compile(r'(?:方程|公式)\s*\(?\d+\)?')),
            ("citation", re.compile(r'\[[\d,]+\]')),
        ]

        for i, p in enumerate(self.paragraphs, 1):
            text = p.text
            for ref_type, pattern in patterns:
                for m in pattern.finditer(text):
                    references.append({
                        "para_index": i,
                        "ref_type": ref_type,
                        "ref_text": m.group(),
                        "position": m.start(),
                    })

        return references

    def get_full_structure(self) -> Dict:
        paragraphs_data = []
        for i, p in enumerate(self.paragraphs, 1):
            runs_data = []
            for run in p.runs:
                run_props = {}
                if run.text:
                    run_props["text"] = run.text
                if run.font:
                    run_props["font"] = run.font
                if run.font_east_asia:
                    run_props["font_east_asia"] = run.font_east_asia
                if run.font_cs:
                    run_props["font_cs"] = run.font_cs
                if run.font_hansi:
                    run_props["font_hansi"] = run.font_hansi
                if run.size is not None:
                    run_props["size"] = run.size
                if run.size_cs is not None:
                    run_props["size_cs"] = run.size_cs
                if run.bold:
                    run_props["bold"] = True
                if run.bold_cs:
                    run_props["bold_cs"] = True
                if run.italic:
                    run_props["italic"] = True
                if run.italic_cs:
                    run_props["italic_cs"] = True
                if run.caps:
                    run_props["caps"] = True
                if run.small_caps:
                    run_props["small_caps"] = True
                if run.underline:
                    run_props["underline"] = run.underline
                if run.color:
                    run_props["color"] = run.color
                if run.highlight:
                    run_props["highlight"] = run.highlight
                if run.strike:
                    run_props["strike"] = True
                if run.double_strike:
                    run_props["double_strike"] = True
                if run.emboss:
                    run_props["emboss"] = True
                if run.imprint:
                    run_props["imprint"] = True
                if run.shadow:
                    run_props["shadow"] = True
                if run.outline:
                    run_props["outline"] = True
                if run.superscript:
                    run_props["superscript"] = True
                if run.subscript:
                    run_props["subscript"] = True
                if run.baseline_offset is not None:
                    run_props["baseline_offset"] = run.baseline_offset
                if run.char_spacing is not None:
                    run_props["char_spacing"] = run.char_spacing
                if run.kerning is not None:
                    run_props["kerning"] = run.kerning
                if run.scaling is not None:
                    run_props["scaling"] = run.scaling
                if run.lang:
                    run_props["lang"] = run.lang
                if run.lang_east_asia:
                    run_props["lang_east_asia"] = run.lang_east_asia
                if run.emphasis_mark:
                    run_props["emphasis_mark"] = run.emphasis_mark
                runs_data.append(run_props)

            paragraphs_data.append({
                "index": i,
                "text": p.text,
                "style_id": p.style_id,
                "alignment": p.alignment,
                "outline_level": p.outline_level,
                "heading_level": p.heading_level(),
                "runs": runs_data,
            })

        tables_data = []
        for idx, tbl in enumerate(self.tables, 1):
            cells_data = []
            for row in tbl.rows:
                row_data = []
                for cell in row:
                    row_data.append({
                        "text": cell.text,
                        "paragraphs": [cp.text for cp in cell.paragraphs],
                    })
                cells_data.append(row_data)
            tables_data.append({
                "index": idx,
                "rows": tbl.row_count,
                "cols": tbl.col_count,
                "cells": cells_data,
            })

        sections_data = []
        for sec in self.sections:
            sections_data.append({
                "page_width": sec.page_width,
                "page_height": sec.page_height,
                "orientation": sec.orientation,
                "margins": {
                    "top": sec.top_margin,
                    "bottom": sec.bottom_margin,
                    "left": sec.left_margin,
                    "right": sec.right_margin,
                },
            })

        headings = self.get_heading_structure()
        statistics = self.get_statistics()

        return {
            "paragraphs": paragraphs_data,
            "tables": tables_data,
            "sections": sections_data,
            "headings": headings,
            "statistics": statistics,
        }
