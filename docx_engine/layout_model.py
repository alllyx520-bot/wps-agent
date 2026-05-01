# -*- coding: utf-8 -*-
"""
Layout Model — page geometry analysis, text overflow detection,
header/footer linkage analysis, image-text wrapping inspection.
"""
from typing import List, Dict, Optional, Any, Tuple
from dataclasses import dataclass, field
from .document_model import Document, Section


@dataclass
class PageGeometry:
    width: float  # points
    height: float  # points
    top_margin: float
    bottom_margin: float
    left_margin: float
    right_margin: float
    printable_width: float
    printable_height: float
    columns: int
    column_gap: float = 0.0

    @classmethod
    def from_section(cls, section: Section) -> "PageGeometry":
        pw = section.page_width
        ph = section.page_height
        tm = section.top_margin
        bm = section.bottom_margin
        lm = section.left_margin
        rm = section.right_margin
        if section.orientation == "landscape":
            pw, ph = ph, pw
        return cls(
            width=pw,
            height=ph,
            top_margin=tm,
            bottom_margin=bm,
            left_margin=lm,
            right_margin=rm,
            printable_width=pw - lm - rm,
            printable_height=ph - tm - bm,
            columns=section.cols,
            column_gap=section.col_space or 0.0,
        )


@dataclass
class HeaderFooterAnalysis:
    has_first_page_different: bool
    has_odd_even_different: bool
    sections: List[Dict]
    link_chain: List[List[int]]  # groups of sections linked together


@dataclass
class LayoutIssue:
    issue_type: str  # widow/orphan/page_break/image_position/column_balance
    severity: str  # warning/error
    location: str  # paragraph index or description
    message: str
    suggested_fix: str


@dataclass
class LayoutReport:
    geometry: PageGeometry
    section_count: int
    estimated_pages: int
    header_footer: HeaderFooterAnalysis
    text_overflow_risks: List[Dict]
    image_placement_issues: List[Dict]
    issues: List[LayoutIssue]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "geometry": {
                "width": self.geometry.width,
                "height": self.geometry.height,
                "top_margin": self.geometry.top_margin,
                "bottom_margin": self.geometry.bottom_margin,
                "left_margin": self.geometry.left_margin,
                "right_margin": self.geometry.right_margin,
                "printable_width": self.geometry.printable_width,
                "printable_height": self.geometry.printable_height,
                "columns": self.geometry.columns,
                "column_gap": self.geometry.column_gap,
            },
            "section_count": self.section_count,
            "estimated_pages": self.estimated_pages,
            "header_footer": {
                "first_page_different": self.header_footer.has_first_page_different,
                "odd_even_different": self.header_footer.has_odd_even_different,
                "sections": self.header_footer.sections,
                "link_chain": self.header_footer.link_chain,
            },
            "text_overflow_risks": self.text_overflow_risks,
            "image_placement_issues": self.image_placement_issues,
            "issues": [{"type": i.issue_type, "severity": i.severity,
                         "location": i.location, "message": i.message,
                         "suggested_fix": i.suggested_fix} for i in self.issues],
        }


class LayoutAnalyzer:
    """Analyze page geometry, text flow, headers/footers, and image placement."""

    # Approximate characters per line for Chinese text at 12pt
    CHARS_PER_LINE_12PT_A4 = 38
    LINES_PER_PAGE_A4 = 42

    def __init__(self, doc: Document):
        self.doc = doc
        self.geometry: Optional[PageGeometry] = None

    def analyze(self) -> LayoutReport:
        self._analyze_geometry()
        hf = self._analyze_headers_footers()
        overflow = self._detect_text_overflow()
        images = self._analyze_image_placement()
        table_issues = self._detect_table_overflow()
        column_issues = self._detect_column_imbalance()
        issues = self._collect_layout_issues(overflow, images)
        for t in table_issues:
            issues.append(LayoutIssue(
                issue_type="table_overflow", severity="warning",
                location=f"table {t['table_index']}",
                message=t.get("description", ""),
                suggested_fix="Enable 'Repeat Header Rows' under Table Properties → Row",
            ))
        for c in column_issues:
            issues.append(LayoutIssue(
                issue_type="column_balance", severity="info",
                location=f"section {c['section_index']}",
                message=c.get("description", ""),
                suggested_fix="Check column content balance",
            ))
        pages = self._estimate_page_count()

        return LayoutReport(
            geometry=self.geometry or PageGeometry(595.3, 841.9, 72, 72, 72, 72, 451.3, 697.9, 1),
            section_count=len(self.doc.sections),
            estimated_pages=pages,
            header_footer=hf,
            text_overflow_risks=overflow,
            image_placement_issues=images,
            issues=issues,
        )

    def _analyze_geometry(self):
        if self.doc.sections:
            sec = self.doc.sections[0]
            self.geometry = PageGeometry.from_section(sec)
        else:
            self.geometry = PageGeometry(595.3, 841.9, 72, 72, 72, 72, 451.3, 697.9, 1)

    def _analyze_headers_footers(self) -> HeaderFooterAnalysis:
        sections_info = []
        for i, sec in enumerate(self.doc.sections, 1):
            sections_info.append({
                "index": i,
                "title_page": sec.title_page,
                "page_number_start": sec.page_number_start,
                "page_number_fmt": sec.page_number_fmt,
                "header_distance": sec.header_distance,
                "footer_distance": sec.footer_distance,
            })

        has_first_page = any(s.title_page for s in self.doc.sections)

        # Default to single-segment link chain
        link_chain = [list(range(1, len(self.doc.sections) + 1))]

        return HeaderFooterAnalysis(
            has_first_page_different=has_first_page,
            has_odd_even_different=False,
            sections=sections_info,
            link_chain=link_chain,
        )

    def _estimate_page_count(self) -> int:
        if not self.geometry:
            return 0
        total_chars = sum(len(p.text) for p in self.doc.paragraphs)
        col_factor = self.geometry.columns
        chars_per_line = self.CHARS_PER_LINE_12PT_A4 * col_factor
        if chars_per_line == 0:
            return 0
        total_lines = total_chars / chars_per_line
        pages = total_lines / self.LINES_PER_PAGE_A4
        return max(1, int(pages + 0.999))

    def _detect_text_overflow(self) -> List[Dict]:
        risks = []
        if not self.geometry:
            return risks

        col_factor = self.geometry.columns
        chars_per_line = self.CHARS_PER_LINE_12PT_A4 * col_factor

        for i, para in enumerate(self.doc.paragraphs, 1):
            text = para.text
            if not text:
                continue

            estimated_lines = len(text) / chars_per_line if chars_per_line > 0 else 1
            # Long paragraphs risk overflow into next page
            if estimated_lines > self.LINES_PER_PAGE_A4 * 0.8:
                risks.append({
                    "index": i,
                    "estimated_lines": round(estimated_lines, 1),
                    "risk": "high" if estimated_lines > self.LINES_PER_PAGE_A4 else "medium",
                    "description": f"Paragraph may span {int(estimated_lines)} lines",
                })

            # Very long paragraphs risk page break issues
            if estimated_lines > self.LINES_PER_PAGE_A4:
                risks[-1]["description"] = (
                    f"Paragraph spans ~{int(estimated_lines)} lines, exceeds one page. "
                    f"Consider breaking into multiple paragraphs."
                )

            # Check for single-line at page end (widow)
            if estimated_lines == 1 and i == len(self.doc.paragraphs):
                risks.append({
                    "index": i,
                    "estimated_lines": 1,
                    "risk": "low",
                    "description": "Single-line paragraph at document end",
                })

        return risks

    def _analyze_image_placement(self) -> List[Dict]:
        issues = []
        for table in self.doc.tables:
            if table.row_count == 1 and table.col_count == 1:
                cell = table.cell(0, 0)
                if cell:
                    text = cell.text.strip().lower()
                    if any(kw in text for kw in ("image", "picture", "drawing", "jpg", "png", "图", "像")):
                        issues.append({
                            "type": "possible_image_in_table",
                            "description": "Image may be embedded in a 1x1 table cell",
                        })
        return issues

    def _collect_layout_issues(self, overflow: List[Dict], images: List[Dict]) -> List[LayoutIssue]:
        issues = []

        for o in overflow:
            if o.get("risk") == "high":
                issues.append(LayoutIssue(
                    issue_type="text_overflow",
                    severity="warning",
                    location=f"paragraph {o['index']}",
                    message=o.get("description", "Long paragraph may cause layout issues"),
                    suggested_fix="Break long paragraph into smaller paragraphs or adjust margins",
                ))

        for img in images:
            issues.append(LayoutIssue(
                issue_type=img.get("type", "image_placement"),
                severity="warning",
                location="document",
                message=img.get("description", "Image placement issue"),
                suggested_fix="Verify image positioning and text wrapping",
            ))

        # Widow/orphan detection
        roles = self._get_heading_roles()
        for i, para in enumerate(self.doc.paragraphs):
            text = para.text.strip()
            if not text:
                continue
            # Heading at the end of a "page": if a heading is followed by very little content
            if roles.get(i + 1, "").startswith("heading_"):
                following_text = ""
                for j in range(i + 1, min(i + 5, len(self.doc.paragraphs))):
                    following_text += self.doc.paragraphs[j].text.strip()
                if len(following_text) < 20:
                    issues.append(LayoutIssue(
                        issue_type="orphan_heading",
                        severity="warning",
                        location=f"paragraph {i + 1}",
                        message=f"Heading '{text[:40]}' has very little content following it",
                        suggested_fix="Ensure heading has sufficient body text below it, or remove if unused",
                    ))

        return issues

    def _get_heading_roles(self) -> Dict[int, str]:
        roles = {}
        for i, para in enumerate(self.doc.paragraphs, 1):
            if para.is_heading():
                roles[i] = f"heading_{para.heading_level()}"
        return roles

    def _detect_table_overflow(self) -> List[Dict]:
        issues = []
        for t_idx, table in enumerate(self.doc.tables, 1):
            if table.row_count > 30:
                issues.append({
                    "table_index": t_idx, "row_count": table.row_count, "risk": "high",
                    "description": f"Table with {table.row_count} rows may break across pages. Enable 'Repeat Header Rows'.",
                })
            elif table.row_count > 15:
                issues.append({
                    "table_index": t_idx, "row_count": table.row_count, "risk": "medium",
                    "description": f"Table with {table.row_count} rows spans pages. Verify header row repeats.",
                })
        return issues

    def _detect_column_imbalance(self) -> List[Dict]:
        issues = []
        for s_idx, sec in enumerate(self.doc.sections, 1):
            if sec.cols > 1:
                issues.append({
                    "section_index": s_idx, "columns": sec.cols, "column_gap": sec.col_space,
                    "description": f"Section {s_idx} has {sec.cols} columns. Verify balance.",
                })
        return issues
