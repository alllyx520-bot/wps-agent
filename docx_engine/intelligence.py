# -*- coding: utf-8 -*-
"""
Document Intelligence — analyze document structure, detect paragraph roles,
classify document type, and suggest formatting improvements.
"""
import re
from typing import List, Dict, Optional, Any

from .document_model import Document, Paragraph
from .style_resolver import StyleResolver


class DocumentAnalyzer:
    """Analyzes document content and structure to provide intelligent insights."""

    def __init__(self, doc: Document, style_resolver: Optional[StyleResolver] = None):
        self.doc = doc
        self.resolver = style_resolver

    # ─── Document Type Detection ───

    def detect_document_type(self) -> str:
        """Classify document type based on content and structure."""
        text = self.doc.text[:5000].lower()
        headings = self.doc.get_heading_structure()

        scores = {
            "thesis": 0,
            "report": 0,
            "resume": 0,
            "contract": 0,
            "letter": 0,
            "memo": 0,
            "paper": 0,
        }

        # Thesis indicators
        thesis_keywords = ["abstract", "keywords", "introduction", "literature review",
                          "methodology", "results", "discussion", "conclusion",
                          "references", "acknowledgment", "chapter"]
        for kw in thesis_keywords:
            if kw in text:
                scores["thesis"] += 1

        # Resume indicators
        resume_keywords = ["education", "experience", "skills", "contact",
                          "phone", "email", "linkedin", "objective", "profile"]
        for kw in resume_keywords:
            if kw in text:
                scores["resume"] += 1

        # Contract indicators
        contract_keywords = ["agreement", "party", "clause", "terms", "conditions",
                            "warranty", "liability", "jurisdiction", "hereby"]
        for kw in contract_keywords:
            if kw in text:
                scores["contract"] += 1

        # Letter indicators
        letter_patterns = [r"dear\s+\w+", r"sincerely", r"regards", r"yours truly"]
        for pat in letter_patterns:
            if re.search(pat, text):
                scores["letter"] += 1

        # Memo indicators
        if "to:" in text and "from:" in text and "subject:" in text:
            scores["memo"] += 3

        # Paper indicators
        paper_keywords = ["abstract", "keywords", "introduction", "method",
                         "results", "discussion", "conclusion", "doi", "et al"]
        for kw in paper_keywords:
            if kw in text:
                scores["paper"] += 1

        # Report indicators
        if len(headings) >= 3:
            scores["report"] += 2
        report_keywords = ["executive summary", "findings", "recommendations",
                          "appendix", "figure", "table"]
        for kw in report_keywords:
            if kw in text:
                scores["report"] += 1

        best = max(scores, key=scores.get)
        return best if scores[best] > 0 else "general"

    # ─── Paragraph Role Detection ───

    def detect_paragraph_roles(self) -> List[Dict[str, Any]]:
        """Assign a semantic role to each paragraph."""
        results = []
        for i, para in enumerate(self.doc.paragraphs, 1):
            role = self._detect_single_role(para, i)
            results.append({
                "index": i,
                "role": role,
                "text_preview": para.text[:80],
                "style_id": para.style_id,
                "heading_level": para.heading_level(),
            })
        return results

    def _detect_single_role(self, para: Paragraph, index: int) -> str:
        text = para.text.strip()
        lower = text.lower()

        # Empty paragraph
        if not text:
            return "empty"

        # Heading
        if para.is_heading():
            level = para.heading_level()
            if level == 1:
                return "heading1"
            elif level == 2:
                return "heading2"
            elif level == 3:
                return "heading3"
            return f"heading{level}"

        # Title detection (first non-empty paragraph, centered/bold/large)
        if index <= 3:
            if para.alignment in ("center", "both"):
                if any(r.bold or (r.size and r.size >= 16) for r in para.runs):
                    return "title"

        # Abstract
        if re.match(r"^(abstract|摘要)\s*[:：]?\s*$", lower):
            return "abstract_label"
        if index > 1:
            prev = self.doc.get_paragraph(index - 1)
            if prev and re.match(r"^(abstract|摘要)\s*[:：]?\s*$", prev.text.strip().lower()):
                return "abstract_content"

        # Keywords
        if re.match(r"^(keywords|关键词)\s*[:：]", lower):
            return "keywords"

        # Table caption
        if re.match(r"^(table|表)\s*\d+", lower):
            return "table_caption"

        # Figure caption
        if re.match(r"^(figure|fig\.?|图)\s*\d+", lower):
            return "figure_caption"

        # List items
        if para.numPr and para.numPr.get("numId", 0) > 0:
            ilvl = para.numPr.get("ilvl", 0)
            if ilvl == 0:
                return "list_item"
            return f"list_item_level_{ilvl + 1}"

        # Bullet patterns (if not using native numbering)
        bullet_patterns = [r"^[•·\-\*]\s", r"^\d+[\.\)]\s", r"^[\(\[]\d+[\)\]]\s"]
        for pat in bullet_patterns:
            if re.match(pat, text):
                return "list_item"

        # Reference section
        if re.match(r"^(references|bibliography|参考文献)\s*$", lower):
            return "references_label"
        if index > 1:
            prev = self.doc.get_paragraph(index - 1)
            if prev and re.match(r"^(references|bibliography|参考文献)\s*$", prev.text.strip().lower()):
                return "reference_entry"

        # Footnote/Endnote references in text
        if re.search(r"\[\d+\]$", text) or re.search(r"\(\d+\)$", text):
            return "reference_entry"

        # Acknowledgment
        if re.match(r"^(acknowledgment|acknowledgement|致谢)\s*$", lower):
            return "acknowledgment_label"

        # Body text heuristics
        if len(text) > 50 and text[-1] in ".。!！?？":
            return "body"

        return "unknown"

    # ─── Formatting Quality Analysis ───

    def analyze_formatting_quality(self) -> Dict[str, Any]:
        """Analyze formatting consistency and detect common issues."""
        issues = []
        stats = {
            "total_paragraphs": len(self.doc.paragraphs),
            "total_tables": len(self.doc.tables),
            "headings": 0,
            "inconsistent_fonts": 0,
            "inconsistent_spacing": 0,
            "orphan_headings": 0,
            "empty_paragraphs": 0,
        }

        # Collect font usage
        font_sizes = []
        for para in self.doc.iter_all_paragraphs():
            for run in para.runs:
                if run.size:
                    font_sizes.append(run.size)

        # Check heading consistency
        headings = self.doc.get_heading_structure()
        stats["headings"] = len(headings)

        prev_level = 0
        for h in headings:
            level = h["level"]
            # Check heading sequence gaps
            if level > prev_level + 1 and prev_level > 0:
                issues.append({
                    "type": "heading_gap",
                    "severity": "warning",
                    "index": h["index"],
                    "message": f"Heading level jumps from {prev_level} to {level}",
                })
            prev_level = level

        # Check orphan headings (heading at end of section without content)
        for h in headings:
            idx = h["index"]
            if idx < len(self.doc.paragraphs):
                next_para = self.doc.paragraphs[idx]  # idx is 1-based, paragraphs list is 0-based; +1 offset gives "next"
                if next_para.is_heading() or next_para.is_empty():
                    stats["orphan_headings"] += 1
                    issues.append({
                        "type": "orphan_heading",
                        "severity": "warning",
                        "index": idx,
                        "message": f"Heading '{h['text'][:40]}' may be orphaned at paragraph {idx}",
                    })

        # Check empty paragraphs
        for i, para in enumerate(self.doc.paragraphs, 1):
            if para.is_empty():
                stats["empty_paragraphs"] += 1

        # Check font consistency in body
        if font_sizes:
            from statistics import mode, stdev
            try:
                common_size = mode(font_sizes)
                stats["common_font_size"] = common_size
                outliers = [s for s in font_sizes if abs(s - common_size) > 2]
                if outliers:
                    stats["inconsistent_fonts"] = len(outliers)
                    issues.append({
                        "type": "font_inconsistency",
                        "severity": "info",
                        "message": f"Found {len(outliers)} font size outliers from common size {common_size}pt",
                    })
            except Exception:
                pass

        return {
            "score": max(0, 100 - len(issues) * 5),
            "stats": stats,
            "issues": issues,
        }

    # ─── Structure Analysis ───

    def get_document_outline(self) -> List[Dict[str, Any]]:
        """Get a hierarchical outline of the document."""
        outline = []
        stack = []

        for i, para in enumerate(self.doc.paragraphs, 1):
            level = para.heading_level()
            if level is None:
                continue

            node = {
                "index": i,
                "level": level,
                "text": para.text,
                "children": [],
            }

            # Pop stack to correct level
            while stack and stack[-1]["level"] >= level:
                stack.pop()

            if stack:
                stack[-1]["children"].append(node)
            else:
                outline.append(node)

            stack.append(node)

        return outline

    def get_toc_candidates(self) -> List[Dict[str, Any]]:
        """Generate table of contents from headings."""
        toc = []
        for h in self.doc.get_heading_structure():
            toc.append({
                "level": h["level"],
                "text": h["text"],
                "page": None,  # Would need layout engine to determine
            })
        return toc
