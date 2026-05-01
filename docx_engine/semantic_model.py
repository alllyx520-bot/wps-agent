# -*- coding: utf-8 -*-
"""
Semantic Document Model — rule-based semantic parser for Office documents.
Builds on top of the Document DOM to produce structured semantic understanding:
  - Element role classification with confidence scores
  - Document relationship graph (heading→body, caption→media, reference→citation)
  - Semantic outline tree
"""
import re
from typing import List, Dict, Optional, Any, Tuple
from dataclasses import dataclass, field

from .document_model import Document, Paragraph


# ─── Semantic Role Enumeration ───

class SemanticRole:
    COVER_TITLE = "cover_title"
    COVER_SUBTITLE = "cover_subtitle"
    COVER_DATE = "cover_date"
    COVER_AUTHOR = "cover_author"
    COVER_INSTITUTION = "cover_institution"
    ABSTRACT_LABEL = "abstract_label"
    ABSTRACT_CONTENT = "abstract_content"
    KEYWORDS_LABEL = "keywords_label"
    KEYWORDS = "keywords"
    TOC_HEADING = "toc_heading"
    TOC_ENTRY = "toc_entry"
    CHAPTER_TITLE = "chapter_title"
    SECTION_TITLE = "section_title"
    SUBSECTION_TITLE = "subsection_title"
    HEADING_1 = "heading_1"
    HEADING_2 = "heading_2"
    HEADING_3 = "heading_3"
    HEADING_4 = "heading_4"
    BODY = "body"
    FIGURE_CAPTION = "figure_caption"
    TABLE_CAPTION = "table_caption"
    EQUATION = "equation"
    REFERENCE_SECTION_HEADER = "reference_section_header"
    REFERENCE_ITEM = "reference_item"
    ACKNOWLEDGEMENTS = "acknowledgements"
    APPENDIX_HEADING = "appendix_heading"
    APPENDIX_CONTENT = "appendix_content"
    LIST_ITEM = "list_item"
    LIST_ITEM_LEVEL_2 = "list_item_level_2"
    EMPTY = "empty"
    UNKNOWN = "unknown"


ROLE_HIERARCHY = {
    SemanticRole.COVER_TITLE: "cover",
    SemanticRole.COVER_SUBTITLE: "cover",
    SemanticRole.COVER_DATE: "cover",
    SemanticRole.COVER_AUTHOR: "cover",
    SemanticRole.COVER_INSTITUTION: "cover",
    SemanticRole.ABSTRACT_LABEL: "front_matter",
    SemanticRole.ABSTRACT_CONTENT: "front_matter",
    SemanticRole.KEYWORDS_LABEL: "front_matter",
    SemanticRole.KEYWORDS: "front_matter",
    SemanticRole.TOC_HEADING: "front_matter",
    SemanticRole.TOC_ENTRY: "front_matter",
    SemanticRole.CHAPTER_TITLE: "heading",
    SemanticRole.SECTION_TITLE: "heading",
    SemanticRole.SUBSECTION_TITLE: "heading",
    SemanticRole.HEADING_1: "heading",
    SemanticRole.HEADING_2: "heading",
    SemanticRole.HEADING_3: "heading",
    SemanticRole.HEADING_4: "heading",
    SemanticRole.BODY: "body",
    SemanticRole.FIGURE_CAPTION: "caption",
    SemanticRole.TABLE_CAPTION: "caption",
    SemanticRole.EQUATION: "caption",
    SemanticRole.REFERENCE_SECTION_HEADER: "back_matter",
    SemanticRole.REFERENCE_ITEM: "back_matter",
    SemanticRole.ACKNOWLEDGEMENTS: "back_matter",
    SemanticRole.APPENDIX_HEADING: "back_matter",
    SemanticRole.APPENDIX_CONTENT: "back_matter",
    SemanticRole.LIST_ITEM: "body",
    SemanticRole.LIST_ITEM_LEVEL_2: "body",
    SemanticRole.EMPTY: "separator",
    SemanticRole.UNKNOWN: "body",
}


# ─── Semantic Element ───

@dataclass
class SemanticElement:
    index: int
    role: str
    text_preview: str
    confidence: float  # 0.0 - 1.0
    heading_level: Optional[int] = None
    style_id: Optional[str] = None
    outline_level: Optional[int] = None
    parent_index: Optional[int] = None  # parent heading paragraph index
    child_indices: List[int] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "index": self.index,
            "role": self.role,
            "text_preview": self.text_preview,
            "confidence": self.confidence,
            "heading_level": self.heading_level,
            "style_id": self.style_id,
            "parent_index": self.parent_index,
            "child_indices": self.child_indices,
            "metadata": self.metadata,
        }


# ─── Document Graph ───

@dataclass
class DocumentGraph:
    elements: List[SemanticElement]
    heading_tree: List[Dict]  # recursive heading hierarchy
    caption_to_media: Dict[int, int]  # caption para index → following table/figure para
    reference_map: Dict[str, int]  # reference label → paragraph containing it
    cover_indices: List[int]
    front_matter_indices: List[int]
    body_indices: List[int]
    back_matter_indices: List[int]

    def get_element(self, index: int) -> Optional[SemanticElement]:
        for e in self.elements:
            if e.index == index:
                return e
        return None

    def get_children(self, index: int) -> List[SemanticElement]:
        children = []
        for e in self.elements:
            if e.parent_index == index:
                children.append(e)
        return children

    def get_by_role(self, role: str) -> List[SemanticElement]:
        return [e for e in self.elements if e.role == role]

    def get_section_body(self, heading_index: int) -> List[SemanticElement]:
        """Get all body elements belonging to a section headed by the given paragraph."""
        results = []
        collecting = False
        heading_level = None
        for e in self.elements:
            if e.index == heading_index:
                collecting = True
                heading_level = e.heading_level
                continue
            if collecting and e.role.startswith("heading_") and e.heading_level and heading_level and e.heading_level <= heading_level:
                break
            if collecting:
                results.append(e)
        return results

    def to_dict(self) -> Dict[str, Any]:
        return {
            "elements": [e.to_dict() for e in self.elements],
            "heading_tree": self.heading_tree,
            "caption_to_media": self.caption_to_media,
            "reference_map": self.reference_map,
            "cover_indices": self.cover_indices,
            "front_matter_indices": self.front_matter_indices,
            "body_indices": self.body_indices,
            "back_matter_indices": self.back_matter_indices,
        }


# ─── Semantic Parser ───

class SemanticParser:
    """Rule-based semantic document parser with confidence scoring."""

    def __init__(self, doc: Document):
        self.doc = doc
        self._results: List[SemanticElement] = []

    def parse(self) -> DocumentGraph:
        self._results = []
        self._classify_all()
        self._apply_contextual_rules()
        self._build_relationships()
        return self._build_graph()

    def _classify_all(self):
        for i, para in enumerate(self.doc.paragraphs, 1):
            role, confidence = self._classify_paragraph(para, i)
            self._results.append(SemanticElement(
                index=i,
                role=role,
                text_preview=para.text.strip()[:80],
                confidence=confidence,
                heading_level=para.heading_level(),
                style_id=para.style_id,
                outline_level=para.outline_level,
            ))

    def _classify_paragraph(self, para: Paragraph, index: int) -> Tuple[str, float]:
        text = para.text.strip()
        total = len(self.doc.paragraphs)

        if not text:
            return SemanticRole.EMPTY, 1.0

        # ── Ordered rule cascade (first match wins) ──

        pos_ratio = index / total if total > 0 else 0

        # 1. Cover page detection (front portion or first 5 paragraphs)
        in_cover_zone = pos_ratio <= 0.15 or index <= 5
        if in_cover_zone:
            # Skip cover rules for known non-cover patterns
            _lower = text.lower()
            cover_skip = False
            if re.match(r'^(abstract|摘要)\s*[:：]?\s*$', _lower, re.IGNORECASE):
                cover_skip = True
            elif re.match(r'^\s*(参考文献|References|Bibliography|致谢|Acknowledgement)\s*$', text, re.IGNORECASE):
                cover_skip = True
            elif re.match(r'^(第[一二三四五六七八九十\d]+章|[Cc]hapter\s*\d+)', text):
                cover_skip = True
            elif re.match(r'^(关键词|关键字|keywords|key\s*words)\s*[：:]', _lower, re.IGNORECASE):
                cover_skip = True
            elif re.match(r'^\[\d+\]', text):
                cover_skip = True
            if not cover_skip:
                if para.alignment in ("center", "both"):
                    large_bold = any(r.bold or (r.size and r.size >= 22) for r in para.runs)
                    medium_bold = any(r.bold or (r.size and r.size >= 14) for r in para.runs)
                    if large_bold:
                        return SemanticRole.COVER_TITLE, 0.9
                    if medium_bold and index <= 5:
                        return SemanticRole.COVER_SUBTITLE, 0.75
                date_pattern = r'^\d{4}\s*[年/\-\.]\s*\d{1,2}\s*[月/\-\.]?\s*\d{0,2}\s*[日号]?$'
                if re.match(date_pattern, text):
                    return SemanticRole.COVER_DATE, 0.95
                if re.match(r'^\d{4}年\d{1,2}月\d{1,2}日$', text):
                    return SemanticRole.COVER_DATE, 0.98
                author_pattern = r'^[\u4e00-\u9fff]{2,4}\s*[：:]\s*.{2,20}$'
                if re.match(author_pattern, text):
                    return SemanticRole.COVER_AUTHOR, 0.8
                if re.match(r'^[\u4e00-\u9fff]{2,4}\s+[\u4e00-\u9fff]{2,4}$', text) and index <= 3:
                    return SemanticRole.COVER_AUTHOR, 0.7
                if re.match(r'^[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]{4,20}$', text) and index <= 2:
                    if not any(c in text for c in "。，；！？"):
                        return SemanticRole.COVER_INSTITUTION, 0.65

        # 2. References (MUST be checked BEFORE heading, as they often have heading style)
        lower = text.lower()
        if re.match(r'^\s*(参考文献|References|Bibliography|References\s*and\s*Notes)\s*$', text, re.IGNORECASE):
            return SemanticRole.REFERENCE_SECTION_HEADER, 0.98

        # 3. Acknowledgements (also before heading check)
        if re.match(r'^\s*(致谢|Acknowledgements?|Acknowledgment)\s*$', text, re.IGNORECASE):
            return SemanticRole.ACKNOWLEDGEMENTS, 0.98

        # 4. Heading detection
        if para.is_heading():
            level = para.heading_level()
            if level == 1:
                if re.match(r'^(第[一二三四五六七八九十\d]+章|[Cc]hapter\s*\d+)', text):
                    return SemanticRole.CHAPTER_TITLE, 0.95
                return SemanticRole.HEADING_1, 0.9
            elif level == 2:
                return SemanticRole.HEADING_2, 0.9
            elif level == 3:
                return SemanticRole.HEADING_3, 0.9
            elif level == 4:
                return SemanticRole.HEADING_4, 0.9
            else:
                return f"heading_{level}", 0.85

        # 5. Abstract
        lower = text.lower()
        if re.match(r'^(abstract|摘要)\s*[:：]?\s*$', lower, re.IGNORECASE):
            return SemanticRole.ABSTRACT_LABEL, 0.98

        # 6. Keywords
        kw_patterns = [
            r'^(keywords|key\s*words|关键词|关键字)\s*[:：]?\s*',
            r'^(keywords|key\s*words|关键词|关键字)\s*[:：].+',
        ]
        for pat in kw_patterns:
            if re.match(pat, lower, re.IGNORECASE):
                return SemanticRole.KEYWORDS_LABEL, 0.98

        # 7. TOC
        if re.match(r'^\s*(目录|Table\s*of\s*Contents|Contents)\s*$', text, re.IGNORECASE):
            return SemanticRole.TOC_HEADING, 0.98
        if re.match(r'^.{1,80}\.{3,}\s*\d+$', text):
            return SemanticRole.TOC_ENTRY, 0.95

        # 8. Figure/Table captions
        if re.match(r'^(图|Figure|Fig\.?)\s*\d+', text, re.IGNORECASE):
            return SemanticRole.FIGURE_CAPTION, 0.95
        if re.match(r'^(表|Table)\s*\d+', text, re.IGNORECASE):
            return SemanticRole.TABLE_CAPTION, 0.95

        # 9. Equations
        if re.match(r'^[（(]?\s*(方程|公式|Equation|Formula)\s*\(?\d*\)?\s*[）)]?\s*[:：]?', text, re.IGNORECASE):
            return SemanticRole.EQUATION, 0.9

        # 10. Reference items (section header already caught earlier)
        if re.match(r'^\[\d+\]', text):
            return SemanticRole.REFERENCE_ITEM, 0.95
        if re.match(r'^\d+\.\s+\w', text) and pos_ratio > 0.8:
            return SemanticRole.REFERENCE_ITEM, 0.7

        # 11. Appendix
        if re.match(r'^\s*(附录|Appendix)\s*[A-Za-z]*\s*$', text, re.IGNORECASE):
            return SemanticRole.APPENDIX_HEADING, 0.98
        if re.match(r'^\s*附录[一二三四五六七八九十]|Appendix\s+[IVXLC]+', text, re.IGNORECASE):
            return SemanticRole.APPENDIX_CONTENT, 0.85

        # 12. List items
        if para.numPr and para.numPr.get("numId", 0) > 0:
            ilvl = para.numPr.get("ilvl", 0)
            if ilvl == 0:
                return SemanticRole.LIST_ITEM, 0.9
            return SemanticRole.LIST_ITEM_LEVEL_2, 0.85
        bullet_patterns = [r'^[•·\-\*]\s', r'^\d+[\.\)]\s', r'^[\(\[]\d+[\)\]]\s',
                          r'^[（(][一二三四五六七八九十]+[）)]']
        for pat in bullet_patterns:
            if re.match(pat, text):
                return SemanticRole.LIST_ITEM, 0.85

        # 13. Body text (default, high confidence when long)
        if len(text) > 50:
            return SemanticRole.BODY, 0.75
        if len(text) > 15:
            return SemanticRole.BODY, 0.6

        return SemanticRole.UNKNOWN, 0.3

    def _apply_contextual_rules(self):
        """Post-classification context-aware refinements."""
        for i, elem in enumerate(self._results):
            idx = elem.index

            # Abstract content: paragraph AFTER abstract label
            if i > 0 and self._results[i - 1].role == SemanticRole.ABSTRACT_LABEL:
                if elem.role in (SemanticRole.BODY, SemanticRole.UNKNOWN):
                    elem.role = SemanticRole.ABSTRACT_CONTENT
                    elem.confidence = 0.85

            # Keywords content: paragraph AFTER keywords label
            if i > 0 and self._results[i - 1].role == SemanticRole.KEYWORDS_LABEL:
                if elem.role in (SemanticRole.BODY, SemanticRole.UNKNOWN):
                    elem.role = SemanticRole.KEYWORDS
                    elem.confidence = 0.85

            # Reference items: after reference section header
            if i > 0:
                prev = self._results[i - 1]
                if prev.role == SemanticRole.REFERENCE_SECTION_HEADER:
                    if elem.role in (SemanticRole.BODY, SemanticRole.UNKNOWN):
                        if re.match(r'^\[\d+\]', elem.text_preview):
                            elem.role = SemanticRole.REFERENCE_ITEM
                            elem.confidence = 0.9

            # Appendix content: after appendix heading
            if i > 0:
                prev = self._results[i - 1]
                if prev.role == SemanticRole.APPENDIX_HEADING:
                    if elem.role in (SemanticRole.BODY, SemanticRole.UNKNOWN):
                        elem.role = SemanticRole.APPENDIX_CONTENT
                        elem.confidence = 0.8

            # Downgrade unknown short text near references
            if elem.role == SemanticRole.UNKNOWN:
                if re.match(r'^\d+\.\s+\w{2,}', elem.text_preview) and idx / len(self._results) > 0.7:
                    elem.role = SemanticRole.REFERENCE_ITEM
                    elem.confidence = 0.65

    def _build_relationships(self):
        """Build parent-child relationships between headings and their content."""
        heading_stack: List[SemanticElement] = []

        for elem in self._results:
            role = elem.role
            if role in (SemanticRole.HEADING_1, SemanticRole.HEADING_2,
                         SemanticRole.HEADING_3, SemanticRole.HEADING_4,
                         SemanticRole.CHAPTER_TITLE, SemanticRole.SECTION_TITLE,
                         SemanticRole.SUBSECTION_TITLE):
                level = elem.heading_level or 0
                while heading_stack:
                    parent = heading_stack[-1]
                    parent_level = parent.heading_level or 0
                    if level > parent_level:
                        break
                    heading_stack.pop()
                if heading_stack:
                    elem.parent_index = heading_stack[-1].index
                    heading_stack[-1].child_indices.append(elem.index)
                heading_stack.append(elem)
            elif heading_stack:
                elem.parent_index = heading_stack[-1].index
                heading_stack[-1].child_indices.append(elem.index)

    def _build_heading_tree(self) -> List[Dict]:
        """Build recursive heading tree for outline display."""
        headings = [e for e in self._results if e.role.startswith("heading_") or
                     e.role in (SemanticRole.CHAPTER_TITLE, SemanticRole.SECTION_TITLE,
                                SemanticRole.SUBSECTION_TITLE)]

        def build_subtree(parent_idx: Optional[int] = None) -> List[Dict]:
            nodes = []
            for h in headings:
                if h.parent_index == parent_idx:
                    node = {
                        "index": h.index,
                        "text": h.text_preview,
                        "role": h.role,
                        "level": h.heading_level,
                        "children": build_subtree(h.index),
                    }
                    nodes.append(node)
            return nodes

        return build_subtree(None)

    def _build_graph(self) -> DocumentGraph:
        heading_tree = self._build_heading_tree()

        caption_to_media: Dict[int, int] = {}
        for i, elem in enumerate(self._results):
            if elem.role in (SemanticRole.FIGURE_CAPTION, SemanticRole.TABLE_CAPTION):
                if i + 1 < len(self._results):
                    next_elem = self._results[i + 1]
                    caption_to_media[elem.index] = next_elem.index

        reference_map: Dict[str, int] = {}
        for elem in self._results:
            if elem.role == SemanticRole.REFERENCE_ITEM:
                m = re.match(r'^\[(\d+)\]', elem.text_preview)
                if m:
                    reference_map[m.group(1)] = elem.index

        cover = [e.index for e in self._results if e.role.startswith("cover_")]
        front = [e.index for e in self._results if ROLE_HIERARCHY.get(e.role) == "front_matter"]
        body = [e.index for e in self._results if ROLE_HIERARCHY.get(e.role) in ("body", "heading", "caption", "separator")]
        back = [e.index for e in self._results if ROLE_HIERARCHY.get(e.role) == "back_matter"]

        return DocumentGraph(
            elements=self._results,
            heading_tree=heading_tree,
            caption_to_media=caption_to_media,
            reference_map=reference_map,
            cover_indices=cover,
            front_matter_indices=front,
            body_indices=body,
            back_matter_indices=back,
        )


# ─── Content Type Classification ───

CONTENT_TYPES = {
    "expository": "论述型",
    "data": "数据型",
    "formula": "公式型",
    "code": "代码型",
    "reference": "引用型",
}


def classify_paragraph_content(text: str) -> Dict[str, Any]:
    """Classify a paragraph by its content nature."""
    if not text or not text.strip():
        return {"type": "empty", "confidence": 1.0}
    stripped = text.strip()
    if re.match(r'^\s*[\[（(]\s*\d+[\d\.\-+*/\^×÷±\s]+\s*[\]）)]\s*$', stripped):
        return {"type": "formula", "confidence": 0.85}
    if re.match(r'^[\[（(]\s*\d+\s*[\]）)]\s*$', stripped):
        return {"type": "formula_label", "confidence": 0.9}
    code_signals = [r'(def|class|import|from|return|if|else|for|while|try|except)\s',
                    r'\{.*\}', r'function\s*\(', r'const\s+\w+\s*=',
                    r'<\w+[^>]*>', r'SELECT\s+.*FROM', r'INSERT\s+INTO']
    for pat in code_signals:
        if re.search(pat, stripped):
            return {"type": "code", "confidence": 0.8}
    if re.search(r'[\d\.]+\s*[±＋]\s*[\d\.]+', stripped):
        return {"type": "data", "confidence": 0.75}
    if re.search(r'^\s*[\d\s\.\,\;%％]+\s*$', stripped):
        return {"type": "data", "confidence": 0.7}
    if re.match(r'^\s*[\[（]\d+[\];,，；\s\d]+[\]）]\s*$', stripped):
        return {"type": "reference", "confidence": 0.85}
    return {"type": "expository", "confidence": 0.65}
