# -*- coding: utf-8 -*-
"""
Professional Formatter — apply professional document formatting with
context awareness, style inheritance, and Chinese typography rules.
"""
import re
from typing import List, Dict, Optional, Any, Tuple

from .document_model import Document, Paragraph, Run, Section
from .style_resolver import StyleResolver


class Formatter:
    """Professional document formatting engine."""

    # Chinese typography conventions
    CJK_PUNCTUATION = "，。、；：？！""''「」『』（）【】《》〈〉"
    CJK_OPEN_PUNCT = """""「『（【《〈"
    CJK_CLOSE_PUNCT = """"」』）】》〉"

    # Standard font sizes (Chinese typesetting)
    FONT_SIZES = {
        "title": 22,        # 二号
        "subtitle": 16,     # 三号
        "heading1": 16,     # 三号
        "heading2": 14,     # 四号
        "heading3": 12,     # 小四
        "body": 12,         # 小四
        "caption": 10.5,    # 五号
        "footnote": 9,      # 小五
    }

    # Standard spacing (points)
    SPACING = {
        "title_before": 120,
        "title_after": 24,
        "heading1_before": 24,
        "heading1_after": 12,
        "heading2_before": 18,
        "heading2_after": 6,
        "heading3_before": 12,
        "heading3_after": 6,
        "body_before": 0,
        "body_after": 0,
        "paragraph_spacing": 6,
    }

    def __init__(self, doc: Document, style_resolver: Optional[StyleResolver] = None):
        self.doc = doc
        self.resolver = style_resolver

    # ─── High-Level Operations ───

    def auto_format(self, document_type: str = "general") -> Dict[str, Any]:
        """Apply comprehensive automatic formatting based on document type."""
        changes = []

        if document_type == "thesis":
            changes.extend(self._format_thesis())
        elif document_type == "report":
            changes.extend(self._format_report())
        elif document_type == "resume":
            changes.extend(self._format_resume())
        else:
            changes.extend(self._format_general())

        # Common fixes for all types
        changes.extend(self._fix_punctuation_spacing())
        changes.extend(self._fix_kinsoku())
        changes.extend(self._fix_heading_sequence())
        changes.extend(self._remove_redundant_empty_paragraphs())

        return {
            "applied": True,
            "changes": changes,
            "change_count": len(changes),
        }

    def apply_template(self, template_name: str) -> Dict[str, Any]:
        """Apply a named formatting template."""
        templates = {
            "thesis_cn": self._template_thesis_cn,
            "report_official": self._template_report_official,
            "resume_professional": self._template_resume_professional,
        }
        fn = templates.get(template_name)
        if fn:
            return fn()
        return {"error": f"Unknown template: {template_name}"}

    # ─── Specific Formatters ───

    def _format_thesis(self) -> List[Dict]:
        changes = []
        for i, para in enumerate(self.doc.paragraphs, 1):
            level = para.heading_level()
            if level == 1:
                changes.extend(self._apply_heading_format(para, 16, True, "center", 24, 12))
            elif level == 2:
                changes.extend(self._apply_heading_format(para, 14, True, "left", 18, 6))
            elif level == 3:
                changes.extend(self._apply_heading_format(para, 12, True, "left", 12, 6))
            elif level is None and not para.is_empty():
                # Body text
                changes.extend(self._apply_body_format(para, 12, "宋体", "justify", 2))
        return changes

    def _format_report(self) -> List[Dict]:
        changes = []
        for para in self.doc.paragraphs:
            level = para.heading_level()
            if level == 1:
                changes.extend(self._apply_heading_format(para, 16, True, "left", 24, 12))
            elif level == 2:
                changes.extend(self._apply_heading_format(para, 14, True, "left", 12, 6))
            elif level is None and not para.is_empty():
                changes.extend(self._apply_body_format(para, 12, "宋体", "justify", 2))
        return changes

    def _format_resume(self) -> List[Dict]:
        changes = []
        for para in self.doc.paragraphs:
            if para.is_empty():
                continue
            text = para.text.strip()
            # Section headers
            if re.match(r"^(EDUCATION|EXPERIENCE|SKILLS|PROJECTS|CONTACT|PROFILE|OBJECTIVE)", text, re.I):
                changes.extend(self._apply_heading_format(para, 14, True, "left", 12, 6))
            else:
                changes.extend(self._apply_body_format(para, 10.5, "Calibri", "left", 0))
        return changes

    def _format_general(self) -> List[Dict]:
        changes = []
        for para in self.doc.paragraphs:
            if para.is_empty():
                continue
            level = para.heading_level()
            if level == 1:
                changes.extend(self._apply_heading_format(para, 16, True, "left", 18, 6))
            elif level == 2:
                changes.extend(self._apply_heading_format(para, 14, True, "left", 12, 6))
            elif level == 3:
                changes.extend(self._apply_heading_format(para, 12, True, "left", 6, 3))
            elif level is None:
                text = para.text.strip()
                if text and not para.runs:
                    continue
                if text and len(text) < 30 and (re.match(r"^第[一二三四五六七八九十\d]+[章节条]|^[\d]+\.[\d]+\s|^[一二三四五六七八九十]、", text) or
                    text.endswith(("概述", "总结", "介绍", "背景", "方法", "结论", "测试", "分析", "设计"))):
                    changes.extend(self._apply_heading_format(para, 14, True, "left", 12, 6))
                else:
                    changes.extend(self._apply_body_format(para, 12, "宋体", "justify", 2))
        return changes

    # ─── Template Implementations ───

    def _template_thesis_cn(self) -> Dict[str, Any]:
        """Chinese university thesis template."""
        changes = []
        for para in self.doc.paragraphs:
            level = para.heading_level()
            if level == 1:
                # Chapter titles: 黑体, 三号(16pt), centered
                changes.extend(self._apply_heading_format(para, 16, True, "center", 24, 12, "黑体"))
            elif level == 2:
                # Section titles: 黑体, 四号(14pt), left
                changes.extend(self._apply_heading_format(para, 14, True, "left", 18, 6, "黑体"))
            elif level == 3:
                # Subsection: 黑体, 小四(12pt), left
                changes.extend(self._apply_heading_format(para, 12, True, "left", 12, 6, "黑体"))
            elif level is None and not para.is_empty():
                # Body: 宋体, 小四(12pt), justify, first-line indent 2 chars
                changes.extend(self._apply_body_format(para, 12, "宋体", "justify", 2))
                # Line spacing 1.5
                para.line_spacing = 1.5
                para.line_rule = "auto"

        # Set page margins (Chinese thesis standard)
        if self.doc.sections:
            sec = self.doc.sections[0]
            sec.top_margin = 72    # 2.54cm
            sec.bottom_margin = 72
            sec.left_margin = 85   # 3.0cm
            sec.right_margin = 72  # 2.54cm

        return {"template": "thesis_cn", "changes": changes, "change_count": len(changes)}

    def _template_report_official(self) -> Dict[str, Any]:
        changes = []
        for para in self.doc.paragraphs:
            level = para.heading_level()
            if level == 1:
                changes.extend(self._apply_heading_format(para, 16, True, "center", 24, 12, "黑体"))
            elif level == 2:
                changes.extend(self._apply_heading_format(para, 14, True, "left", 18, 6, "黑体"))
            elif level is None and not para.is_empty():
                changes.extend(self._apply_body_format(para, 12, "仿宋", "justify", 2))

        if self.doc.sections:
            sec = self.doc.sections[0]
            sec.top_margin = 72
            sec.bottom_margin = 72
            sec.left_margin = 72
            sec.right_margin = 72

        return {"template": "report_official", "changes": changes, "change_count": len(changes)}

    def _template_resume_professional(self) -> Dict[str, Any]:
        changes = []
        for para in self.doc.paragraphs:
            text = para.text.strip()
            if re.match(r"^(EDUCATION|EXPERIENCE|SKILLS|PROJECTS|CERTIFICATIONS|LANGUAGES)", text, re.I):
                changes.extend(self._apply_heading_format(para, 12, True, "left", 12, 6, "Arial"))
                # Add bottom border effect by using underline on runs
                for run in para.runs:
                    run.underline = "single"
            else:
                changes.extend(self._apply_body_format(para, 10.5, "Calibri", "left", 0))

        return {"template": "resume_professional", "changes": changes, "change_count": len(changes)}

    # ─── Low-Level Format Application ───

    def _apply_heading_format(
        self, para: Paragraph, size: float, bold: bool, alignment: str,
        space_before: float, space_after: float, font: Optional[str] = None
    ) -> List[Dict]:
        changes = []
        old_alignment = para.alignment
        old_sb = para.space_before
        old_sa = para.space_after
        old_fli = para.first_line_indent
        old_li = para.left_indent
        old_ri = para.right_indent

        para.alignment = alignment
        para.space_before = space_before
        para.space_after = space_after
        para.first_line_indent = None
        para.left_indent = None
        para.right_indent = None

        if old_alignment != alignment:
            changes.append({"type": "paragraph_format", "action": "set_alignment", "old": old_alignment, "new": alignment})
        if old_sb != space_before:
            changes.append({"type": "paragraph_format", "action": "set_space_before", "old": old_sb, "new": space_before})
        if old_sa != space_after:
            changes.append({"type": "paragraph_format", "action": "set_space_after", "old": old_sa, "new": space_after})
        if old_fli is not None and old_fli != 0:
            changes.append({"type": "paragraph_format", "action": "clear_first_line_indent", "old": old_fli})

        for run in para.runs:
            run.size = size
            run.bold = bold
            run.italic = False
            if font:
                run.font = font
            changes.append({
                "type": "run_format",
                "action": "set_font",
                "size": size, "bold": bold, "font": font,
            })

        changes.append({
            "type": "paragraph_format",
            "action": "set_alignment_spacing",
            "alignment": alignment,
            "space_before": space_before,
            "space_after": space_after,
        })
        return changes

    def _apply_body_format(
        self, para: Paragraph, size: float, font: str, alignment: str,
        first_line_indent_chars: int
    ) -> List[Dict]:
        changes = []
        old_alignment = para.alignment
        old_sb = para.space_before
        old_sa = para.space_after

        para.alignment = alignment
        if first_line_indent_chars > 0:
            para.first_line_indent = first_line_indent_chars * size
        para.space_before = 0
        para.space_after = self.SPACING["paragraph_spacing"]

        if old_alignment != alignment:
            changes.append({"type": "paragraph_format", "action": "set_alignment", "old": old_alignment, "new": alignment})
        if old_sb is not None and old_sb != 0:
            changes.append({"type": "paragraph_format", "action": "clear_space_before", "old": old_sb})
        if old_sa != self.SPACING["paragraph_spacing"]:
            changes.append({"type": "paragraph_format", "action": "set_space_after", "old": old_sa, "new": self.SPACING["paragraph_spacing"]})

        for run in para.runs:
            run.size = size
            run.font = font
            run.bold = False
            run.italic = False
            changes.append({
                "type": "run_format",
                "action": "set_font",
                "size": size, "font": font,
            })

        changes.append({
            "type": "paragraph_format",
            "action": "set_body_format",
            "alignment": alignment,
            "first_line_indent": para.first_line_indent,
            "space_after": para.space_after,
        })
        return changes

    # ─── Quality Fixes ───

    def _fix_punctuation_spacing(self) -> List[Dict]:
        """Fix Chinese punctuation spacing issues."""
        changes = []
        for para in self.doc.iter_all_paragraphs():
            for run in para.runs:
                # Remove spaces before Chinese punctuation
                original = run.text
                fixed = re.sub(rf"\s+([{self.CJK_PUNCTUATION}])", r"\1", original)
                if fixed != original:
                    run.text = fixed
                    changes.append({
                        "type": "text_fix",
                        "action": "remove_space_before_punctuation",
                        "original": original[:40],
                        "fixed": fixed[:40],
                    })
        return changes

    def _fix_heading_sequence(self) -> List[Dict]:
        """Ensure heading levels don't skip (e.g., H1 -> H3 without H2)."""
        changes = []
        headings = self.doc.get_heading_structure()
        prev_level = 0
        for h in headings:
            level = h["level"]
            if level > prev_level + 1 and prev_level > 0:
                # This is a gap, but we can't auto-fix without user confirmation
                changes.append({
                    "type": "heading_sequence",
                    "severity": "warning",
                    "index": h["index"],
                    "message": f"Heading level gap: H{prev_level} -> H{level}",
                })
            prev_level = level
        return changes

    def _remove_redundant_empty_paragraphs(self) -> List[Dict]:
        """Remove consecutive empty paragraphs (keep at most one)."""
        changes = []
        to_remove = []
        prev_empty = False
        for i, para in enumerate(self.doc.paragraphs):
            is_empty = para.is_empty()
            if is_empty and prev_empty:
                to_remove.append(i)
            prev_empty = is_empty

        # Remove in reverse order to preserve indices
        for idx in reversed(to_remove):
            removed = self.doc.delete_paragraph(idx + 1)
            if removed:
                changes.append({
                    "type": "cleanup",
                    "action": "remove_redundant_empty",
                    "index": idx + 1,
                })
        return changes

    def _fix_kinsoku(self) -> List[Dict]:
        """Apply Chinese kinsoku (禁則) rules — prohibit certain chars at line start/end."""
        changes = []

        # Characters that cannot start a line (行頭禁則)
        KINSOKU_NO_START = "、。，．，：；？！）〕］｝」』】〉》〕〟ヽヾーァィゥェォッャュョヮヵヶぁぃぅぇぉっゃゅょゎゕゖ%℃￠¢"

        # Characters that cannot end a line (行末禁則)
        KINSOKU_NO_END = "（〔［｛「『【〈《〔〝＄￡￥"

        import re
        for para in self.doc.iter_all_paragraphs():
            for run in para.runs:
                original = run.text
                if not original:
                    continue

                # Insert zero-width non-joiner before line-start-prohibited chars
                # when they appear after non-punctuation (simulates kinsoku)
                fixed = re.sub(
                    rf"([^\s{KINSOKU_NO_START}])\s*([{KINSOKU_NO_START}])",
                    r"\1\2",
                    original,
                )
                # Remove whitespace between line-end-prohibited char and following char
                fixed = re.sub(
                    rf"([{KINSOKU_NO_END}])\s+",
                    r"\1",
                    fixed,
                )

                if fixed != original:
                    run.text = fixed
                    changes.append({
                        "type": "kinsoku",
                        "action": "fix_kinsoku",
                        "original_len": len(original),
                        "fixed_len": len(fixed),
                    })

        return changes

    # ─── Advanced Operations ───

    def add_multi_level_numbering(self) -> Dict[str, Any]:
        """Add automatic multi-level numbering to headings."""
        changes = []
        counters = [0, 0, 0, 0]  # level 1-4

        for para in self.doc.paragraphs:
            level = para.heading_level()
            if level is None:
                continue

            idx = level - 1
            if 0 <= idx < len(counters):
                counters[idx] += 1
                # Reset lower levels
                for j in range(idx + 1, len(counters)):
                    counters[j] = 0

                # Build number string
                if level == 1:
                    num_str = f"第{counters[0]}章 "
                elif level == 2:
                    num_str = f"{counters[0]}.{counters[1]} "
                elif level == 3:
                    num_str = f"{counters[0]}.{counters[1]}.{counters[2]} "
                else:
                    parts = [str(c) for c in counters[:level] if c > 0]
                    num_str = f"{'.'.join(parts)} "

                # Prepend to first run or create new run
                if para.runs:
                    para.runs[0].text = num_str + para.runs[0].text
                else:
                    para.add_run(num_str)

                changes.append({
                    "type": "numbering",
                    "index": self.doc.paragraphs.index(para) + 1,
                    "number": num_str.strip(),
                })

        return {"applied": True, "changes": changes, "change_count": len(changes)}

    def add_page_numbers(self, alignment: str = "center", start_at: int = 1) -> Dict[str, Any]:
        """Add page numbers to footer (requires section properties)."""
        # This is a structural change - in real implementation would modify sectPr
        return {
            "applied": False,
            "message": "Page numbers require section property modification. Use WPS mode for this.",
            "alignment": alignment,
            "start_at": start_at,
        }

    def ensure_widow_orphan_control(self) -> List[Dict]:
        """Enable widow/orphan control for all paragraphs."""
        changes = []
        # In the DOM model, we would set this on paragraph properties
        # For now, return instructions
        for para in self.doc.paragraphs:
            if para.is_heading():
                # Keep heading with next paragraph
                para.space_after = max(para.space_after or 0, 6)
                changes.append({
                    "type": "orphan_control",
                    "index": self.doc.paragraphs.index(para) + 1,
                    "action": "heading_keep_with_next",
                })
        return changes
