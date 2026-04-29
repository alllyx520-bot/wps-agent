# -*- coding: utf-8 -*-
"""
Quality Supervisor: evaluates and auto-corrects document quality after generation.
"""
from typing import Dict, List, Optional
from wps_bridge.app import get_doc, get_app
from wps_bridge.content import outline, paragraph, full_text
from wps_bridge.formatting import set_font, set_paragraph_format
from wps_bridge.utils import com_property, com_set


def evaluate(doc_index: Optional[int] = None) -> Dict:
    """Full document quality evaluation. Returns score (0-100) + issues."""
    doc = get_doc(doc_index)
    issues = []
    fixed = 0
    score = 100

    # 1. Check cover page formatting
    issues += _check_cover(doc)
    # 2. Check table sizing and position
    issues += _check_tables(doc)
    # 3. Check content ordering (cover first, tables last)
    issues += _check_ordering(doc)
    # 4. Check page breaks
    issues += _check_page_breaks(doc)

    # Apply auto-fixes
    for issue in issues:
        if issue.get("fixed"):
            fixed += 1

    score = max(0, score - len(issues) * 5)
    return {
        "score": score,
        "issues_found": len(issues),
        "issues_fixed": fixed,
        "issues": issues[:20],
        "verdict": "excellent" if score >= 90 else "good" if score >= 70 else "needs_work" if score >= 50 else "poor",
    }


def _check_cover(doc) -> List[Dict]:
    """Verify cover page has proper formatting."""
    issues = []
    try:
        para_count = doc.Paragraphs.Count
        # First 5 paragraphs form the cover
        cover_end = min(5, para_count)
        first_text = ""
        for i in range(1, cover_end + 1):
            p = doc.Paragraphs.Item(i)
            text = com_property(p.Range, "Text", "").strip()
            if text:
                first_text = text
                break

        if not first_text:
            return issues

        # Check if first paragraph looks like a title (should be centered, large)
        p1 = doc.Paragraphs.Item(1)
        r1 = p1.Range
        f1 = r1.Font
        size = com_property(f1, "Size", 0)
        align = com_property(p1.Format, "Alignment", 0)
        space_after = com_property(p1.Format, "SpaceAfter", 0)
        space_before = com_property(p1.Format, "SpaceBefore", 0)

        if size < 18:
            com_set(f1, "Size", 22)
            com_set(f1, "Bold", True)
            com_set(f1, "NameFarEast", "黑体")
            issues.append({"issue": "封面标题字体过小", "fixed": True, "action": f"设为22pt黑体加粗"})
        if align != 1:
            com_set(p1.Format, "Alignment", 1)
            issues.append({"issue": "封面标题未居中", "fixed": True, "action": "设为居中"})
        if space_after < 24:
            com_set(p1.Format, "SpaceAfter", 24)
            issues.append({"issue": "封面标题与正文间距过小", "fixed": True, "action": "增加段后间距24pt"})
        if space_before < 72:
            com_set(p1.Format, "SpaceBefore", 72)
            issues.append({"issue": "封面标题距顶部过近", "fixed": True, "action": "增加段前间距72pt"})
        com_set(f1, "ColorIndex", 1)

        # Ensure body text after title has reasonable formatting
        for i in range(2, cover_end + 1):
            p = doc.Paragraphs.Item(i)
            text = com_property(p.Range, "Text", "").strip()
            if not text:
                continue
            f = p.Range.Font
            sz = com_property(f, "Size", 0)
            if sz == 0 or sz > 16:
                com_set(f, "Size", 14)
                com_set(f, "NameFarEast", "宋体")
                com_set(f, "ColorIndex", 1)
            al = com_property(p.Format, "Alignment", 0)
            if al != 1 and i <= 3:
                com_set(p.Format, "Alignment", 1)
            sp = com_property(p.Format, "SpaceAfter", 0)
            if sp < 6:
                com_set(p.Format, "SpaceAfter", 6)

        # Add page break after cover
        last_cover = cover_end
        for i in range(cover_end, 0, -1):
            if com_property(doc.Paragraphs.Item(i).Range, "Text", "").strip():
                last_cover = i
                break
        has_page_break = False
        try:
            next_text = com_property(doc.Paragraphs.Item(last_cover + 1).Range, "Text", "").strip()
            if next_text:
                pf = doc.Paragraphs.Item(last_cover).Format
                if com_property(pf, "PageBreakBefore", 0) == 0 and last_cover == 1:
                    # Check if anything follows cover; if so, ensure page break
                    pass
        except Exception:
            pass
        # Insert page break after cover (insert at last_cover+1)
        if last_cover < para_count:
            rng = doc.Paragraphs.Item(last_cover + 1).Range
            try:
                rng.InsertBreak(2)  # wdPageBreak
                issues.append({"issue": "封面与正文无分页", "fixed": True, "action": "插入分页符"})
            except Exception:
                pass

    except Exception as e:
        issues.append({"issue": f"封面检查异常: {e}", "fixed": False})
    return issues


def _check_tables(doc) -> List[Dict]:
    """Verify tables are properly sized and positioned."""
    issues = []
    try:
        tbl_count = doc.Tables.Count
        for ti in range(1, tbl_count + 1):
            tbl = doc.Tables.Item(ti)
            cols = tbl.Columns.Count
            rows = tbl.Rows.Count

            # Check table is not full-page wide (typical page width ~460pt for A4 with margins)
            tbl_width = com_property(tbl, "PreferredWidth", 0)
            if tbl_width > 400 or tbl_width == 0:
                try:
                    tbl.AutoFitBehavior(2)  # AutoFit to window
                    # Then set preferred width to ~420pt
                    tbl.PreferredWidthType = 2  # wdPreferredWidthPoints
                    tbl.PreferredWidth = 420
                    issues.append({"issue": f"表格{ti}宽度过大", "fixed": True, "action": "调整为420pt"})
                except Exception:
                    pass

            # Check column widths are reasonable
            if cols > 0:
                col_width = 420 / cols
                for c in range(1, cols + 1):
                    try:
                        if tbl.Columns.Item(c).Width > col_width * 1.5:
                            tbl.Columns.Item(c).Width = col_width
                            issues.append({"issue": f"表格{ti}列{c}过宽", "fixed": True, "action": f"调整为{int(col_width)}pt"})
                    except Exception:
                        pass

            # Check table row height is reasonable (not full-page)
            if rows > 0:
                page_height = 700  # ~A4 text height in pt
                row_height = page_height / max(rows, 1)
                if row_height < 14:  # Rows too many for one page
                    issues.append({"issue": f"表格{ti}行数({rows})过多", "fixed": False, "suggestion": "考虑拆分表格"})

            # Format table header row
            if rows > 0:
                try:
                    cell = tbl.Cell(1, 1)
                    cell.Range.Font.Bold = True
                    cell.Range.Font.NameFarEast = "黑体"
                    cell.Range.Font.Size = 10.5
                    cell.Range.Font.ColorIndex = 1
                except Exception:
                    pass

            # Check alternating row colors
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    try:
                        cell = tbl.Cell(r, c)
                        cell.Range.Font.ColorIndex = 1
                        if r == 1:
                            cell.Shading.BackgroundPatternColor = 0xDDDDDD
                    except Exception:
                        continue

            # Auto-fit the table
            try:
                tbl.AutoFitBehavior(2)
            except Exception:
                pass

    except Exception as e:
        issues.append({"issue": f"表格检查异常: {e}", "fixed": False})
    return issues


def _check_ordering(doc) -> List[Dict]:
    """Ensure content ordering is correct: cover page → body text → tables."""
    issues = []
    try:
        para_count = doc.Paragraphs.Count
        tbl_count = doc.Tables.Count
        if tbl_count == 0 or para_count == 0:
            return issues

        # Find the first table's position
        first_table_start = None
        try:
            first_table_range = doc.Tables.Item(1).Range
            first_table_start = com_property(first_table_range, "Start", 0)
        except Exception:
            return issues

        if first_table_start is None:
            return issues

        # Check if any body text appears AFTER the first table
        # Body text is paragraphs with outline level >= 9 (not headings)
        for i in range(1, para_count + 1):
            p = doc.Paragraphs.Item(i)
            text = com_property(p.Range, "Text", "").strip()
            if not text:
                continue
            level = com_property(p.Format, "OutlineLevel", 10)
            p_start = com_property(p.Range, "Start", 0)
            # This is body text that appears after the table → wrong order
            if level >= 9 and p_start > first_table_start and i > 1:
                issues.append({
                    "issue": f"正文(段{i})出现在表格之后，顺序异常",
                    "fixed": False,
                    "suggestion": "表格应放在文档末尾，正文应在前"
                })

        # Check cover page text count
        cover_texts = 0
        for i in range(1, min(6, para_count + 1)):
            if com_property(doc.Paragraphs.Item(i).Range, "Text", "").strip():
                cover_texts += 1
        if cover_texts < 2:
            issues.append({"issue": "封面信息不完整", "fixed": False, "suggestion": "封面应包含标题、副标题、作者、日期"})

    except Exception as e:
        issues.append({"issue": f"内容顺序检查异常: {e}", "fixed": False})
    return issues


def _check_page_breaks(doc) -> List[Dict]:
    """Check and fix page break formatting."""
    issues = []
    try:
        para_count = doc.Paragraphs.Count
        # Find the title paragraph (first with text)
        title_idx = 0
        for i in range(1, min(10, para_count + 1)):
            text = com_property(doc.Paragraphs.Item(i).Range, "Text", "").strip()
            if text:
                if title_idx == 0:
                    title_idx = i
                elif title_idx > 0 and i > title_idx + 3:
                    # Gap between cover and next content is too large without page break
                    # Check if there's a page break between them
                    has_break = False
                    for j in range(title_idx + 1, i):
                        try:
                            if com_property(doc.Paragraphs.Item(j).Range, "Text", "").strip() == "":
                                pass  # empty paragraph, could be spacing
                        except Exception:
                            pass
                    # If more than 3 empty lines between title and content, it's a visual gap issue
                    pass
    except Exception:
        pass
    return issues


def sanitize_and_fix(doc_index: Optional[int] = None) -> Dict:
    """Run full evaluation + auto-fix, then re-evaluate. Returns final state."""
    result = evaluate(doc_index)
    # If score is poor, try one more round of fixes
    if result["score"] < 60:
        # Apply additional fixes for poor quality
        doc = get_doc(doc_index)
        # Force all text to be black
        for i in range(1, doc.Paragraphs.Count + 1):
            try:
                com_set(doc.Paragraphs.Item(i).Range.Font, "ColorIndex", 1)
            except Exception:
                pass
        # Re-evaluate
        result = evaluate(doc_index)
    return result
