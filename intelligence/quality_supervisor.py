# -*- coding: utf-8 -*-
"""Quality Supervisor: evaluates and lightly fixes document formatting issues."""
from typing import Dict, List, Optional
from wps_bridge.app import get_doc, get_app
from wps_bridge.content import outline, paragraph
from wps_bridge.utils import com_property, com_set


def evaluate(doc_index: Optional[int] = None) -> Dict:
    """Evaluate document quality. Returns score (0-100) and issues."""
    doc = get_doc(doc_index)
    issues = []
    fixed = 0
    score = 100

    # 1. Check paragraph count is reasonable
    para_count = doc.Paragraphs.Count
    if para_count <= 1:
        issues.append({"issue": "文档段落过少", "fixed": False, "suggestion": "文档应有标题、正文等多个段落"})
        score -= 15

    # 2. Check content is not all in one paragraph
    all_in_one = _check_all_content_in_one_para(doc)
    if all_in_one:
        issues.append({"issue": "所有内容挤在一个段落", "fixed": False, "suggestion": "每条内容应独立成段"})
        score -= 20

    # 3. Check cover page quality
    cover_issues = _check_cover(doc)
    issues.extend(cover_issues)
    score -= len(cover_issues) * 5

    # 4. Check table quality
    table_issues = _check_tables(doc)
    issues.extend(table_issues)
    score -= len(table_issues) * 5

    # 5. Check content ordering
    order_issues = _check_ordering(doc)
    issues.extend(order_issues)
    score -= len(order_issues) * 5

    fixed = sum(1 for i in issues if i.get("fixed"))
    score = max(0, score)

    verdict = "excellent" if score >= 90 else "good" if score >= 70 else "needs_work" if score >= 50 else "poor"
    return {
        "score": score,
        "issues_found": len(issues),
        "issues_fixed": fixed,
        "issues": issues[:20],
        "verdict": verdict,
        "suggestion": _get_suggestion(verdict),
    }


def _get_suggestion(verdict: str) -> str:
    if verdict == "excellent":
        return "文档格式良好，无需修改"
    elif verdict == "good":
        return "有小问题已自动修复"
    elif verdict == "needs_work":
        return "建议运行 ai_format.reformat 优化排版"
    else:
        return "建议先用 ai_format.apply_template 套用模板，再用 ai_format.reformat 调整"


def _check_all_content_in_one_para(doc) -> bool:
    text_count = 0
    for i in range(1, doc.Paragraphs.Count + 1):
        try:
            if com_property(doc.Paragraphs.Item(i).Range, "Text", "").strip():
                text_count += 1
        except Exception:
            pass
    return text_count == 1 and doc.Paragraphs.Count <= 2


def _check_cover(doc) -> List[Dict]:
    issues = []
    try:
        para_count = doc.Paragraphs.Count
        if para_count < 2:
            return issues

        first_text_idx = 0
        for i in range(1, min(10, para_count + 1)):
            text = com_property(doc.Paragraphs.Item(i).Range, "Text", "").strip()
            if text and first_text_idx == 0:
                first_text_idx = i
                break

        if first_text_idx == 0:
            return issues

        p1 = doc.Paragraphs.Item(first_text_idx)
        f1 = p1.Range.Font
        size = com_property(f1, "Size", 0)
        align = com_property(p1.Format, "Alignment", 0)
        space_after = com_property(p1.Format, "SpaceAfter", 0)
        space_before = com_property(p1.Format, "SpaceBefore", 0)

        # Only fix if truly broken (title should be prominent)
        if size > 0 and size < 16:
            com_set(f1, "Size", 22)
            com_set(f1, "Bold", True)
            com_set(f1, "NameFarEast", "黑体")
            issues.append({"issue": "标题字体过小", "fixed": True, "action": "设为22pt黑体"})
        if align != 1 and first_text_idx == 1:
            com_set(p1.Format, "Alignment", 1)
            issues.append({"issue": "标题未居中", "fixed": True, "action": "居中"})
        if space_after < 12:
            com_set(p1.Format, "SpaceAfter", 24)
            issues.append({"issue": "标题与正文无间距", "fixed": True, "action": "增加段后间距"})
        if space_before < 36:
            com_set(p1.Format, "SpaceBefore", 72)
            issues.append({"issue": "标题距顶部过近", "fixed": True, "action": "增加段前间距"})
        com_set(f1, "ColorIndex", 1)

    except Exception as e:
        issues.append({"issue": f"封面检查异常: {e}", "fixed": False})
    return issues


def _check_tables(doc) -> List[Dict]:
    issues = []
    try:
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables.Item(ti)
            cols = tbl.Columns.Count
            rows = tbl.Rows.Count
            tbl_width = com_property(tbl, "PreferredWidth", 0)

            if tbl_width > 470 or tbl_width == 0:
                try:
                    tbl.AutoFitBehavior(2)
                    tbl.PreferredWidthType = 2
                    tbl.PreferredWidth = 451
                    issues.append({"issue": f"表格{ti}宽度调整", "fixed": True, "action": "设为451pt"})
                except Exception:
                    pass

            # Header formatting
            if rows > 0:
                try:
                    for c in range(1, cols + 1):
                        cell = tbl.Cell(1, c)
                        cell.Range.Font.Bold = True
                        cell.Range.Font.NameFarEast = "黑体"
                        cell.Range.Font.ColorIndex = 1
                        cell.Shading.BackgroundPatternColor = 0xE8E8E8
                except Exception:
                    pass

            # Row height warning
            if rows > 40:
                issues.append({"issue": f"表格{ti}行数({rows})过多", "fixed": False, "suggestion": "考虑拆分表格"})

    except Exception as e:
        issues.append({"issue": f"表格检查异常: {e}", "fixed": False})
    return issues


def _check_ordering(doc) -> List[Dict]:
    issues = []
    try:
        if doc.Tables.Count == 0:
            return issues
        first_tbl_start = com_property(doc.Tables.Item(1).Range, "Start", 0)
        if first_tbl_start is None or first_tbl_start == 0:
            return issues

        # Check if any heading (outline_level 1-9) appears after the first table
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs.Item(i)
            level = com_property(p.Format, "OutlineLevel", 10)
            text = com_property(p.Range, "Text", "").strip()
            p_start = com_property(p.Range, "Start", 0)
            if 1 <= level <= 9 and p_start > first_tbl_start and text:
                issues.append({"issue": f"标题出现在表格之后", "fixed": False, "suggestion": "正文和标题应在表格之前"})
                break

    except Exception as e:
        issues.append({"issue": f"顺序检查异常: {e}", "fixed": False})
    return issues


def sanitize_and_fix(doc_index: Optional[int] = None) -> Dict:
    """Evaluate + auto-fix, then re-evaluate."""
    result = evaluate(doc_index)
    if result["score"] < 50:
        # Force all text to have black color
        doc = get_doc(doc_index)
        for i in range(1, doc.Paragraphs.Count + 1):
            try:
                com_set(doc.Paragraphs.Item(i).Range.Font, "ColorIndex", 1)
            except Exception:
                pass
        result = evaluate(doc_index)
    return result
