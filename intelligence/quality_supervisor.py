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

    # 6. Check widow/orphan headings
    widow_issues = _check_widow_orphan(doc)
    issues.extend(widow_issues)
    score -= len(widow_issues) * 5

    # 7. Check paragraph spacing consistency
    spacing_issues = _check_spacing_consistency(doc)
    issues.extend(spacing_issues)
    score -= len(spacing_issues) * 3

    # 8. Check font consistency
    font_issues = _check_font_consistency(doc)
    issues.extend(font_issues)
    score -= len(font_issues) * 3

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


def _check_widow_orphan(doc) -> List[Dict]:
    """Detect widows (lonely last line) and orphans (lonely heading at bottom).
    Works in COM mode by checking heading proximity to subsequent content."""
    issues = []
    try:
        para_count = doc.Paragraphs.Count
        if para_count < 3:
            return issues

        for i in range(1, para_count):
            p = doc.Paragraphs.Item(i)
            level = com_property(p.Format, "OutlineLevel", 10)
            text = com_property(p.Range, "Text", "").strip()
            is_heading = 1 <= level <= 9

            if is_heading:
                # Check if heading is followed by very little content
                following_text = ""
                for j in range(i + 1, min(i + 5, para_count + 1)):
                    try:
                        next_level = com_property(doc.Paragraphs.Item(j).Format, "OutlineLevel", 10)
                        if 1 <= next_level <= level:
                            break
                        t = com_property(doc.Paragraphs.Item(j).Range, "Text", "").strip()
                        following_text += t
                    except Exception:
                        break

                # Heading with no body content → orphan
                if len(following_text) < 10:
                    issues.append({
                        "issue": f"标题后有实质性内容",
                        "fixed": False,
                        "suggestion": f"第{i}段'{text[:30]}'后正文内容过少({len(following_text)}字)，可能为孤行标题",
                    })

            # Check single-line body paragraphs at position near "page boundaries"
            if not is_heading and len(text) < 25 and i > 1 and i < para_count - 1:
                prev_empty = not com_property(doc.Paragraphs.Item(i - 1).Range, "Text", "").strip()
                next_empty = not com_property(doc.Paragraphs.Item(i + 1).Range, "Text", "").strip()
                if prev_empty and next_empty:
                    issues.append({
                        "issue": f"孤立短段落在两段空行之间",
                        "fixed": False,
                        "suggestion": f"第{i}段是孤立短段落，考虑删除或合并",
                    })

        # Check table cross-page break risk
        for ti in range(1, doc.Tables.Count + 1):
            tbl = doc.Tables.Item(ti)
            rows = tbl.Rows.Count
            if rows > 30:
                issues.append({
                    "issue": f"表格{ti}行数({rows})较多可能跨页断裂",
                    "fixed": False,
                    "suggestion": "设置表头重复或拆分表格",
                })

    except Exception as e:
        issues.append({"issue": f"孤行检查异常: {e}", "fixed": False})
    return issues


def _check_spacing_consistency(doc) -> List[Dict]:
    """Check paragraph spacing consistency across the document."""
    issues = []
    try:
        spaces_before = []
        spaces_after = []
        for i in range(1, min(doc.Paragraphs.Count + 1, 200)):
            try:
                pf = doc.Paragraphs.Item(i).Format
                sb = com_property(pf, "SpaceBefore", 0)
                sa = com_property(pf, "SpaceAfter", 0)
                text = com_property(doc.Paragraphs.Item(i).Range, "Text", "").strip()
                if text:
                    spaces_before.append(sb)
                    spaces_after.append(sa)
            except Exception:
                pass

        if len(spaces_before) > 5:
            from statistics import mean, stdev
            try:
                avg_sb = mean(spaces_before)
                std_sb = stdev(spaces_before) if len(spaces_before) > 1 else 0
                if std_sb > 12:
                    issues.append({
                        "issue": f"段前间距不一致(标准差={std_sb:.1f}pt)",
                        "fixed": False,
                        "suggestion": "统一段前间距",
                    })
            except Exception:
                pass

    except Exception as e:
        issues.append({"issue": f"间距检查异常: {e}", "fixed": False})
    return issues


def _check_font_consistency(doc) -> List[Dict]:
    """Check font consistency across body text paragraphs."""
    issues = []
    try:
        fonts = {}
        for i in range(1, min(doc.Paragraphs.Count + 1, 200)):
            try:
                p = doc.Paragraphs.Item(i)
                level = com_property(p.Format, "OutlineLevel", 10)
                if 1 <= level <= 9:
                    continue
                f = p.Range.Font
                name = com_property(f, "NameFarEast", "") or com_property(f, "Name", "")
                size = com_property(f, "Size", 0)
                if name and size > 0:
                    key = (name, size)
                    fonts[key] = fonts.get(key, 0) + 1
            except Exception:
                pass

        if len(fonts) > 3:
            issues.append({
                "issue": f"正文字体不统一(发现{len(fonts)}种字体/字号组合)",
                "fixed": False,
                "suggestion": "全选正文统一为宋体小四",
            })

    except Exception as e:
        issues.append({"issue": f"字体检查异常: {e}", "fixed": False})
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
