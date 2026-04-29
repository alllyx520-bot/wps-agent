# -*- coding: utf-8 -*-
"""
Template Manager: extract/save/load/compare document formatting templates.
Stores templates as JSON in intelligence/templates/ directory.
"""
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from .chinese_rules import CHINESE_FORMATTING
from wps_bridge.document import doc_info
from wps_bridge.layout import section_info
from wps_bridge.table import table_info
from wps_bridge.formatting import get_font, get_paragraph_format

TEMPLATES_DIR = Path(__file__).parent / "templates"


def _ensure_dir():
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)


def _extract_full_paragraph(doc, i: int, doc_index: Optional[int]) -> Dict:
    """Extract every available font + paragraph property for a single paragraph."""
    from wps_bridge.utils import com_property
    from wps_bridge import formatting as fmt
    from wps_bridge import content as cnt

    font = fmt.get_font(para_index=i, doc_index=doc_index)
    pf = fmt.get_paragraph_format(i, doc_index=doc_index)
    text = ""
    style_name = ""
    try:
        p = doc.Paragraphs.Item(i)
        r = p.Range
        text = com_property(r, "Text", "")
        style_name = com_property(r.Style, "NameLocal", "")
    except Exception:
        pass

    return {
        "index": i,
        "text": text[:120],
        "style_name": style_name,
        "font": {
            "name": font.get("name", ""),
            "name_far_east": font.get("name_far_east", ""),
            "size": font.get("size", 0),
            "bold": font.get("bold", False),
            "italic": font.get("italic", False),
            "underline": font.get("underline", 0),
            "color_index": font.get("color_index", 1),
            "superscript": font.get("superscript", False),
            "subscript": font.get("subscript", False),
            "strike_through": font.get("strike_through", False),
            "spacing": font.get("spacing", 0),
            "scaling": font.get("scaling", 100),
            "kerning": font.get("kerning", 0),
        },
        "paragraph_format": {
            "alignment": pf.get("alignment", "left"),
            "alignment_raw": pf.get("alignment_raw", 0),
            "first_line_indent": pf.get("first_line_indent", 0),
            "left_indent": pf.get("left_indent", 0),
            "right_indent": pf.get("right_indent", 0),
            "line_spacing_rule": pf.get("line_spacing_rule", "single"),
            "line_spacing_rule_raw": pf.get("line_spacing_rule_raw", 0),
            "line_spacing": pf.get("line_spacing", 0),
            "space_before": pf.get("space_before", 0),
            "space_after": pf.get("space_after", 0),
            "outline_level": pf.get("outline_level", 10),
            "widow_control": pf.get("widow_control", False),
            "keep_with_next": pf.get("keep_with_next", False),
        },
    }


def _make_fingerprint(pdata: Dict, ignore_text: bool = True) -> str:
    """Build a hash from font + paragraph format to identify identical formatting."""
    import hashlib
    parts = []
    f = pdata.get("font", {})
    pf = pdata.get("paragraph_format", {})
    parts.append(str(f.get("name", "")))
    parts.append(str(f.get("name_far_east", "")))
    parts.append(str(f.get("size", 0)))
    parts.append(str(int(f.get("bold", False))))
    parts.append(str(int(f.get("italic", False))))
    parts.append(str(f.get("underline", 0)))
    parts.append(str(f.get("color_index", 1)))
    parts.append(str(int(f.get("superscript", False))))
    parts.append(str(int(f.get("subscript", False))))
    parts.append(str(int(f.get("strike_through", False))))
    parts.append(str(f.get("spacing", 0)))
    parts.append(str(f.get("scaling", 100)))
    parts.append(str(f.get("kerning", 0)))
    parts.append(str(pf.get("alignment_raw", 0)))
    parts.append(str(pf.get("first_line_indent", 0)))
    parts.append(str(pf.get("left_indent", 0)))
    parts.append(str(pf.get("right_indent", 0)))
    parts.append(str(pf.get("line_spacing_rule_raw", 0)))
    parts.append(str(pf.get("line_spacing", 0)))
    parts.append(str(pf.get("space_before", 0)))
    parts.append(str(pf.get("space_after", 0)))
    parts.append(str(pf.get("outline_level", 10)))
    parts.append(str(int(pf.get("widow_control", False))))
    parts.append(str(int(pf.get("keep_with_next", False))))
    parts.append(pdata.get("style_name", ""))
    return hashlib.md5("|".join(parts).encode()).hexdigest()[:8]


def _extract_section_details(doc, doc_index: Optional[int]) -> Dict:
    """Get detailed section info including headers/footers, page setup, columns."""
    from wps_bridge.utils import com_property
    from wps_bridge import layout as lay

    sections_data = lay.section_info(doc_index=doc_index)
    result = {
        "page_width": sections_data.get("page_width", 0),
        "page_height": sections_data.get("page_height", 0),
        "top_margin": sections_data.get("top_margin", 0),
        "bottom_margin": sections_data.get("bottom_margin", 0),
        "left_margin": sections_data.get("left_margin", 0),
        "right_margin": sections_data.get("right_margin", 0),
        "orientation": "portrait" if sections_data.get("orientation", 0) == 0 else "landscape",
        "paper": "A4" if abs(sections_data.get("page_width", 595) - 595) < 5 else "custom",
    }

    # Try to detect headers/footers
    try:
        sec = doc.Sections.Item(1)
        hdr = sec.Headers
        ftr = sec.Footers
        has_header = False
        has_footer = False
        try:
            has_header = len(com_property(hdr.Item(1).Range, "Text", "").strip()) > 0
        except Exception:
            pass
        try:
            has_footer = len(com_property(ftr.Item(1).Range, "Text", "").strip()) > 0
        except Exception:
            pass
        result["has_header"] = has_header
        result["has_footer"] = has_footer
    except Exception:
        pass

    return result


def extract(doc_index: Optional[int] = None) -> Dict:
    """Extract every formatting detail from a document into a comprehensive template JSON.

    Scans ALL paragraphs for full font (13 props) + paragraph (13 props) formatting.
    Returns distinct patterns by fingerprint, outline-level rules, cover detection,
    style usage, and section details.
    """
    from wps_bridge.app import get_doc
    from wps_bridge.utils import com_property

    info = doc_info(doc_index)
    doc = get_doc(doc_index)
    para_count = doc.Paragraphs.Count

    # ── 1. Page / Section Setup ──
    section_detail = _extract_section_details(doc, doc_index)

    # ── 2. Full paragraph scan (every paragraph, every property) ──
    all_paragraphs = []
    for i in range(1, para_count + 1):
        try:
            pdata = _extract_full_paragraph(doc, i, doc_index)
            all_paragraphs.append(pdata)
        except Exception:
            continue

    if not all_paragraphs:
        return {"error": "No paragraphs found", "doc_name": info.get("name", "")}

    # ── 3. Format fingerprinting → distinct patterns ──
    pattern_groups = {}
    for pdata in all_paragraphs:
        fp = _make_fingerprint(pdata)
        if fp not in pattern_groups:
            pattern_groups[fp] = []
        pattern_groups[fp].append(pdata)

    distinct_patterns = []
    for fp, group in sorted(pattern_groups.items(), key=lambda x: -len(x[1])):
        sample = group[0]
        distinct_patterns.append({
            "fingerprint": fp,
            "count": len(group),
            "percentage": round(len(group) / len(all_paragraphs) * 100, 1),
            "font": sample["font"],
            "paragraph_format": sample["paragraph_format"],
            "style_name": sample["style_name"],
            "para_indices": [p["index"] for p in group[:20]],
            "examples": [p["text"] for p in group[:5] if p["text"].strip()],
        })

    # ── 4. Outline-level rules (backward compatible) ──
    level_groups = {}
    for pdata in all_paragraphs:
        ol = pdata["paragraph_format"]["outline_level"]
        if 1 <= ol <= 9:
            if ol not in level_groups:
                level_groups[ol] = []
            level_groups[ol].append(pdata)

    outline_rules = {}
    for level in sorted(level_groups):
        sample = level_groups[level][0]
        font = sample["font"]
        pf = sample["paragraph_format"]
        outline_rules[f"heading_level_{level}"] = {
            "count": len(level_groups[level]),
            "font": dict(font),
            "paragraph_format": dict(pf),
            "style_name": sample["style_name"],
        }

    # ── 5. Body text rules (first body paragraph pattern) ──
    body_rules = {}
    body_found = None
    for pdata in all_paragraphs:
        if pdata["paragraph_format"]["outline_level"] > 9 and pdata["text"].strip():
            body_found = pdata
            break
    if body_found:
        body_rules = {
            "font": dict(body_found["font"]),
            "paragraph_format": dict(body_found["paragraph_format"]),
            "style_name": body_found["style_name"],
            "indices": list(range(body_found["index"], min(body_found["index"] + 5, para_count + 1))),
        }

    # ── 6. Cover detection: first non-empty paragraphs before first heading ──
    cover_paragraphs = []
    for pdata in all_paragraphs:
        if 1 <= pdata["paragraph_format"]["outline_level"] <= 9:
            break
        if pdata["text"].strip():
            cover_paragraphs.append({
                "index": pdata["index"],
                "font": dict(pdata["font"]),
                "paragraph_format": dict(pdata["paragraph_format"]),
                "style_name": pdata["style_name"],
                "text": pdata["text"],
            })
            if len(cover_paragraphs) >= 10:
                break

    is_cover_detected = bool(cover_paragraphs) and all(
        cp["paragraph_format"]["alignment"] == "center"
        for cp in cover_paragraphs[:1]  # first line centered suggests a cover
    )

    # ── 7. Named style usage ──
    style_counter = {}
    for pdata in all_paragraphs:
        sn = pdata["style_name"]
        if sn:
            style_counter[sn] = style_counter.get(sn, 0) + 1
    styles_used = [
        {"name": name, "count": count, "pct": round(count / len(all_paragraphs) * 100, 1)}
        for name, count in sorted(style_counter.items(), key=lambda x: -x[1])
    ]

    # ── 8. Table summary ──
    table_count = doc.Tables.Count
    tables_summary = []
    for ti in range(1, min(table_count + 1, 20)):
        try:
            tinfo = table_info(ti, doc_index)
            tables_summary.append({
                "index": ti,
                "rows": tinfo.get("rows", 0),
                "cols": tinfo.get("cols", 0),
            })
        except Exception:
            continue

    # ── 9. Build result ──
    return {
        "name": f"Extracted from {info.get('name', 'unknown')}",
        "extracted_from": info.get("full_name", info.get("name", "")),
        "extracted_at": None,
        "doc_type": "general",
        "description": f"Auto-extracted from {info.get('name', 'unknown')}",
        "summary": {
            "total_paragraphs": para_count,
            "total_paragraphs_scanned": len(all_paragraphs),
            "distinct_patterns": len(distinct_patterns),
            "outline_levels": sorted(level_groups.keys()),
            "has_cover": is_cover_detected,
            "cover_paragraphs_count": len(cover_paragraphs),
            "styles_used": len(styles_used),
            "table_count": table_count,
        },
        "section": section_detail,
        "cover": cover_paragraphs,
        "outline_rules": outline_rules,
        "body_rules": body_rules,
        "distinct_patterns": distinct_patterns,
        "styles_used": styles_used,
        "tables": tables_summary,
        # backward compatible simplified rules
        "rules": _build_legacy_rules(section_detail, outline_rules, distinct_patterns, cover_paragraphs),
    }


def _build_legacy_rules(section_detail, outline_rules, distinct_patterns, cover_paragraphs) -> Dict:
    """Build the old-style 'rules' dict for backward compatibility with apply_template."""
    rules = {}
    rules["page"] = {
        "paper": section_detail.get("paper", "A4"),
        "top_margin": section_detail.get("top_margin", 25.4),
        "bottom_margin": section_detail.get("bottom_margin", 25.4),
        "left_margin": section_detail.get("left_margin", 31.7),
        "right_margin": section_detail.get("right_margin", 31.7),
    }

    for label, rule in outline_rules.items():
        font = rule.get("font", {})
        pf = rule.get("paragraph_format", {})
        rules[label] = {
            "font_name": font.get("name_far_east", font.get("name", "宋体")),
            "font_size": font.get("size", 12),
            "bold": font.get("bold", False),
            "italic": font.get("italic", False),
            "underline": font.get("underline", 0),
            "alignment": pf.get("alignment", "left"),
            "outline_level": pf.get("outline_level", 10),
            "space_before": pf.get("space_before", 0),
            "space_after": pf.get("space_after", 0),
            "line_spacing_rule": pf.get("line_spacing_rule", "single"),
            "line_spacing": pf.get("line_spacing", 1),
            "first_line_indent": pf.get("first_line_indent", 0),
            "left_indent": pf.get("left_indent", 0),
            "right_indent": pf.get("right_indent", 0),
        }

    # Find the most common non-heading pattern as body text
    body_pattern = None
    for pat in distinct_patterns:
        if pat["paragraph_format"]["outline_level"] > 9 and pat["font"].get("size", 0) > 0:
            body_pattern = pat
            break
    if body_pattern:
        font = body_pattern.get("font", {})
        pf = body_pattern.get("paragraph_format", {})
        rules["正文"] = {
            "font_name": font.get("name_far_east", font.get("name", "宋体")),
            "font_size": font.get("size", 12),
            "bold": font.get("bold", False),
            "italic": font.get("italic", False),
            "underline": font.get("underline", 0),
            "alignment": pf.get("alignment", "justify"),
            "first_line_indent": pf.get("first_line_indent", 28),
            "line_spacing_rule": pf.get("line_spacing_rule", "multiple"),
            "line_spacing": pf.get("line_spacing", 1.5),
            "space_before": pf.get("space_before", 0),
            "space_after": pf.get("space_after", 0),
        }

    if cover_paragraphs:
        cp = cover_paragraphs[0]
        font = cp.get("font", {})
        pf = cp.get("paragraph_format", {})
        rules["封面标题"] = {
            "font_name": font.get("name_far_east", font.get("name", "黑体")),
            "font_size": font.get("size", 22),
            "bold": font.get("bold", True),
            "alignment": pf.get("alignment", "center"),
            "space_before": pf.get("space_before", 0),
            "space_after": pf.get("space_after", 12),
            "is_cover": True,
        }

    return rules


def save(template_name: str, template_data: Dict) -> Dict:
    """Save a template as JSON file."""
    _ensure_dir()
    from datetime import datetime
    template_data["extracted_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    filepath = TEMPLATES_DIR / f"{template_name}.json"
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(template_data, f, ensure_ascii=False, indent=2)
    return {"saved": template_name, "path": str(filepath)}


def load(template_name: str) -> Optional[Dict]:
    """Load a saved template by name."""
    _ensure_dir()
    filepath = TEMPLATES_DIR / f"{template_name}.json"
    if filepath.exists():
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def list_all() -> List[Dict]:
    """List all saved custom templates."""
    _ensure_dir()
    result = []
    for f in sorted(TEMPLATES_DIR.glob("*.json")):
        try:
            with open(f, "r", encoding="utf-8") as fp:
                data = json.load(fp)
                result.append({
                    "name": f.stem,
                    "description": data.get("description", ""),
                    "extracted_from": data.get("extracted_from", ""),
                    "extracted_at": data.get("extracted_at", ""),
                })
        except Exception:
            continue
    # Also include built-in presets
    for name in CHINESE_FORMATTING:
        result.append({"name": name, "type": "builtin", "description": f"Built-in: {name}"})
    return result


def delete(template_name: str) -> Dict:
    """Delete a saved template."""
    filepath = TEMPLATES_DIR / f"{template_name}.json"
    if filepath.exists():
        filepath.unlink()
        return {"deleted": template_name}
    return {"error": f"Template not found: {template_name}"}


def export_template(template_name: str, filepath: str) -> Dict:
    """Export a template to an external JSON file."""
    data = load(template_name)
    if not data:
        # Try builtin
        if template_name in CHINESE_FORMATTING:
            data = {
                "name": template_name,
                "doc_type": template_name,
                "rules": CHINESE_FORMATTING[template_name],
            }
        else:
            return {"error": f"Template not found: {template_name}"}
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return {"exported": template_name, "path": filepath}


def import_template(filepath: str) -> Dict:
    """Import a template from a JSON file."""
    _ensure_dir()
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
    name = data.get("name", Path(filepath).stem)
    filepath_dest = TEMPLATES_DIR / f"{name}.json"
    with open(filepath_dest, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return {"imported": name, "path": str(filepath_dest)}


def compare_with_template(doc_index: Optional[int] = None, template_name: str = "") -> Dict:
    """Compare current document formatting against a template. Uses detailed extraction."""
    extracted = extract(doc_index)
    doc_rules = extracted.get("rules", {})
    doc_detailed = extracted.get("distinct_patterns", [])

    cmp_rules = None
    saved = load(template_name)
    if saved:
        cmp_rules = saved.get("rules", {})
    elif template_name in CHINESE_FORMATTING:
        cmp_rules = CHINESE_FORMATTING[template_name]
    else:
        return {"error": f"Template not found: {template_name}", "available_builtins": list(CHINESE_FORMATTING.keys())}

    mismatches = []
    compare_props = [
        ("font_name", "字体"),
        ("font_size", "字号"),
        ("bold", "加粗"),
        ("italic", "斜体"),
        ("underline", "下划线"),
        ("alignment", "对齐"),
        ("first_line_indent", "首行缩进"),
        ("left_indent", "左缩进"),
        ("right_indent", "右缩进"),
        ("line_spacing_rule", "行距规则"),
        ("line_spacing", "行距值"),
        ("space_before", "段前间距"),
        ("space_after", "段后间距"),
    ]

    for key, expected in cmp_rules.items():
        if key == "page":
            continue
        actual = doc_rules.get(key, {})
        if not actual:
            mismatches.append({"rule": key, "issue": "文档中未找到该规则"})
            continue
        for prop_en, prop_cn in compare_props:
            exp_val = expected.get(prop_en)
            act_val = actual.get(prop_en)
            if exp_val is not None and act_val is not None and exp_val != act_val:
                mismatches.append({
                    "rule": key,
                    "property": prop_en,
                    "property_cn": prop_cn,
                    "expected": exp_val,
                    "actual": act_val,
                })

    return {
        "template": template_name,
        "doc_rules_count": len(doc_rules),
        "template_rules_count": len(cmp_rules),
        "mismatches": len(mismatches),
        "details": mismatches[:100],
    }
