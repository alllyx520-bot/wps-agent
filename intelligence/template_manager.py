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
from wps_bridge.content import outline, paragraph
from wps_bridge.document import doc_info
from wps_bridge.layout import section_info
from wps_bridge.table import table_info, table_read
from wps_bridge.formatting import get_font, get_paragraph_format

TEMPLATES_DIR = Path(__file__).parent / "templates"


def _ensure_dir():
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)


def extract(doc_index: Optional[int] = None) -> Dict:
    """Extract all formatting rules from a document into a template JSON."""
    info = doc_info(doc_index)
    outline_data = outline(doc_index)
    sections = section_info(doc_index)

    rules = {}
    page_rules = {}
    try:
        if "top_margin" in sections:
            page_rules = {
                "paper": "A4",
                "top_margin": sections.get("top_margin", 25.4),
                "bottom_margin": sections.get("bottom_margin", 25.4),
                "left_margin": sections.get("left_margin", 31.7),
                "right_margin": sections.get("right_margin", 31.7),
            }
    except Exception:
        page_rules = {"paper": "A4"}

    # Group headings by outline level
    level_groups = {}
    for h in outline_data:
        level = h["outline_level"]
        if level not in level_groups:
            level_groups[level] = []
        level_groups[level].append(h)

    for level, headings in level_groups.items():
        sample = paragraph(headings[0]["index"], doc_index)
        font = sample.get("font", {})
        pf = sample.get("paragraph_format", {})
        label = f"OutlineLevel {level}"
        rules[label] = {
            "font_name": font.get("name", "宋体"),
            "font_size": font.get("size", 12),
            "bold": font.get("bold", False),
            "alignment": pf.get("alignment", "left"),
            "outline_level": level,
            "space_before": pf.get("space_before", 0),
            "space_after": pf.get("space_after", 0),
            "line_spacing_rule": pf.get("line_spacing_rule", "single"),
            "line_spacing": pf.get("line_spacing", 1),
            "first_line_indent": pf.get("first_line_indent", 0),
        }

    # Extract body text formatting
    try:
        from wps_bridge.app import get_doc
        doc = get_doc(doc_index)
        body_sample = None
        for i in range(1, min(doc.Paragraphs.Count + 1, 50)):
            p = doc.Paragraphs.Item(i)
            ol = None
            try:
                ol = p.Format.OutlineLevel
            except Exception:
                pass
            if ol is None or ol > 9:
                body_sample = paragraph(i, doc_index)
                break
        if body_sample:
            font = body_sample.get("font", {})
            pf = body_sample.get("paragraph_format", {})
            rules["正文"] = {
                "font_name": font.get("name", "宋体"),
                "font_size": font.get("size", 12),
                "bold": font.get("bold", False),
                "alignment": pf.get("alignment", "justify"),
                "first_line_indent": pf.get("first_line_indent", 28),
                "line_spacing_rule": pf.get("line_spacing_rule", "multiple"),
                "line_spacing": pf.get("line_spacing", 1.5),
            }
    except Exception:
        pass

    rules["page"] = page_rules

    return {
        "name": f"Extracted from {info.get('name', 'unknown')}",
        "extracted_from": info.get("full_name", info.get("name", "")),
        "extracted_at": None,  # filled by caller
        "doc_type": "general",
        "description": f"Auto-extracted from {info.get('name', 'unknown')}",
        "rules": rules,
    }


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
    """Compare current document formatting against a template."""
    doc_rules = extract(doc_index).get("rules", {})

    cmp_rules = None
    saved = load(template_name)
    if saved:
        cmp_rules = saved.get("rules", {})
    elif template_name in CHINESE_FORMATTING:
        cmp_rules = CHINESE_FORMATTING[template_name]
    else:
        return {"error": f"Template not found: {template_name}", "available_builtins": list(CHINESE_FORMATTING.keys())}

    mismatches = []
    for key, expected in cmp_rules.items():
        if key == "page":
            continue
        actual = doc_rules.get(key, {})
        if not actual:
            mismatches.append(f"Missing rule: {key}")
            continue
        for prop in ["font_name", "font_size", "bold", "alignment", "first_line_indent"]:
            if prop in expected and prop in actual:
                if expected[prop] != actual[prop]:
                    mismatches.append(f"{key}.{prop}: expected {expected[prop]}, actual {actual[prop]}")

    return {
        "template": template_name,
        "doc_rules_count": len(doc_rules),
        "template_rules_count": len(cmp_rules),
        "mismatches": len(mismatches),
        "details": mismatches[:50],
    }
