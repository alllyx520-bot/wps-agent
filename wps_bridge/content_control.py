# -*- coding: utf-8 -*-
"""Content Control bridge — Rich Text, Date Picker, Dropdown, Checkbox, Picture controls."""
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set, com_set_batch

# WPS ContentControlType constants
WDCONTENTCONTROL = {
    "RICH_TEXT": 0,
    "TEXT": 1,
    "PICTURE": 2,
    "COMBO_BOX": 3,
    "DROPDOWN_LIST": 4,
    "BUILDING_BLOCK_GALLERY": 5,
    "DATE_PICKER": 6,
    "CHECKBOX": 8,
    "GROUP": 9,
    "REPEATING_SECTION": 10,
}


def _get_cc(doc, index=None):
    """Get a content control by index (1-based) or return all."""
    if index is not None:
        return doc.ContentControls.Item(index)
    return doc.ContentControls


def _safe_placeholder(cc) -> str:
    """Extract placeholder text from a ContentControl, handling COM Range objects."""
    try:
        pt = com_property(cc, "PlaceholderText", None)
        if pt is None:
            return ""
        if isinstance(pt, str):
            return pt
        return com_property(pt, "Text", "")
    except Exception:
        return ""


def count(doc_index: Optional[int] = None) -> Dict:
    """Count content controls in the document."""
    doc = get_doc(doc_index)
    total = doc.ContentControls.Count
    types = {}
    for i in range(1, total + 1):
        try:
            cc = doc.ContentControls.Item(i)
            t = com_property(cc, "Type", -1)
            types[t] = types.get(t, 0) + 1
        except Exception:
            pass
    type_names = {}
    for type_val in types:
        for name, val in WDCONTENTCONTROL.items():
            if val == type_val:
                type_names[type_val] = name
                break
    return {
        "total": total,
        "by_type": {
            type_names.get(t, f"unknown_{t}"): c for t, c in types.items()
        },
    }


def list_controls(doc_index: Optional[int] = None) -> Dict:
    """List all content controls with details."""
    doc = get_doc(doc_index)
    result = []
    total = doc.ContentControls.Count
    for i in range(1, total + 1):
        try:
            cc = doc.ContentControls.Item(i)
            t = com_property(cc, "Type", -1)
            result.append({
                "index": i,
                "type": t,
                "title": com_property(cc, "Title", ""),
                "tag": com_property(cc, "Tag", ""),
                "placeholder": _safe_placeholder(cc),
                "text": com_property(cc, "Text", "")[:100],
                "checked": com_property(cc, "Checked", False) if t == 8 else None,
                "removable": com_property(cc, "LockContents", False),
            })
        except Exception:
            continue
    return {"controls": result, "total": len(result)}


def info(cc_index: int, doc_index: Optional[int] = None) -> Dict:
    """Get detailed info about a specific content control."""
    doc = get_doc(doc_index)
    try:
        cc = _get_cc(doc, cc_index)
    except Exception:
        return {"error": f"Content control {cc_index} not found", "error_code": "CC_NOT_FOUND"}
    try:
        t = com_property(cc, "Type", -1)
    except Exception:
        return {"error": "Failed to read content control type", "error_code": "COM_ERROR"}
    info = {
        "index": cc_index,
        "type": t,
        "title": com_property(cc, "Title", ""),
        "tag": com_property(cc, "Tag", ""),
        "text": com_property(cc.Range, "Text", ""),
        "placeholder": _safe_placeholder(cc),
        "color": com_property(cc, "Color", 0),
    }
    if t == 4:  # Dropdown
        items = []
        try:
            for j in range(1, cc.DropdownListEntries.Count + 1):
                items.append(com_property(cc.DropdownListEntries.Item(j), "Text", ""))
        except Exception:
            pass
        info["dropdown_items"] = items
    elif t == 8:  # Checkbox
        info["checked"] = com_property(cc, "Checked", False)
    elif t == 6:  # Date picker
        info["date_value"] = com_property(cc.Range, "Text", "")
    return info


def add(type_name: str, text: str = "", title: str = "",
        para_index: Optional[int] = None,
        position: str = "end",
        dropdown_items: Optional[List[str]] = None,
        date_format: Optional[str] = None,
        doc_index: Optional[int] = None) -> Dict:
    """Add a content control to the document.
    type_name: RICH_TEXT/TEXT/DATE_PICKER/DROPDOWN_LIST/CHECKBOX/PICTURE
    """
    doc = get_doc(doc_index)
    cc_type = WDCONTENTCONTROL.get(type_name.upper())
    if cc_type is None:
        return {"error": f"Unknown type: {type_name}", "valid_types": list(WDCONTENTCONTROL.keys())}

    if position == "end":
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    elif position == "before" and para_index:
        rng = doc.Paragraphs.Item(para_index).Range
    elif position == "after" and para_index:
        rng = doc.Paragraphs.Item(para_index).Range
        rng.Collapse(0)
    else:
        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)

    cc = doc.ContentControls.Add(cc_type, rng)
    if title:
        com_set(cc, "Title", title)
    if text:
        cc.Range.Text = text

    if dropdown_items and cc_type in (3, 4):
        for item in dropdown_items:
            cc.DropdownListEntries.Add(item)

    if date_format and cc_type == 6:
        try:
            com_set(cc, "DateFormat", date_format)
        except Exception:
            pass

    return {
        "added": True,
        "type": type_name,
        "title": title,
        "text": text[:80],
    }


def set_text(cc_index: int, text: str, doc_index: Optional[int] = None) -> Dict:
    """Set text of a content control."""
    doc = get_doc(doc_index)
    cc = _get_cc(doc, cc_index)
    cc.Range.Text = text
    return {"set": True, "text": text[:80]}


def set_checkbox(cc_index: int, checked: bool, doc_index: Optional[int] = None) -> Dict:
    """Set checkbox state."""
    doc = get_doc(doc_index)
    cc = _get_cc(doc, cc_index)
    com_set(cc, "Checked", checked)
    return {"checked": checked}


def select_dropdown(cc_index: int, item_text: str, doc_index: Optional[int] = None) -> Dict:
    """Select a dropdown item by text."""
    doc = get_doc(doc_index)
    cc = _get_cc(doc, cc_index)
    items = cc.DropdownListEntries
    for i in range(1, items.Count + 1):
        if com_property(items.Item(i), "Text", "") == item_text:
            items.Item(i).Select()
            return {"selected": item_text}
    return {"error": f"Item '{item_text}' not found in dropdown", "available": [
        com_property(items.Item(j), "Text", "") for j in range(1, items.Count + 1)
    ]}


def delete(cc_index: int, doc_index: Optional[int] = None) -> Dict:
    """Delete a content control, keeping its text."""
    doc = get_doc(doc_index)
    cc = _get_cc(doc, cc_index)
    text = com_property(cc.Range, "Text", "")
    cc.Delete(True)  # True = keep content
    return {"deleted": True, "kept_text": text[:80]}


def set_tag(cc_index: int, tag: str, doc_index: Optional[int] = None) -> Dict:
    """Set the tag of a content control (useful for programmatic access)."""
    doc = get_doc(doc_index)
    cc = _get_cc(doc, cc_index)
    com_set(cc, "Tag", tag)
    return {"tag": tag}


def find_by_tag(tag: str, doc_index: Optional[int] = None) -> Dict:
    """Find content controls by tag value."""
    doc = get_doc(doc_index)
    results = []
    total = doc.ContentControls.Count
    for i in range(1, total + 1):
        try:
            cc = doc.ContentControls.Item(i)
            if com_property(cc, "Tag", "") == tag:
                results.append({
                    "index": i,
                    "title": com_property(cc, "Title", ""),
                    "text": com_property(cc.Range, "Text", "")[:80],
                })
        except Exception:
            continue
    return {"tag": tag, "found": len(results), "controls": results}
