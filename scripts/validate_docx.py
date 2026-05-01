# -*- coding: utf-8 -*-
"""Validate .docx XML structure — check schema issues, auto-repair when possible.
Usage: python scripts/validate_docx.py <file.docx>
"""
import zipfile
import json
import os
import sys
import re
import shutil
from pathlib import Path


def validate_docx(filepath: str, auto_fix: bool = False) -> dict:
    """Validate a .docx file's XML structure. Returns issues found.
    Checks: durableId overflow, missing xml:space, run count sanity."""
    if not os.path.exists(filepath):
        return {"error": f"File not found: {filepath}"}
    if not filepath.lower().endswith(".docx"):
        return {"error": "Not a .docx file"}

    issues = []
    fixed = []

    # Verify it's a valid ZIP
    try:
        with zipfile.ZipFile(filepath, "r") as zf:
            zf.testzip()
    except zipfile.BadZipFile as e:
        return {"error": f"Corrupted .docx (invalid ZIP): {e}"}

    # If auto-fix, work on a copy
    if auto_fix:
        backup = filepath + ".bak"
        shutil.copy2(filepath, backup)

    # Read main document XML
    try:
        with zipfile.ZipFile(filepath, "r") as zf:
            if "word/document.xml" in zf.namelist():
                xml_content = zf.read("word/document.xml").decode("utf-8", errors="replace")
                # Check durableId overflow (> 0x7FFFFFFF)
                for m in re.finditer(r'durableId="(\d+)"', xml_content):
                    val = int(m.group(1))
                    if val >= 0x7FFFFFFF:
                        issues.append({
                            "type": "durableId_overflow",
                            "file": "word/document.xml",
                            "detail": f"durableId={val} exceeds max safe value (0x7FFFFFFF)",
                            "fixed": False,
                        })
                # Check for missing xml:space on w:t with leading/trailing spaces
                for m in re.finditer(r'<w:t(?![^>]*xml:space)[^>]*>([ \t].*?|[^<]*?[ \t])</w:t>', xml_content):
                    issues.append({
                        "type": "missing_xml_space",
                        "file": "word/document.xml",
                        "detail": f"w:t with whitespace but no xml:space='preserve'",
                        "fixed": False,
                    })
            else:
                issues.append({"type": "structure", "detail": "word/document.xml not found"})
    except Exception as e:
        issues.append({"type": "read_error", "detail": str(e)})

    # Paragraph count sanity
    try:
        with zipfile.ZipFile(filepath, "r") as zf:
            xml_content = zf.read("word/document.xml").decode("utf-8", errors="replace")
            para_count = len(re.findall(r'<w:p[\s>]', xml_content))
            if para_count <= 1:
                issues.append({"type": "empty_document", "detail": f"Only {para_count} paragraph(s) found"})
    except Exception:
        pass

    # Smart quote detection
    try:
        with zipfile.ZipFile(filepath, "r") as zf:
            xml_content = zf.read("word/document.xml").decode("utf-8", errors="replace")
            if re.search(r'[\\u2018\\u2019\\u201C\\u201D]', xml_content):
                issues.append({
                    "type": "smart_quotes_ascii",
                    "detail": "Smart quotes found as unicode chars — should be XML entities (&#x201C; etc.)",
                    "fixed": False,
                })
    except Exception:
        pass

    if auto_fix and issues:
        # Attempt to repair durableId overflow by patching XML
        _repair_durable_ids(filepath, issues, fixed)

    status = "ok" if not issues else "issues_found"
    result = {"status": status, "file": filepath, "issues_count": len(issues), "issues": issues, "fixed_count": len(fixed), "fixed": fixed}
    if auto_fix:
        result["backup"] = filepath + ".bak"
    return result


def _repair_durable_ids(filepath: str, issues: list, fixed: list):
    """Patch durableId values in document.xml."""
    try:
        tmp = filepath + ".tmp"
        with zipfile.ZipFile(filepath, "r") as zin:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "word/document.xml":
                        text = data.decode("utf-8", errors="replace")
                        repaired = 0
                        def replace_id(m):
                            nonlocal repaired
                            val = int(m.group(2))
                            if val >= 0x7FFFFFFF:
                                repaired += 1
                                new_val = val & 0x7FFFFFFE
                                return f'{m.group(1)}"{new_val}"'
                            return m.group(0)
                        text = re.sub(r'(durableId)="(\d+)"', replace_id, text)
                        data = text.encode("utf-8")
                        if repaired > 0:
                            fixed.append({"type": "durableId_overflow", "count": repaired})
                    zout.writestr(item, data)
        os.replace(tmp, filepath)
    except Exception:
        pass


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python validate_docx.py <file.docx> [--fix]")
        sys.exit(1)
    auto_fix = "--fix" in sys.argv
    result = validate_docx(sys.argv[1], auto_fix)
    print(json.dumps(result, indent=2, ensure_ascii=False))
