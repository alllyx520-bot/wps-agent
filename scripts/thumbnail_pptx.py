# -*- coding: utf-8 -*-
"""Generate slide thumbnail images from .pptx using LibreOffice + pdftoppm.
Usage: python scripts/thumbnail_pptx.py <presentation.pptx> [output_prefix]
"""
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path


def _find_soffice() -> str:
    candidates = [
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
        "/usr/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    for cmd in ["soffice", "libreoffice"]:
        try:
            result = subprocess.run([cmd, "--version"], capture_output=True, timeout=5)
            if result.returncode == 0:
                return cmd
        except Exception:
            continue
    return None


def generate_thumbnails(filepath: str, output_prefix: str = "slide", dpi: int = 150) -> dict:
    """Convert PPTX to PDF, then render each slide as JPEG image.
    Images saved as {output_prefix}-01.jpg, {output_prefix}-02.jpg, etc.
    """
    if not os.path.exists(filepath):
        return {"error": f"File not found: {filepath}"}
    if not filepath.lower().endswith(".pptx"):
        return {"error": "Not a .pptx file"}

    soffice = _find_soffice()
    if not soffice:
        return {"error": "LibreOffice not found"}

    abs_path = os.path.abspath(filepath)
    work_dir = tempfile.mkdtemp(prefix="thumbnail_")

    # Step 1: PPTX → PDF
    try:
        subprocess.run([
            soffice, "--headless", "--norestore",
            "--convert-to", "pdf", "--outdir", work_dir, abs_path,
        ], capture_output=True, text=True, timeout=120)
    except subprocess.TimeoutExpired:
        return {"error": "PDF conversion timed out"}

    pdf_files = list(Path(work_dir).glob("*.pdf"))
    if not pdf_files:
        return {"error": "PDF conversion produced no output"}

    pdf_path = str(pdf_files[0])

    # Step 2: PDF → JPEG images
    output_dir = os.path.dirname(os.path.abspath(filepath))
    try:
        subprocess.run([
            "pdftoppm", "-jpeg", "-r", str(dpi), pdf_path,
            os.path.join(output_dir, output_prefix),
        ], capture_output=True, text=True, timeout=60)
    except FileNotFoundError:
        return {
            "error": "pdftoppm not found. Install poppler-utils.",
            "pdf_output": pdf_path,
            "note": "PDF was generated but slide images could not be rendered",
        }
    except subprocess.TimeoutExpired:
        return {"error": "Image conversion timed out", "pdf_output": pdf_path}

    # Find generated images
    images = sorted(Path(output_dir).glob(f"{output_prefix}-*.jpg"))
    slides = [{"file": str(img), "slide_number": i + 1} for i, img in enumerate(images)]

    return {
        "status": "success",
        "output_dir": output_dir,
        "prefix": output_prefix,
        "slide_count": len(slides),
        "slides": slides,
        "dpi": dpi,
    }


def visual_qa_check(images_dir: str, prefix: str = "slide", expectations: list = None) -> dict:
    """Check slide images for visual issues. Returns list of potential problems.
    This is meant to be used by subagent with vision capabilities.
    expectations = ["Slide 1: Title and subtitle", "Slide 2: 3-column layout with icons"]
    """
    from pathlib import Path
    slides = sorted(Path(images_dir).glob(f"{prefix}-*.jpg"))
    return {
        "slides": [str(s) for s in slides],
        "expectations": expectations or [],
        "checklist": [
            "Check for overlapping elements (text through shapes)",
            "Check for text overflow or cut-off at edges",
            "Check for elements too close (< 0.3in gaps)",
            "Check for insufficient margin from slide edges (< 0.5in)",
            "Check for low-contrast text on backgrounds",
            "Check for leftover placeholder content",
        ],
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python thumbnail_pptx.py <presentation.pptx> [output_prefix]")
        sys.exit(1)
    prefix = sys.argv[2] if len(sys.argv) > 2 else "slide"
    result = generate_thumbnails(sys.argv[1], prefix)
    print(json.dumps(result, indent=2, ensure_ascii=False))
