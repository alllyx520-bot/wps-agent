# -*- coding: utf-8 -*-
from typing import Optional, Dict, List
import logging
import sys

logger = logging.getLogger("wps-agent.analyzer")


def analyze_offline(filepath: str) -> Dict:
    """Analyze document using offline docx_engine (no WPS required).
    Uses SemanticParser as primary analyzer, falls back to LLM enhancement.
    """
    from docx_engine import parse_docx, build_document_model
    from docx_engine.intelligence import DocumentAnalyzer
    from docx_engine.semantic_model import SemanticParser

    parsed = parse_docx(filepath)
    doc = build_document_model(parsed)

    # ── Rule-based semantic parsing ──
    parser = SemanticParser(doc)
    graph = parser.parse()

    # ── Legacy analysis ──
    analyzer = DocumentAnalyzer(doc)
    doc_type = analyzer.detect_document_type()
    quality = analyzer.analyze_formatting_quality()

    # ── Layout analysis ──
    layout = None
    try:
        from docx_engine.layout_model import LayoutAnalyzer
        la = LayoutAnalyzer(doc)
        layout = la.analyze().to_dict()
    except Exception as e:
        logger.warning(f"Layout analysis failed: {e}")

    return {
        "document_type": doc_type,
        "paragraph_count": len(doc.paragraphs),
        "table_count": len(doc.tables),
        "semantic_graph": graph.to_dict(),
        "quality": quality,
        "layout": layout,
    }


def analyze_online(doc_index: Optional[int] = None) -> Dict:
    """Analyze document using WPS COM bridge (requires WPS running).
    Includes semantic parsing via offline engine if available.
    """
    from wps_bridge.content import outline, paragraph, full_text
    from wps_bridge.document import doc_info

    info = doc_info(doc_index)
    outline_data = outline(doc_index)
    full = full_text(doc_index)

    format_samples = []
    for h in outline_data[:10]:
        try:
            p = paragraph(h["index"], doc_index)
            format_samples.append({
                "para_index": h["index"],
                "text": p["text"][:60],
                "outline_level": h["outline_level"],
                "font": p["font"],
                "paragraph_format": p["paragraph_format"],
            })
        except Exception:
            continue

    llm_result = None
    try:
        from .llm_client import analyze_document_structure
        llm_result = analyze_document_structure(outline_data, format_samples)
    except Exception as e:
        logger.warning(f"LLM analysis failed: {e}")

    return {
        "document": info,
        "outline": outline_data,
        "format_samples": format_samples,
        "full_text_length": len(full),
        "llm_analysis": llm_result,
    }


def analyze(doc_index: Optional[int] = None, filepath: Optional[str] = None) -> Dict:
    """Unified analyze entry: prefers offline when filepath is given."""
    if filepath:
        return analyze_offline(filepath)
    return analyze_online(doc_index)


def detect_semantic_roles(filepath: str) -> Dict:
    """Detect semantic roles for all paragraphs using rule-based parser."""
    from docx_engine import parse_docx, build_document_model
    from docx_engine.semantic_model import SemanticParser

    parsed = parse_docx(filepath)
    doc = build_document_model(parsed)
    parser = SemanticParser(doc)
    graph = parser.parse()

    roles = []
    for e in graph.elements:
        roles.append({
            "para_index": e.index,
            "role": e.role,
            "confidence": e.confidence,
            "text_preview": e.text_preview,
            "heading_level": e.heading_level,
        })

    return {
        "document_type": __import__('docx_engine.intelligence', fromlist=['DocumentAnalyzer']).DocumentAnalyzer(doc).detect_document_type(),
        "roles": roles,
        "heading_tree": graph.heading_tree,
        "document_zones": {
            "cover": graph.cover_indices,
            "front_matter": graph.front_matter_indices,
            "body": graph.body_indices,
            "back_matter": graph.back_matter_indices,
        },
    }


def analyze_layout(filepath: str) -> Dict:
    """Analyze document layout: page geometry, text flow, header/footer chain."""
    from docx_engine import parse_docx, build_document_model
    from docx_engine.layout_model import LayoutAnalyzer

    parsed = parse_docx(filepath)
    doc = build_document_model(parsed)
    la = LayoutAnalyzer(doc)
    return la.analyze().to_dict()


def generate_reformat_actions(instructions: str, doc_index: Optional[int] = None) -> List[Dict]:
    from wps_bridge.content import outline, paragraph

    outline_data = outline(doc_index)
    format_samples = []
    for h in outline_data[:10]:
        try:
            p = paragraph(h["index"], doc_index)
            format_samples.append({
                "para_index": h["index"],
                "font": p["font"],
                "paragraph_format": p["paragraph_format"],
            })
        except Exception:
            continue

    from .llm_client import parse_natural_language_instructions
    return parse_natural_language_instructions(instructions, outline_data, format_samples) or []
