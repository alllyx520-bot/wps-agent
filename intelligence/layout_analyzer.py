# -*- coding: utf-8 -*-
from typing import Optional, Dict, List
import logging
from .llm_client import analyze_document_structure

logger = logging.getLogger("wps-agent.analyzer")


def analyze(doc_index: Optional[int] = None) -> Dict:
    from wps_bridge.content import outline, paragraph
    from wps_bridge.document import doc_info

    info = doc_info(doc_index)
    outline_data = outline(doc_index)

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
        llm_result = analyze_document_structure(outline_data, format_samples)
    except Exception as e:
        logger.warning(f"LLM analysis failed: {e}")

    return {
        "document": info,
        "outline": outline_data,
        "format_samples": format_samples,
        "llm_analysis": llm_result,
    }


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
