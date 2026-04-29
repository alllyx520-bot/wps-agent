# -*- coding: utf-8 -*-
from typing import Optional, List, Dict
import logging
from .llm_client import suggest_formatting

logger = logging.getLogger("wps-agent.suggester")


def suggest(doc_index: Optional[int] = None) -> Dict:
    from wps_bridge.content import outline, paragraph

    outline_data = outline(doc_index)
    format_samples = []
    for h in outline_data[:15]:
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

    llm_suggestions = None
    try:
        llm_suggestions = suggest_formatting(outline_data, format_samples)
    except Exception as e:
        logger.warning(f"LLM suggestions failed: {e}")

    return {
        "outline_count": len(outline_data),
        "format_samples": format_samples[:20],
        "llm_suggestions": llm_suggestions,
    }
