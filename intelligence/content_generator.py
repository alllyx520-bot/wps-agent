# -*- coding: utf-8 -*-
from typing import Optional, Dict
from .llm_client import chat
from wps_bridge.content import full_text, paragraph, outline, insert_text


def generate_content(instructions: str, position: str = "end",
                     para_index: Optional[int] = None,
                     doc_index: Optional[int] = None) -> Dict:
    """Generate new content at a position."""
    context = full_text(doc_index)[:4000]
    prompt = f"Document context:\n{context}\n\nTask: {instructions}\n\nGenerate the requested content in Chinese. Output ONLY the content text, no explanation."
    result = chat("You are a professional Chinese document writer. Generate high-quality, relevant content based on the instructions.", prompt)
    if not result:
        return {"error": "LLM generation failed", "instructions": instructions}
    r = insert_text(result, position, para_index, doc_index)
    return {"generated": True, "position": position, "text_preview": result[:200], "insert_result": r}


def summarize_document(doc_index: Optional[int] = None) -> Dict:
    """Generate a summary of the document."""
    text = full_text(doc_index)[:8000]
    outline_data = outline(doc_index)
    prompt = f"Document text:\n{text}\n\nOutline:\n{str(outline_data)[:3000]}\n\nGenerate a concise summary in Chinese."
    result = chat("You are a professional document summarizer. Create a concise, well-structured summary highlighting the key points.", prompt)
    if not result:
        return {"error": "LLM summarization failed"}
    return {"summary": result, "original_length": len(text), "summary_length": len(result)}


def rewrite_paragraph(para_index: int, instructions: str = "",
                      doc_index: Optional[int] = None) -> Dict:
    """Rewrite or improve a specific paragraph."""
    p = paragraph(para_index, doc_index)
    original = p.get("text", "")
    prompt = f"Original paragraph:\n{original}\n\nTask: {instructions or 'Rewrite this paragraph to be more professional and polished'}.\n\nOutput ONLY the rewritten text, no explanation."
    result = chat("You are a professional Chinese editor. Improve the given text.", prompt)
    if not result:
        return {"error": "LLM rewrite failed", "para_index": para_index}
    from wps_bridge.app import get_doc
    doc = get_doc(doc_index)
    doc.Paragraphs.Item(para_index).Range.Text = result
    return {"para_index": para_index, "original_preview": original[:100], "rewritten_preview": result[:200]}


def expand_section(para_index: int, doc_index: Optional[int] = None) -> Dict:
    """Expand a section with more detailed content."""
    p = paragraph(para_index, doc_index)
    context = p.get("text", "")
    # Also get surrounding paragraphs for context
    try:
        next_p = paragraph(para_index + 1, doc_index)
        context += "\n\nAfter: " + next_p.get("text", "")[:500]
    except Exception:
        pass
    prompt = f"Section:\n{context}\n\nExpand this section with 1-2 more paragraphs of detailed, relevant content in Chinese. Output ONLY the expanded paragraphs (the new content to insert)."
    result = chat("You are a professional Chinese document writer. Expand the given section with well-structured, detailed content.", prompt)
    if not result:
        return {"error": "LLM expansion failed", "para_index": para_index}
    from wps_bridge.app import get_doc
    doc = get_doc(doc_index)
    rng = doc.Paragraphs.Item(para_index).Range
    rng.InsertAfter("\r" + result)
    return {"para_index": para_index, "expanded": True, "added_text_preview": result[:200]}


def translate_section(para_index: int, target_lang: str = "en",
                      doc_index: Optional[int] = None) -> Dict:
    """Translate a section to another language."""
    p = paragraph(para_index, doc_index)
    original = p.get("text", "")
    lang_map = {"en": "English", "ja": "Japanese", "ko": "Korean", "fr": "French", "de": "German", "zh": "Chinese"}
    target = lang_map.get(target_lang, target_lang)
    prompt = f"Translate the following text to {target}. Output ONLY the translation, no explanation.\n\nText:\n{original}"
    result = chat(f"You are a professional translator. Translate to {target} accurately.", prompt)
    if not result:
        return {"error": "LLM translation failed", "para_index": para_index}
    from wps_bridge.app import get_doc
    doc = get_doc(doc_index)
    rng = doc.Paragraphs.Item(para_index).Range
    rng.InsertAfter("\r[Translation:\r" + result + "]")
    return {"para_index": para_index, "target_lang": target, "original_preview": original[:100], "translation_preview": result[:200]}
