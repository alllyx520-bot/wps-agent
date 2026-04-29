# -*- coding: utf-8 -*-
from typing import Optional, Dict
from .llm_client import chat
from wps_bridge.content import full_text, paragraph, outline, insert_text


def generate_content(instructions: str, position: str = "end",
                     para_index: Optional[int] = None,
                     doc_index: Optional[int] = None) -> Dict:
    """Generate new content at a position."""
    # Use outline for structural context, full text for detailed context
    outline_data = outline(doc_index)
    context = full_text(doc_index)[:2000]
    prompt = (
        f"文档大纲:\n{str(outline_data)[:1500]}\n\n"
        f"文档片段:\n{context}\n\n"
        f"任务: {instructions}\n\n"
        f"请根据以上文档结构和内容，生成专业、得体、符合中文公文规范的内容。"
        f"保持与原文风格一致，使用正式书面语。只输出生成的内容文本，不要说明。"
    )
    result = chat(
        "你是一位资深中文文档编辑，精通各类公文写作规范。"
        "你产出的内容格式规范、用词精准、逻辑清晰、风格得体。",
        prompt
    )
    if not result:
        return {"error": "LLM generation failed", "instructions": instructions}
    r = insert_text(result, position, para_index, doc_index)
    return {"generated": True, "position": position, "text_preview": result[:200], "insert_result": r}


def summarize_document(doc_index: Optional[int] = None) -> Dict:
    """Generate a summary of the document."""
    text = full_text(doc_index)[:6000]
    outline_data = outline(doc_index)
    prompt = (
        f"文档内容:\n{text}\n\n"
        f"文档大纲:\n{str(outline_data)[:2000]}\n\n"
        f"请根据以上文档生成一份专业摘要。要求：\n"
        f"1. 先概括文档主题和目的\n"
        f"2. 按章节顺序概述各章节要点\n"
        f"3. 语言精炼，200-400字\n"
        f"4. 使用正式书面语"
    )
    result = chat(
        "你是一位资深文档分析专家，擅长提炼和总结中文文档的核心内容。"
        "你的摘要结构清晰、重点突出、语言精炼。",
        prompt
    )
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
    from wps_bridge.utils import com_property, com_set
    doc = get_doc(doc_index)
    p = doc.Paragraphs.Item(para_index)
    # Save original formatting
    f = p.Range.Font
    saved = {"Name": com_property(f, "Name", ""), "NameFarEast": com_property(f, "NameFarEast", ""),
             "Size": com_property(f, "Size", 0), "Bold": com_property(f, "Bold", 0),
             "Italic": com_property(f, "Italic", 0), "ColorIndex": com_property(f, "ColorIndex", 1)}
    # Replace text
    p.Range.Text = result
    # Restore original formatting
    for k, v in saved.items():
        if v:
            com_set(p.Range.Font, k, v)
    return {"para_index": para_index, "original_preview": original[:100], "rewritten_preview": result[:200]}


def expand_section(para_index: int, doc_index: Optional[int] = None) -> Dict:
    """Expand a section with more detailed content."""
    p = paragraph(para_index, doc_index)
    context = p.get("text", "")
    # Get surrounding paragraphs for context (with bounds check)
    try:
        from wps_bridge.app import get_doc
        doc = get_doc(doc_index)
        if para_index < doc.Paragraphs.Count:
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
    # Collect all parts to insert, then insert at once to avoid position drift
    parts_to_insert = [part for part in result.split("\n") if part.strip()]
    if parts_to_insert:
        insertion_text = "\r" + "\r".join(parts_to_insert)
        rng.InsertAfter(insertion_text)
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
    from wps_bridge.utils import com_property, com_set
    doc = get_doc(doc_index)
    p = doc.Paragraphs.Item(para_index)
    # Save original formatting for the translation block
    f = p.Range.Font
    saved = {"Name": com_property(f, "Name", ""), "NameFarEast": com_property(f, "NameFarEast", ""),
             "Size": com_property(f, "Size", 0), "Bold": com_property(f, "Bold", 0),
             "ColorIndex": com_property(f, "ColorIndex", 1)}
    # Insert translation as a new paragraph after the original, preserving original as-is
    rng = p.Range
    rng.InsertAfter("\r" + result)
    # Format the newly inserted translation paragraph to match
    try:
        next_p = doc.Paragraphs.Item(para_index + 1)
        for k, v in saved.items():
            if v:
                com_set(next_p.Range.Font, k, v)
    except Exception:
        pass
    return {"para_index": para_index, "target_lang": target, "original_preview": original[:100], "translation_preview": result[:200]}
