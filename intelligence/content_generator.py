# -*- coding: utf-8 -*-
from typing import Optional, Dict, List
from .llm_client import chat
from wps_bridge.content import full_text, paragraph, outline, insert_text
from wps_bridge.app import get_doc, get_app
from wps_bridge.utils import com_property, com_set
from wps_bridge.formatting import set_font, set_paragraph_format


def _smart_cover(instructions: str, doc_index: Optional[int] = None) -> Dict:
    """Generate a cover page with proper spacing and formatting."""
    prompt = f"""根据以下要求，生成一份标准中文封面内容。每行一条，不要编号前缀。
{instructions}

标准封面结构：
第一行：主标题（如"实验报告""项目建议书"等，要精确具体）
第二行：副标题或补充信息（如有）
空行
第三行：单位/部门名称
第四行：姓名（用"姓名：___"格式）
空行
第五行：日期（用"202X年X月X日"格式）

输出纯文本，每行是一条，最多6行。标题要准确、正式。不要加任何说明文字。"""
    result = chat("你是资深文档排版专家，生成的封面内容格式规范、信息完整、可直接使用。", prompt)
    if not result:
        return {"error": "LLM generation failed"}
    return _build_cover_page(result.strip(), doc_index)


def _build_cover_page(content_text: str, doc_index: Optional[int] = None) -> Dict:
    """Build a cover page from LLM output: center everything, proper spacing, page break."""
    doc = get_doc(doc_index)
    lines = [l.strip() for l in content_text.split("\n") if l.strip()]
    if not lines:
        return {"error": "No content to insert"}

    # Clear existing content first
    try:
        doc.Content.Delete()
    except Exception:
        pass

    # Insert each line as a separate paragraph
    app = get_app()
    sel = app.Selection
    sel.HomeKey(6)  # Go to start

    for i, line in enumerate(lines):
        if i > 0:
            sel.TypeParagraph()
        sel.TypeText(line)

    # Now format each paragraph
    para_count = doc.Paragraphs.Count
    formatted = 0
    for i in range(1, para_count + 1):
        try:
            p = doc.Paragraphs.Item(i)
            text = com_property(p.Range, "Text", "").strip()
            if not text:
                # Empty paragraph: add small spacing
                com_set(p.Format, "SpaceBefore", 6)
                com_set(p.Format, "SpaceAfter", 6)
                continue

            f = p.Range.Font
            pf = p.Format
            com_set(f, "ColorIndex", 1)

            if i == 1:
                # Title paragraph
                com_set(f, "NameFarEast", "黑体")
                com_set(f, "Size", 26)
                com_set(f, "Bold", True)
                com_set(pf, "Alignment", 1)  # center
                com_set(pf, "SpaceBefore", 144)  # ~5cm from top
                com_set(pf, "SpaceAfter", 36)
            elif i == 2 and len(lines) >= 4:
                # Subtitle paragraph
                com_set(f, "NameFarEast", "宋体")
                com_set(f, "Size", 16)
                com_set(pf, "Alignment", 1)
                com_set(pf, "SpaceAfter", 24)
            elif i == len(lines) - 1:
                # Date paragraph (last)
                com_set(f, "NameFarEast", "宋体")
                com_set(f, "Size", 14)
                com_set(pf, "Alignment", 1)
                com_set(pf, "SpaceBefore", 36)
            elif i == len(lines) - 2:
                # Author paragraph (second to last)
                com_set(f, "NameFarEast", "宋体")
                com_set(f, "Size", 14)
                com_set(pf, "Alignment", 1)
                com_set(pf, "SpaceBefore", 12)
            else:
                # Middle paragraphs (unit/org)
                com_set(f, "NameFarEast", "宋体")
                com_set(f, "Size", 14)
                com_set(pf, "Alignment", 1)
                com_set(pf, "SpaceBefore", 6)
                com_set(pf, "SpaceAfter", 6)

            formatted += 1
        except Exception:
            continue

    # Add page break after cover
    try:
        last_para = doc.Paragraphs.Item(para_count)
        last_para.Range.InsertParagraphAfter()
        next_para = doc.Paragraphs.Item(para_count + 1)
        next_para.Range.InsertBreak(2)  # wdPageBreak
    except Exception:
        pass

    return {
        "cover_created": True,
        "paragraphs": formatted,
        "content": lines,
        "page_break_added": True,
    }


def generate_content(instructions: str, position: str = "end",
                     para_index: Optional[int] = None,
                     doc_index: Optional[int] = None) -> Dict:
    """Generate content. Automatically detects cover page requests."""
    is_cover = any(w in instructions for w in ["封面", "首页", "标题页", "cover", "第一页"])

    if is_cover:
        return _smart_cover(instructions, doc_index)

    context = full_text(doc_index)[:2000]
    outline_data = outline(doc_index)
    prompt = (
        f"文档大纲:\n{str(outline_data)[:1500]}\n\n"
        f"文档片段:\n{context}\n\n"
        f"任务: {instructions}\n\n"
        f"生成专业、得体的中文内容。保持与原文风格一致。只输出内容文本。"
    )
    result = chat("你是资深中文文档编辑，精通各类公文写作规范，产出格式规范、用词精准、逻辑清晰。", prompt)
    if not result:
        return {"error": "LLM generation failed"}
    r = insert_text(result, position, para_index, doc_index)
    return {"generated": True, "position": position, "text_preview": result[:200], "insert_result": r}


def summarize_document(doc_index: Optional[int] = None) -> Dict:
    text = full_text(doc_index)[:6000]
    outline_data = outline(doc_index)
    prompt = (
        f"文档内容:\n{text}\n\n"
        f"文档大纲:\n{str(outline_data)[:2000]}\n\n"
        f"生成一份专业摘要（200-400字）。先概括主题和目的，再按章节概述要点。使用正式书面语。"
    )
    result = chat("你是资深文档分析专家，擅长提炼和总结中文文档核心内容。摘要结构清晰、重点突出、语言精炼。", prompt)
    if not result:
        return {"error": "LLM summarization failed"}
    return {"summary": result, "original_length": len(text), "summary_length": len(result)}


def rewrite_paragraph(para_index: int, instructions: str = "",
                      doc_index: Optional[int] = None) -> Dict:
    p = paragraph(para_index, doc_index)
    original = p.get("text", "")
    prompt = f"原文:\n{original}\n\n要求: {instructions or '改写得更专业、更通顺'}.\n\n只输出改写后的文本。"
    result = chat("你是资深中文编辑，改写文本使其更专业、通顺。", prompt)
    if not result:
        return {"error": "LLM rewrite failed"}
    doc = get_doc(doc_index)
    pp = doc.Paragraphs.Item(para_index)
    f = pp.Range.Font
    saved = {"Name": com_property(f, "Name", ""), "NameFarEast": com_property(f, "NameFarEast", ""),
             "Size": com_property(f, "Size", 0), "Bold": com_property(f, "Bold", 0),
             "Italic": com_property(f, "Italic", 0), "ColorIndex": com_property(f, "ColorIndex", 1)}
    pp.Range.Text = result
    for k, v in saved.items():
        if v:
            com_set(pp.Range.Font, k, v)
    return {"para_index": para_index, "original_preview": original[:100], "rewritten_preview": result[:200]}


def expand_section(para_index: int, doc_index: Optional[int] = None) -> Dict:
    p = paragraph(para_index, doc_index)
    context = p.get("text", "")
    try:
        doc = get_doc(doc_index)
        if para_index < doc.Paragraphs.Count:
            next_p = paragraph(para_index + 1, doc_index)
            context += "\n\nAfter: " + next_p.get("text", "")[:500]
    except Exception:
        pass
    prompt = f"章节内容:\n{context}\n\n扩写1-2段更详细的相关内容。只输出新增的段落文本。"
    result = chat("你是资深中文文档作家，能产出结构完整、内容详实的扩写内容。", prompt)
    if not result:
        return {"error": "LLM expansion failed"}
    doc = get_doc(doc_index)
    rng = doc.Paragraphs.Item(para_index).Range
    parts = [p for p in result.split("\n") if p.strip()]
    if parts:
        rng.InsertAfter("\r" + "\r".join(parts))
    return {"para_index": para_index, "expanded": True, "added_text_preview": result[:200]}


def translate_section(para_index: int, target_lang: str = "en",
                      doc_index: Optional[int] = None) -> Dict:
    p = paragraph(para_index, doc_index)
    original = p.get("text", "")
    lang_map = {"en": "English", "ja": "Japanese", "ko": "Korean", "fr": "French", "de": "German", "zh": "Chinese"}
    target = lang_map.get(target_lang, target_lang)
    prompt = f"翻译以下文本为{target}。只输出译文。\n\n{original}"
    result = chat(f"你是专业翻译，准确翻译为{target}。", prompt)
    if not result:
        return {"error": "LLM translation failed"}
    doc = get_doc(doc_index)
    pp = doc.Paragraphs.Item(para_index)
    f = pp.Range.Font
    saved = {"Name": com_property(f, "Name", ""), "NameFarEast": com_property(f, "NameFarEast", ""),
             "Size": com_property(f, "Size", 0), "Bold": com_property(f, "Bold", 0),
             "ColorIndex": com_property(f, "ColorIndex", 1)}
    rng = pp.Range
    rng.InsertAfter("\r" + result)
    try:
        next_p = doc.Paragraphs.Item(para_index + 1)
        for k, v in saved.items():
            if v:
                com_set(next_p.Range.Font, k, v)
    except Exception:
        pass
    return {"para_index": para_index, "target_lang": target, "original_preview": original[:100], "translation_preview": result[:200]}
