# -*- coding: utf-8 -*-
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set


def find_text(query, match_case=False, whole_word=False, doc_index=None):
    """Selection-based find: searches forward using Range, advancing past each match.
    This approach works reliably with Chinese/CJK text where Duplicate+Collapse fails."""
    doc = get_doc(doc_index)
    if not query:
        return []
    results = []
    search_pos = 0
    content_end = int(com_property(doc.Content, "End", 0))
    max_iter = 5000
    for _ in range(max_iter):
        if search_pos >= content_end:
            break
        rng = doc.Range(search_pos, content_end)
        find_obj = rng.Find
        find_obj.ClearFormatting()
        find_obj.Text = query
        find_obj.MatchCase = match_case
        find_obj.MatchWholeWord = whole_word
        find_obj.Forward = True
        find_obj.Wrap = 0
        find_obj.Format = False
        try:
            found = find_obj.Execute()
        except Exception:
            break
        if not found:
            break
        start = int(com_property(rng, "Start", 0))
        end = int(com_property(rng, "End", 0))
        if end <= search_pos:
            search_pos += 1
            continue
        txt = str(com_property(rng, "Text", ""))
        results.append({"text": txt[:100], "start": start, "end": end})
        search_pos = end
        if len(results) > 500:
            break
    return results


def replace_text(find_text_val, replace_text_val, match_case=False, replace_all=False, doc_index=None):
    """Replace text using Word Find+Replace via Executed method."""
    doc = get_doc(doc_index)
    if not find_text_val:
        return {"replaced": False, "success": False, "error": "find_text is empty"}
    rng = doc.Content
    find_obj = rng.Find
    find_obj.ClearFormatting()
    find_obj.Replacement.ClearFormatting()
    find_obj.Text = find_text_val
    find_obj.Replacement.Text = replace_text_val
    find_obj.Forward = True
    find_obj.Wrap = 0
    find_obj.MatchCase = match_case
    find_obj.MatchWholeWord = False
    find_obj.MatchWildcards = False
    find_obj.MatchSoundsLike = False
    find_obj.MatchAllWordForms = False
    find_obj.Format = False
    replace_mode = 2 if replace_all else 1
    count = 0
    try:
        ok = find_obj.Execute(FindText=find_text_val, MatchCase=match_case,
                              MatchWholeWord=False, MatchWildcards=False,
                              MatchSoundsLike=False, MatchAllWordForms=False,
                              Forward=True, Wrap=0, Format=False,
                              ReplaceWith=replace_text_val, Replace=replace_mode)
        if ok:
            count = -1 if replace_all else 1
    except Exception as e:
        return {"replaced": False, "success": False, "error": str(e)}
    return {"replaced": "all" if replace_all else bool(count > 0), "success": True}


def find_format(font_name=None, font_size=None, bold=None, style_name=None, doc_index=None):
    doc = get_doc(doc_index)
    find = doc.Content.Find
    find.ClearFormatting()
    if font_name:
        find.Font.Name = font_name
    if font_size:
        find.Font.Size = font_size
    if bold is not None:
        find.Font.Bold = bold
    if style_name:
        find.Style = doc.Styles.Item(style_name)
    results = []
    find.Wrap = 0
    while find.Execute(FindText="", Format=True):
        parent = find.Parent
        results.append({
            "text": str(com_property(parent, "Text", ""))[:100],
            "start": int(com_property(parent, "Start", 0)),
            "end": int(com_property(parent, "End", 0)),
        })
    return results


def goto_heading(text=None, level=None, doc_index=None):
    doc = get_doc(doc_index)
    sel = get_app().Selection
    if text:
        sel.HomeKey(6)
        sel.Find.ClearFormatting()
        sel.Find.Text = text
        sel.Find.Execute()
        return {"found": text, "paragraph": str(sel.Range.Paragraphs.Item(1).Range.Text).strip()[:80]}
    elif level:
        for i in range(1, doc.Paragraphs.Count + 1):
            if com_property(doc.Paragraphs.Item(i).Format, "OutlineLevel", 10) == level:
                sel.GoTo(-1, 0, 0, doc.Paragraphs.Item(i).Range.Text)
                return {"heading": str(doc.Paragraphs.Item(i).Range.Text).strip()[:80],
                        "level": level, "index": i}
    return {"found": False}
