# -*- coding: utf-8 -*-
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set


def find_text(query, match_case=False, whole_word=False, doc_index=None):
    doc = get_doc(doc_index)
    find = doc.Content.Find
    find.Text = query
    find.MatchCase = match_case
    find.MatchWholeWord = whole_word
    find.Forward = True
    find.Wrap = 0
    results = []
    while find.Execute():
        rng = find.Parent
        results.append({"text": com_property(rng, "Text", "")[:100], "start": com_property(rng, "Start", 0), "end": com_property(rng, "End", 0)})
        if len(results) > 5000:
            break
    return results


def replace_text(find_text, replace_text, match_case=False, replace_all=False, doc_index=None):
    doc = get_doc(doc_index)
    find = doc.Content.Find
    find.Text = find_text
    find.MatchCase = match_case
    if replace_all:
        find.Execute(Replace=2)
        return {"replaced": "all"}
    else:
        found = find.Execute(Replace=1)
        return {"replaced": bool(found)}


def find_format(font_name=None, font_size=None, bold=None, style_name=None, doc_index=None):
    doc = get_doc(doc_index)
    find = doc.Content.Find
    find.ClearFormatting()
    if font_name: find.Font.Name = font_name
    if font_size: find.Font.Size = font_size
    if bold is not None: find.Font.Bold = bold
    if style_name: find.Style = doc.Styles.Item(style_name)
    results = []
    find.Wrap = 0
    while find.Execute(FindText="", Format=True):
        rng = find.Parent
        results.append({"text": com_property(rng, "Text", "")[:100], "start": com_property(rng, "Start", 0), "end": com_property(rng, "End", 0)})
    return results


def goto_heading(text=None, level=None, doc_index=None):
    doc = get_doc(doc_index)
    sel = get_app().Selection
    if text:
        sel.HomeKey(6)
        sel.Find.ClearFormatting()
        sel.Find.Text = text
        sel.Find.Execute()
        return {"found": text, "paragraph": sel.Range.Paragraphs.Item(1).Range.Text.strip()[:80]}
    elif level:
        for i in range(1, doc.Paragraphs.Count + 1):
            if com_property(doc.Paragraphs.Item(i).Format, "OutlineLevel", 10) == level:
                sel.GoTo(-1, 0, 0, doc.Paragraphs.Item(i).Range.Text)
                return {"heading": doc.Paragraphs.Item(i).Range.Text.strip()[:80], "level": level, "index": i}
    return {"found": False}
