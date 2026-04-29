# -*- coding: utf-8 -*-
"""
Cross-application data migration: Word↔Excel, Word→PPT
"""
from typing import Any, Dict
from .app import get_doc as get_word_doc
from .excel_app import ExcelApplication
from .docspace import get_word_doc_by_id, get_excel_wb_by_id
from .utils import com_property, com_set
from .ppt_app import PPTApplication


def word_table_to_excel(word_doc_id: str, table_index: int,
                        excel_doc_id: str, target_cell: str = "A1",
                        keep_format: bool = False) -> Dict:
    wd_doc = get_word_doc_by_id(word_doc_id)
    xl_wb = get_excel_wb_by_id(excel_doc_id)
    ws = xl_wb.ActiveSheet

    tbl = wd_doc.Tables.Item(table_index)
    rows = tbl.Rows.Count
    cols = tbl.Columns.Count

    data = []
    for r in range(1, rows + 1):
        row_data = []
        for c in range(1, cols + 1):
            try:
                row_data.append(tbl.Cell(r, c).Range.Text.replace("\r\x07", "").strip())
            except Exception:
                row_data.append("")
        data.append(row_data)

    # Calculate target range
    col_letter = chr(ord('A') + (ord(target_cell[0].upper()) - ord('A') + cols - 1) % 26)
    if len(target_cell) == 2 and target_cell[1:].isdigit():
        start_row = int(target_cell[1:])
    elif len(target_cell) >= 3 and target_cell[1:].isdigit():
        start_row = int(target_cell[1:])
    else:
        start_row = 1
    end_cell = f"{col_letter}{start_row + rows - 1}"
    rng = ws.Range(target_cell, end_cell)
    rng.Value = data

    return {"action": "word_table_to_excel", "rows": rows, "cols": cols, "target_range": f"{target_cell}:{end_cell}"}


def excel_range_to_word_table(excel_doc_id: str, range_start: str, range_end: str,
                               word_doc_id: str, position: str = "end",
                               keep_format: bool = False) -> Dict:
    xl_wb = get_excel_wb_by_id(excel_doc_id)
    wd_doc = get_word_doc_by_id(word_doc_id)
    ws = xl_wb.ActiveSheet

    rng = ws.Range(range_start, range_end)
    data = rng.Value
    if data is None:
        return {"error": "No data in specified range"}

    if not isinstance(data, tuple):
        data = ((data,),)
    elif not isinstance(data[0], tuple):
        data = tuple((v,) for v in data)

    rows = len(data)
    cols = len(data[0])

    # Insert position
    if position == "end":
        insert_rng = wd_doc.Range(wd_doc.Content.End - 1, wd_doc.Content.End - 1)
    elif position.isdigit():
        insert_rng = wd_doc.Paragraphs.Item(int(position)).Range
    else:
        insert_rng = wd_doc.Range(wd_doc.Content.End - 1, wd_doc.Content.End - 1)

    table = wd_doc.Tables.Add(insert_rng, rows, cols)
    table.AutoFitBehavior(2)

    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            try:
                val = data[r - 1][c - 1]
                if val is None:
                    val = ""
                table.Cell(r, c).Range.Text = str(val)
            except Exception:
                continue

    return {"action": "excel_range_to_word_table", "rows": rows, "cols": cols,
            "target_table_index": wd_doc.Tables.Count}


def word_outline_to_ppt(word_doc_id: str) -> Dict:
    wd_doc = get_word_doc_by_id(word_doc_id)
    ppt = PPTApplication()
    pres = ppt.app.Presentations.Add()

    # Read Word outline
    outlines = []
    for i in range(1, wd_doc.Paragraphs.Count + 1):
        try:
            p = wd_doc.Paragraphs.Item(i)
            level = com_property(p.Format, "OutlineLevel", 10)
            text = com_property(p.Range, "Text", "").strip()
            if 1 <= level <= 3 and text:
                outlines.append({"level": level, "text": text})
        except Exception:
            continue

    if not outlines:
        return {"error": "No outline found in source document", "slides_created": 0}

    slides_created = 0
    for item in outlines:
        if item["level"] == 1:
            # New section → new slide
            slide = pres.Slides.Add(pres.Slides.Count + 1, 1)
            for i in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes.Item(i)
                if shp.HasTextFrame:
                    shp.TextFrame.TextRange.Text = item["text"]
                    shp.TextFrame.TextRange.Font.Size = 28
                    break
            slides_created += 1
            current_slide = slide
        elif item["level"] == 2 and current_slide is not None:
            # Sub-point → add to current slide body
            for i in range(1, current_slide.Shapes.Count + 1):
                shp = current_slide.Shapes.Item(i)
                if shp.HasTextFrame and shp.Name.lower().find("body") >= 0:
                    existing = com_property(shp.TextFrame.TextRange, "Text", "")
                    new_text = existing + "\n• " + item["text"] if existing else "• " + item["text"]
                    shp.TextFrame.TextRange.Text = new_text
                    break

    return {
        "action": "word_outline_to_ppt",
        "slides_created": slides_created,
        "outline_items": len(outlines),
        "ppt_name": pres.Name,
    }
