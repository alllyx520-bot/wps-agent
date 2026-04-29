# -*- coding: utf-8 -*-
"""
Cross-application data migration: Word↔Excel, Word→PPT
"""
from typing import Any, Dict
from .app import get_doc as get_word_doc
from .excel_app import ExcelApplication
from .docspace import get_word_doc_by_id, get_excel_wb_by_id
from .utils import com_property, com_set, col_letter as _col_letter, parse_cell as _parse_cell
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
    col_letter_target, start_row = _parse_cell(target_cell)
    end_letter = _col_letter(ord(col_letter_target) - 64 + cols - 1)
    end_cell = f"{end_letter}{start_row + rows - 1}"
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

    if isinstance(data, (int, float, str)):
        data = ((data,),)
    elif not isinstance(data, tuple):
        data = ((data,),)
    elif not isinstance(data[0], tuple):
        data = (data,)

    rows = len(data)
    cols = len(data[0]) if rows > 0 else 0
    if rows == 0 or cols == 0:
        return {"error": "Range has no data", "rows": rows, "cols": cols}

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
    current_slide = None
    # Add title slide first
    if outlines:
        title_slide = pres.Slides.Add(1, 1)
        for i in range(1, title_slide.Shapes.Count + 1):
            shp = title_slide.Shapes.Item(i)
            if shp.HasTextFrame and shp.Name.lower().find("title") >= 0:
                shp.TextFrame.TextRange.Text = outlines[0]["text"]
                shp.TextFrame.TextRange.Font.Size = 32
                shp.TextFrame.TextRange.Font.Bold = True
                shp.TextFrame.TextRange.Font.NameFarEast = "黑体"
                break
        slides_created = 1
        current_slide = title_slide

    for item in outlines[1:] if len(outlines) > 1 else []:
        if item["level"] == 1:
            slide = pres.Slides.Add(pres.Slides.Count + 1, 2)
            for i in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes.Item(i)
                if shp.HasTextFrame:
                    if shp.Name.lower().find("title") >= 0:
                        shp.TextFrame.TextRange.Text = item["text"]
                        shp.TextFrame.TextRange.Font.Size = 28
                        shp.TextFrame.TextRange.Font.NameFarEast = "黑体"
                        shp.TextFrame.TextRange.Font.Bold = True
                    elif shp.Name.lower().find("body") >= 0 or shp.Name.lower().find("content") >= 0:
                        shp.TextFrame.TextRange.Text = ""
            slides_created += 1
            current_slide = slide
        elif item["level"] == 2 and current_slide is not None:
            for i in range(1, current_slide.Shapes.Count + 1):
                shp = current_slide.Shapes.Item(i)
                if shp.HasTextFrame and (shp.Name.lower().find("body") >= 0 or shp.Name.lower().find("content") >= 0):
                    existing = com_property(shp.TextFrame.TextRange, "Text", "")
                    bullet = "• " + item["text"]
                    new_text = existing + "\r" + bullet if existing else bullet
                    shp.TextFrame.TextRange.Text = new_text
                    shp.TextFrame.TextRange.Font.Size = 18
                    shp.TextFrame.TextRange.Font.NameFarEast = "宋体"
                    break

    return {
        "action": "word_outline_to_ppt",
        "slides_created": slides_created,
        "outline_items": len(outlines),
        "ppt_name": pres.Name,
    }
