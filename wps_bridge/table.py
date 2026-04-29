# -*- coding: utf-8 -*-
from typing import Any, Optional, Dict, List
from .app import get_app, get_doc
from .utils import com_property, com_set, com_set_batch, WDALIGNMENT


def table_count(doc_index=None): return com_property(get_doc(doc_index).Tables, "Count", 0)

def batch_read(table_indices: list, doc_index=None):
    if not table_indices:
        return []
    return [{"table_index": i, "data": table_read(i, doc_index)} for i in table_indices]

def table_info(table_index, doc_index=None):
    tbl = get_doc(doc_index).Tables.Item(table_index)
    return {"index": table_index, "rows": com_property(tbl.Rows, "Count", 0), "columns": com_property(tbl.Columns, "Count", 0)}

def table_read(table_index, doc_index=None):
    tbl = get_doc(doc_index).Tables.Item(table_index)
    rows, cols = tbl.Rows.Count, tbl.Columns.Count
    data = []
    for r in range(1, rows + 1):
        row_data = []
        for c in range(1, cols + 1):
            try: row_data.append(tbl.Cell(r, c).Range.Text.replace("\r\x07", "").strip())
            except: row_data.append("")
        data.append(row_data)
    return {"rows": rows, "columns": cols, "data": data}

def table_create(rows, cols, position=None, doc_index=None):
    doc = get_doc(doc_index)
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1) if position == "end" else get_app().Selection.Range
    doc.Tables.Add(rng, rows, cols).AutoFitBehavior(2)
    return {"table_index": doc.Tables.Count, "rows": rows, "columns": cols}

def table_delete(table_index, doc_index=None):
    get_doc(doc_index).Tables.Item(table_index).Delete()
    return {"deleted": table_index}

def set_cell_text(table_index, row, col, text, doc_index=None):
    get_doc(doc_index).Tables.Item(table_index).Cell(row, col).Range.Text = text
    return {"table": table_index, "row": row, "col": col, "text": text}

def format_cell(table_index, row, col, font_name=None, font_size=None, bold=None, align=None, shading_color=None, doc_index=None):
    cell = get_doc(doc_index).Tables.Item(table_index).Cell(row, col)
    r = cell.Range
    if font_name: com_set(r.Font, "Name", font_name)
    if font_size: com_set(r.Font, "Size", font_size)
    if bold is not None: com_set(r.Font, "Bold", bold)
    if align: com_set(r.ParagraphFormat, "Alignment", WDALIGNMENT.get(align, 0))
    if shading_color is not None: com_set(cell.Shading, "BackgroundPatternColorIndex", shading_color)
    return {"table": table_index, "row": row, "col": col, "formatted": True}

def set_header(table_index, row_count=1, doc_index=None):
    tbl = get_doc(doc_index).Tables.Item(table_index)
    for r in range(1, row_count + 1):
        for c in range(1, tbl.Columns.Count + 1):
            tbl.Cell(r, c).Range.Font.Bold = True
    return {"header_rows": row_count}

def format_borders(table_index, inside=None, outside=None, doc_index=None):
    import win32com.client
    tbl = get_doc(doc_index).Tables.Item(table_index)
    wd = win32com.client.constants
    if outside:
        for bn in ["wdBorderTop", "wdBorderBottom", "wdBorderLeft", "wdBorderRight"]:
            try:
                b = tbl.Borders(getattr(wd, bn))
                if b and "style" in outside: com_set(b, "LineStyle", outside["style"])
            except: pass
    if inside:
        for bn in ["wdBorderHorizontal", "wdBorderVertical"]:
            try:
                b = tbl.Borders(getattr(wd, bn))
                if b and "style" in inside: com_set(b, "LineStyle", inside["style"])
            except: pass
    return {"table": table_index, "borders_formatted": True}

def merge_cells(table_index, start_row, start_col, end_row, end_col, doc_index=None):
    tbl = get_doc(doc_index).Tables.Item(table_index)
    tbl.Cell(start_row, start_col).Merge(tbl.Cell(end_row, end_col))
    return {"merged": f"({start_row},{start_col})-({end_row},{end_col})"}

def auto_fit(table_index, behavior=2, doc_index=None):
    get_doc(doc_index).Tables.Item(table_index).AutoFitBehavior(behavior)
    return {"table": table_index, "auto_fit": behavior}

def set_column_width(table_index, col, width, doc_index=None):
    get_doc(doc_index).Tables.Item(table_index).Columns.Item(col).Width = width
    return {"table": table_index, "column": col, "width": width}

def alternate_rows(table_index, color1="FFFFFF", color2="F2F2F2", doc_index=None):
    tbl = get_doc(doc_index).Tables.Item(table_index)
    for r in range(1, tbl.Rows.Count + 1):
        for c in range(1, tbl.Columns.Count + 1):
            try: tbl.Cell(r, c).Shading.BackgroundPatternColorIndex = 16 + (int(color2 if r % 2 == 0 else color1, 16) & 0xFF)
            except: pass
    return {"table": table_index, "note": "Basic support only"}
