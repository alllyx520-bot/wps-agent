# -*- coding: utf-8 -*-
import win32com.client
import pythoncom
from typing import Any, Optional, Dict, List
from .utils import co_init, com_property, com_set, com_set_batch


class ExcelApplication:
    _app: Any = None
    _visible: bool = True

    @classmethod
    def get_instance(cls, visible: bool = True) -> Any:
        if cls._app is not None:
            try:
                cls._app.Workbooks.Count
                return cls._app
            except Exception:
                cls._app = None
        co_init()
        cls._visible = visible
        try:
            cls._app = win32com.client.GetObject(None, "Ket.Application")
        except Exception:
            cls._app = win32com.client.Dispatch("Ket.Application")
        com_set(cls._app, "Visible", visible)
        return cls._app

    @property
    def app(self) -> Any:
        return ExcelApplication.get_instance(self._visible)

    @property
    def active_workbook(self) -> Any:
        try:
            return self.app.ActiveWorkbook
        except Exception:
            return None

    @property
    def active_sheet(self) -> Any:
        try:
            return self.app.ActiveSheet
        except Exception:
            return None

    def list_workbooks(self) -> List[Dict]:
        result = []
        try:
            count = self.app.Workbooks.Count
        except Exception:
            return result
        for i in range(1, count + 1):
            try:
                wb = self.app.Workbooks.Item(i)
                result.append({
                    "index": i,
                    "name": com_property(wb, "Name", ""),
                    "full_name": com_property(wb, "FullName", ""),
                    "sheets": com_property(wb.Worksheets, "Count", 0),
                    "saved": com_property(wb, "Saved", False),
                })
            except Exception:
                continue
        return result

    def quit(self):
        try:
            self.app.Quit()
        except Exception:
            pass
        ExcelApplication._app = None


_excel = ExcelApplication()


def wb_create() -> Dict:
    wb = _excel.app.Workbooks.Add()
    return {"name": wb.Name, "sheets": wb.Worksheets.Count}


def wb_open(filepath: str) -> Dict:
    wb = _excel.app.Workbooks.Open(filepath)
    return {"name": wb.Name, "sheets": wb.Worksheets.Count}


def wb_list() -> List[Dict]:
    return _excel.list_workbooks()


def wb_save(filepath: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    if filepath:
        wb.SaveAs(filepath)
    else:
        wb.Save()
    return {"name": wb.Name, "saved": True}


def wb_close(save_changes: bool = False) -> Dict:
    wb = _excel.active_workbook
    name = wb.Name
    wb.Close(save_changes)
    return {"closed": name}


def sheet_list() -> List[str]:
    wb = _excel.active_workbook
    sheets = []
    for i in range(1, wb.Worksheets.Count + 1):
        sheets.append(wb.Worksheets.Item(i).Name)
    return sheets


def sheet_activate(name: str) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(name)
    ws.Activate()
    return {"active_sheet": name}


def sheet_add(name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets.Add()
    if name:
        ws.Name = name
    return {"name": ws.Name}


def cell_read(cell_ref: str, sheet_name: Optional[str] = None) -> Any:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    value = ws.Range(cell_ref).Value
    return {"cell": cell_ref, "value": value}


def cell_write(cell_ref: str, value: Any, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(cell_ref).Value = value
    return {"cell": cell_ref, "value": value, "written": True}


def range_read(start: str, end: str, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(start, end)
    data = rng.Value
    if data is None:
        return {"range": f"{start}:{end}", "data": []}
    if not isinstance(data, tuple):
        data = [[data]]
    return {"range": f"{start}:{end}", "rows": len(data), "cols": len(data[0]) if data else 0, "data": data}


def range_write(start: str, data: List[List[Any]], sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rows = len(data)
    cols = len(data[0]) if data else 1
    end_col = chr(ord('A') + cols - 1)
    end_row = int(start[1:]) + rows - 1 if len(start) > 1 else rows
    end_cell = f"{end_col}{end_row}"
    rng = ws.Range(start, end_cell)
    rng.Value = data
    return {"range": f"{start}:{end_cell}", "rows": rows, "cols": cols, "written": True}


def font_set(cell_ref: str, name: Optional[str] = None, size: Optional[float] = None,
             bold: Optional[bool] = None, italic: Optional[bool] = None,
             color: Optional[int] = None, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(cell_ref)
    props = {"Name": name, "Size": size, "Bold": bold, "Italic": italic, "ColorIndex": color}
    failed = com_set_batch(rng.Font, props)
    return {"cell": cell_ref, "font_set": True, "failed": failed}


def interior_set(cell_ref: str, color: Optional[int] = None,
                 sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(cell_ref)
    if color is not None:
        rng.Interior.ColorIndex = color
    return {"cell": cell_ref, "interior_color": color}


def borders_set(cell_ref: str, style: int = 1, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(cell_ref)
    for border_pos in [7, 8, 9, 10]:  # left, top, bottom, right
        try:
            rng.Borders(border_pos).LineStyle = style
        except Exception:
            continue
    return {"cell": cell_ref, "border_style": style}


def column_width(col: str, width: float, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(f"{col}:{col}").ColumnWidth = width
    return {"column": col, "width": width}


def auto_fit_range(start: str, end: str, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(start, end).Columns.AutoFit()
    return {"auto_fit": f"{start}:{end}"}


def merge_cells(start: str, end: str, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(start, end).Merge()
    return {"merged": f"{start}:{end}"}


def formula_set(cell_ref: str, formula: str, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(cell_ref).Formula = formula
    return {"cell": cell_ref, "formula": formula}


def chart_add(chart_type: int = 4, left: int = 100, top: int = 100,
              width: int = 400, height: int = 300, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    chart_obj = ws.ChartObjects().Add(left, top, width, height)
    chart_obj.Chart.ChartType = chart_type
    return {"chart_index": ws.ChartObjects().Count, "type": chart_type}


# ====== Phase 6: Advanced Excel Features ======

def sort_range(start: str, end: str, key_col: int, order: int = 1,
               sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(start, end)
    rng.Sort(Key1=rng.Columns(key_col), Order1=order)
    return {"range": f"{start}:{end}", "sorted_by_col": key_col, "order": "asc" if order == 1 else "desc"}


def auto_filter(start: str, end: str, field: int = 1,
                criteria: str = "", sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(start, end)
    if criteria:
        rng.AutoFilter(Field=field, Criteria1=criteria)
    else:
        rng.AutoFilter()
    return {"range": f"{start}:{end}", "field": field, "criteria": criteria, "filtered": True}


def remove_filter(sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.AutoFilterMode = False
    return {"filter_removed": True}


def conditional_format(start: str, end: str, rule_type: int = 1,
                       formula: str = "", color: int = 3,
                       sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(start, end)
    fc = rng.FormatConditions.Add(Type=rule_type, Formula1=formula)
    fc.Interior.ColorIndex = color
    return {"range": f"{start}:{end}", "rule_type": rule_type, "color": color}


def sheet_copy(name: str, before: Optional[str] = None,
               after: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(name)
    if before:
        ws.Copy(Before=wb.Worksheets(before))
    elif after:
        ws.Copy(After=wb.Worksheets(after))
    else:
        ws.Copy()
    return {"copied": name}


def sheet_delete(name: str) -> Dict:
    wb = _excel.active_workbook
    wb.Worksheets(name).Delete()
    return {"deleted": name}


def sheet_move(name: str, before: Optional[str] = None,
               after: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(name)
    if before:
        ws.Move(Before=wb.Worksheets(before))
    elif after:
        ws.Move(After=wb.Worksheets(after))
    return {"moved": name}


def chart_set_source(chart_index: int, range_start: str, range_end: str,
                     sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    chart = ws.ChartObjects().Item(chart_index).Chart
    src = ws.Range(range_start, range_end)
    chart.SetSourceData(src)
    return {"chart": chart_index, "source": f"{range_start}:{range_end}"}


def chart_set_title(chart_index: int, title: str,
                    sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    chart = ws.ChartObjects().Item(chart_index).Chart
    if not chart.HasTitle:
        chart.HasTitle = True
    chart.ChartTitle.Text = title
    return {"chart": chart_index, "title": title}


def freeze_panes(cell_ref: str, sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Activate()
    ws.Range(cell_ref).Select()
    wb.ActiveWindow.FreezePanes = True
    return {"frozen_at": cell_ref}


def get_used_range(sheet_name: Optional[str] = None) -> Dict:
    wb = _excel.active_workbook
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    used = ws.UsedRange
    return {
        "address": com_property(used, "Address", ""),
        "rows": com_property(used.Rows, "Count", 0),
        "cols": com_property(used.Columns, "Count", 0),
    }
