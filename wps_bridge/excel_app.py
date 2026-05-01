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
        for progid in ("Ket.Application", "Excel.Application", "ET.Application"):
            try:
                cls._app = win32com.client.GetObject(None, progid)
                break
            except Exception:
                continue
        if cls._app is None:
            raise RuntimeError("Excel/WPS Spreadsheet is not running. Please open WPS Excel first.")
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
    if wb is None:
        return {"error": "No workbook open"}
    if filepath:
        wb.SaveAs(filepath)
    else:
        wb.Save()
    return {"name": wb.Name, "saved": True}


def wb_close(save_changes: bool = False) -> Dict:
    wb = _excel.active_workbook
    if wb is None:
        return {"error": "No workbook open"}
    name = wb.Name
    wb.Close(save_changes)
    return {"closed": name}


def _ensure_wb():
    """Returns (wb, error). If error is not None, caller should return it."""
    wb = _excel.active_workbook
    if wb is None:
        return None, {"error": "No workbook open"}
    return wb, None


def _resolve_sheet(wb, sheet_name=None):
    if wb is None:
        return None
    return wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet


def _parse_range(start, end=None):
    """Parse cell/range references into (col_letter, row_num, cols, rows)."""
    from .utils import col_letter, parse_cell
    col, row = parse_cell(start)
    if end:
        end_col, end_row = parse_cell(end)
        return col, row, ord(end_col) - ord(col) + 1, end_row - row + 1
    return col, row, 1, 1


def sheet_list() -> List[str]:
    wb, err = _ensure_wb()
    if err:
        return err
    sheets = []
    for i in range(1, wb.Worksheets.Count + 1):
        sheets.append(wb.Worksheets.Item(i).Name)
    return sheets


def sheet_activate(name: str) -> Dict:
    wb, err = _ensure_wb()
    if err:
        return err
    ws = wb.Worksheets(name)
    ws.Activate()
    return {"active_sheet": name}


def sheet_add(name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err:
        return err
    ws = wb.Worksheets.Add()
    if name:
        ws.Name = name
    return {"name": ws.Name}


def cell_read(cell_ref: str, sheet_name: Optional[str] = None) -> Any:
    wb, err = _ensure_wb()
    if err:
        return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    value = ws.Range(cell_ref).Value
    return {"cell": cell_ref, "value": value}


def cell_write(cell_ref: str, value: Any, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err:
        return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(cell_ref).Value = value
    return {"cell": cell_ref, "value": value, "written": True}


def range_read(start: str, end: str, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err:
        return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(start, end)
    data = rng.Value
    if data is None:
        return {"range": f"{start}:{end}", "data": []}
    if not isinstance(data, tuple):
        data = [[data]]
    return {"range": f"{start}:{end}", "rows": len(data), "cols": len(data[0]) if data else 0, "data": data}


def range_write(start: str, data: List[List[Any]], sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err:
        return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rows = len(data)
    cols = len(data[0]) if data else 1
    from .utils import col_letter, parse_cell
    end_col = col_letter(ord(start[0].upper()) - 64 + cols - 1)
    _, start_row = parse_cell(start)
    end_cell = f"{end_col}{start_row + rows - 1}"
    rng = ws.Range(start, end_cell)
    rng.Value = data
    return {"range": f"{start}:{end_cell}", "rows": rows, "cols": cols, "written": True}


def font_set(cell_ref: str, name: Optional[str] = None, size: Optional[float] = None,
             bold: Optional[bool] = None, italic: Optional[bool] = None,
             color: Optional[int] = None, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err:
        return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(cell_ref)
    props = {"Name": name, "Size": size, "Bold": bold, "Italic": italic, "ColorIndex": color}
    failed = com_set_batch(rng.Font, props)
    return {"cell": cell_ref, "font_set": True, "failed": failed}


def interior_set(cell_ref: str, color: Optional[int] = None,
                 sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err:
        return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(cell_ref)
    if color is not None:
        rng.Interior.ColorIndex = color
    return {"cell": cell_ref, "interior_color": color}


def borders_set(cell_ref: str, style: int = 1, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(cell_ref)
    for border_pos in [7, 8, 9, 10]:  # left, top, bottom, right
        try:
            rng.Borders(border_pos).LineStyle = style
        except Exception:
            continue
    return {"cell": cell_ref, "border_style": style}


def column_width(col: str, width: float, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(f"{col}:{col}").ColumnWidth = width
    return {"column": col, "width": width}


def auto_fit_range(start: str, end: str, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(start, end).Columns.AutoFit()
    return {"auto_fit": f"{start}:{end}"}


def merge_cells(start: str, end: str, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(start, end).Merge()
    return {"merged": f"{start}:{end}"}


def formula_set(cell_ref: str, formula: str, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Range(cell_ref).Formula = formula
    return {"cell": cell_ref, "formula": formula}


def chart_add(chart_type: int = 4, left: int = 100, top: int = 100,
              width: int = 400, height: int = 300, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    chart_obj = ws.ChartObjects().Add(left, top, width, height)
    chart_obj.Chart.ChartType = chart_type
    return {"chart_index": ws.ChartObjects().Count, "type": chart_type}


# ====== Phase 6: Advanced Excel Features ======

def sort_range(start: str, end: str, key_col: int, order: int = 1,
               sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(start, end)
    rng.Sort(Key1=rng.Columns(key_col), Order1=order)
    return {"range": f"{start}:{end}", "sorted_by_col": key_col, "order": "asc" if order == 1 else "desc"}


def auto_filter(start: str, end: str, field: int = 1,
                criteria: str = "", sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(start, end)
    if criteria:
        rng.AutoFilter(Field=field, Criteria1=criteria)
    else:
        rng.AutoFilter()
    return {"range": f"{start}:{end}", "field": field, "criteria": criteria, "filtered": True}


def remove_filter(sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.AutoFilterMode = False
    return {"filter_removed": True}


def conditional_format(start: str, end: str, rule_type: int = 1,
                       formula: str = "", color: int = 3,
                       sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(start, end)
    fc = rng.FormatConditions.Add(Type=rule_type, Formula1=formula)
    fc.Interior.ColorIndex = color
    return {"range": f"{start}:{end}", "rule_type": rule_type, "color": color}


def sheet_copy(name: str, before: Optional[str] = None,
               after: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(name)
    if before:
        ws.Copy(Before=wb.Worksheets(before))
    elif after:
        ws.Copy(After=wb.Worksheets(after))
    else:
        ws.Copy()
    return {"copied": name}


def sheet_delete(name: str) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    wb.Worksheets(name).Delete()
    return {"deleted": name}


def sheet_move(name: str, before: Optional[str] = None,
               after: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(name)
    if before:
        ws.Move(Before=wb.Worksheets(before))
    elif after:
        ws.Move(After=wb.Worksheets(after))
    return {"moved": name}


def chart_set_source(chart_index: int, range_start: str, range_end: str,
                     sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    chart = ws.ChartObjects().Item(chart_index).Chart
    src = ws.Range(range_start, range_end)
    chart.SetSourceData(src)
    return {"chart": chart_index, "source": f"{range_start}:{range_end}"}


def chart_set_title(chart_index: int, title: str,
                    sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    chart = ws.ChartObjects().Item(chart_index).Chart
    if not chart.HasTitle:
        chart.HasTitle = True
    chart.ChartTitle.Text = title
    return {"chart": chart_index, "title": title}


def freeze_panes(cell_ref: str, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    ws.Activate()
    ws.Range(cell_ref).Select()
    wb.ActiveWindow.FreezePanes = True
    return {"frozen_at": cell_ref}


def get_used_range(sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    used = ws.UsedRange
    return {
        "address": com_property(used, "Address", ""),
        "rows": com_property(used.Rows, "Count", 0),
        "cols": com_property(used.Columns, "Count", 0),
    }


def insert_rows(row: int, count: int = 1, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    for _ in range(count):
        ws.Rows(row).Insert()
    return {"inserted": count, "at_row": row}


def delete_rows(row: int, count: int = 1, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    for _ in range(count):
        ws.Rows(row).Delete()
    return {"deleted": count, "from_row": row}


def add_cell_comment(cell_ref: str, text: str, sheet_name: Optional[str] = None) -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    rng = ws.Range(cell_ref)
    rng.AddComment(text)
    return {"cell": cell_ref, "comment": text}


def import_csv(filepath: str, delimiter: str = ",", has_header: bool = True, sheet_name: Optional[str] = None) -> Dict:
    import csv
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    with open(filepath, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=delimiter)
        rows = list(reader)
    if not rows:
        return {"error": "CSV file is empty"}
    start_row = _find_next_empty_row(ws)
    for i, row_data in enumerate(rows):
        for j, val in enumerate(row_data):
            ws.Cells(start_row + i, j + 1).Value = val
    col_letter = chr(64 + len(rows[0])) if len(rows[0]) <= 26 else "Z"
    end_ref = f"{col_letter}{start_row + len(rows) - 1}"
    return {"imported_from": filepath, "rows": len(rows), "range": f"A{start_row}:{end_ref}"}


def export_csv(filepath: str, start: str, end: str, delimiter: str = ",", sheet_name: Optional[str] = None) -> Dict:
    import csv
    data = range_read(start, end, sheet_name)
    if isinstance(data, dict) and "error" in data:
        return data
    rows_data = data.get("data", [])
    with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=delimiter)
        for row in rows_data:
            writer.writerow(row if isinstance(row, (list, tuple)) else [row])
    return {"exported_to": filepath, "rows": len(rows_data)}


def _find_next_empty_row(ws) -> int:
    try:
        used = ws.UsedRange
        rows = used.Rows.Count
        return rows + 1
    except Exception:
        return 1


def validate_formulas(sheet_name: Optional[str] = None) -> Dict:
    ERRORS = ["#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!"]
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    total_formulas = 0
    errors_found = {}
    try:
        used = ws.UsedRange
        for row in range(1, used.Rows.Count + 1):
            for col in range(1, used.Columns.Count + 1):
                try:
                    cell = used.Cells(row, col)
                    formula = com_property(cell, "Formula", "")
                    text = str(com_property(cell, "Text", ""))
                    if formula and formula.startswith("="):
                        total_formulas += 1
                    for err in ERRORS:
                        if err in text:
                            addr = cell.Address.replace("$", "")
                            if err not in errors_found:
                                errors_found[err] = {"count": 0, "locations": []}
                            errors_found[err]["count"] += 1
                            errors_found[err]["locations"].append(addr)
                            break
                except Exception:
                    continue
    except Exception as e:
        return {"error": str(e)}
    status = "errors_found" if errors_found else "success"
    return {"status": status, "total_errors": sum(v["count"] for v in errors_found.values()), "total_formulas": total_formulas, "error_summary": errors_found if errors_found else None}


def recalc_formulas() -> Dict:
    wb, err = _ensure_wb()
    if err: return err
    try:
        wb.Application.Calculate()
        return {"recalculated": True}
    except Exception as e:
        return {"error": str(e)}


def apply_financial_colors(sheet_name: Optional[str] = None) -> Dict:
    """Apply financial model color conventions to the worksheet.
    Blue=hardcoded inputs, Black=formulas, Green=internal refs, Red=external refs."""
    wb, err = _ensure_wb()
    if err: return err
    ws = wb.Worksheets(sheet_name) if sheet_name else _excel.active_sheet
    stats = {"hardcoded": 0, "formula": 0, "internal_ref": 0, "external_ref": 0}
    try:
        used = ws.UsedRange
        for row in range(1, used.Rows.Count + 1):
            for col in range(1, used.Columns.Count + 1):
                try:
                    cell = used.Cells(row, col)
                    formula = str(com_property(cell, "Formula", ""))
                    if formula.startswith("="):
                        if ".xls" in formula.lower():
                            cell.Font.Color = 0xFF0000  # Red: external
                            stats["external_ref"] += 1
                        elif "!" in formula:
                            cell.Font.Color = 0x008000  # Green: internal
                            stats["internal_ref"] += 1
                        else:
                            cell.Font.Color = 0x000000  # Black: local formula
                            stats["formula"] += 1
                    else:
                        val = com_property(cell, "Value", None)
                        if val is not None and val != "":
                            cell.Font.Color = 0x0000FF  # Blue: hardcoded
                            stats["hardcoded"] += 1
                except Exception:
                    continue
    except Exception as e:
        return {"error": str(e)}
    return {"financial_colors_applied": True, "stats": stats}
