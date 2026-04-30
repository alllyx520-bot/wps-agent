# -*- coding: utf-8 -*-
"""
PPT COM bridge for WPS Presentation (Kwpp.Application)
"""
import win32com.client
from typing import Any, Optional, Dict, List
from .utils import co_init, com_property, com_set, com_set_batch


class PPTApplication:
    _app: Any = None
    _visible: bool = True

    @classmethod
    def get_instance(cls, visible: bool = True) -> Any:
        if cls._app is not None:
            try:
                cls._app.Presentations.Count
                return cls._app
            except Exception:
                cls._app = None
        co_init()
        cls._visible = visible
        for progid in ("Kwpp.Application", "PowerPoint.Application", "WPP.Application"):
            try:
                cls._app = win32com.client.GetObject(None, progid)
                break
            except Exception:
                continue
        if cls._app is None:
            raise RuntimeError("PowerPoint/WPS Presentation is not running. Please open WPS PPT first.")
        com_set(cls._app, "Visible", visible)
        return cls._app

    @property
    def app(self) -> Any:
        return PPTApplication.get_instance(self._visible)

    @property
    def active_presentation(self) -> Any:
        try:
            return self.app.ActivePresentation
        except Exception:
            return None

    @property
    def active_slide(self) -> Any:
        try:
            view = self.app.ActiveWindow.View
            return view.Slide if view else None
        except Exception:
            return None

    def list_presentations(self) -> List[Dict]:
        result = []
        try:
            count = self.app.Presentations.Count
        except Exception:
            return result
        for i in range(1, count + 1):
            try:
                pres = self.app.Presentations.Item(i)
                result.append({
                    "index": i,
                    "name": com_property(pres, "Name", ""),
                    "full_name": com_property(pres, "FullName", ""),
                    "slides": com_property(pres.Slides, "Count", 0),
                    "saved": com_property(pres, "Saved", False),
                })
            except Exception:
                continue
        return result

    def quit(self):
        try:
            self.app.Quit()
        except Exception:
            pass
        PPTApplication._app = None


_ppt = PPTApplication()


def pres_create() -> Dict:
    pres = _ppt.app.Presentations.Add()
    return {"name": pres.Name, "slides": pres.Slides.Count}


def pres_open(filepath: str) -> Dict:
    pres = _ppt.app.Presentations.Open(filepath)
    return {"name": pres.Name, "slides": pres.Slides.Count}


def pres_list() -> List[Dict]:
    return _ppt.list_presentations()


def pres_save(filepath: Optional[str] = None) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    if filepath:
        pres.SaveAs(filepath)
    else:
        pres.Save()
    return {"name": pres.Name, "saved": True}


def pres_close(save_changes: bool = False) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    name = pres.Name
    pres.Close(save_changes)
    return {"closed": name}


def slide_count() -> int:
    pres = _ppt.active_presentation
    if pres is None:
        return 0
    return pres.Slides.Count


def slide_info(slide_index: int) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    shapes = []
    for i in range(1, slide.Shapes.Count + 1):
        try:
            shp = slide.Shapes.Item(i)
            shape_info = {
                "index": i,
                "type": com_property(shp, "Type", 0),
                "name": com_property(shp, "Name", ""),
                "left": com_property(shp, "Left", 0),
                "top": com_property(shp, "Top", 0),
                "width": com_property(shp, "Width", 0),
                "height": com_property(shp, "Height", 0),
            }
            if shp.HasTextFrame:
                shape_info["text"] = com_property(shp.TextFrame.TextRange, "Text", "")[:100]
            if shp.HasTable:
                tbl = shp.Table
                shape_info["table"] = f"{tbl.Rows.Count}x{tbl.Columns.Count}"
            shapes.append(shape_info)
        except Exception:
            continue
    return {
        "index": slide_index,
        "slide_id": com_property(slide, "SlideID", 0),
        "layout_name": com_property(slide.Layout, "Name", ""),
        "shapes_count": slide.Shapes.Count,
        "shapes": shapes,
        "notes": _get_notes(slide),
    }


def _get_notes(slide) -> str:
    try:
        notes_page = slide.NotesPage
        for i in range(1, notes_page.Shapes.Count + 1):
            shp = notes_page.Shapes.Item(i)
            if shp.HasTextFrame:
                return com_property(shp.TextFrame.TextRange, "Text", "")
    except Exception:
        pass
    return ""


def add_slide(layout_index: int = 1) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Add(pres.Slides.Count + 1, layout_index)
    return {"slide_index": slide.SlideIndex, "layout": layout_index}


def delete_slide(slide_index: int) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    pres.Slides.Item(slide_index).Delete()
    return {"deleted": slide_index}


def set_title(slide_index: int, text: str) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    # Find title placeholder by type (ppPlaceholderTitle = 1), then fallback to name
    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes.Item(i)
        if shp.HasTextFrame:
            try:
                if shp.PlaceholderFormat.Type == 1:
                    shp.TextFrame.TextRange.Text = text
                    return {"slide": slide_index, "title": text}
            except Exception:
                pass
    # Fallback: search by name
    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes.Item(i)
        if shp.HasTextFrame and shp.Name.lower().find("title") >= 0:
            shp.TextFrame.TextRange.Text = text
            return {"slide": slide_index, "title": text}
    # Create textbox as last resort
    shp = slide.Shapes.AddTextbox(1, 50, 40, 620, 60)
    shp.TextFrame.TextRange.Text = text
    shp.TextFrame.TextRange.Font.Size = 28
    return {"slide": slide_index, "title": text, "added_as_textbox": True}


def set_body(slide_index: int, text: str) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    body_shape = None
    for i in range(1, slide.Shapes.Count + 1):
        shp = slide.Shapes.Item(i)
        if shp.HasTextFrame and shp.Name.lower().find("body") >= 0:
            body_shape = shp
            break
    if not body_shape:
        body_shape = slide.Shapes.AddTextbox(1, 50, 110, 620, 400)
    body_shape.TextFrame.TextRange.Text = text
    body_shape.TextFrame.TextRange.Font.Size = 18
    return {"slide": slide_index, "body_set": True}


def add_textbox(slide_index: int, text: str, left: int = 50, top: int = 100,
                width: int = 620, height: int = 300) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    shp = slide.Shapes.AddTextbox(1, left, top, width, height)
    shp.TextFrame.TextRange.Text = text
    return {"slide": slide_index, "shape_index": slide.Shapes.Count, "text": text[:50]}


def format_text(slide_index: int, shape_index: int, font_name: Optional[str] = None,
                font_size: Optional[float] = None, bold: Optional[bool] = None,
                color: Optional[int] = None) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    shp = slide.Shapes.Item(shape_index)
    if not shp.HasTextFrame:
        return {"error": "Shape has no text frame"}
    tr = shp.TextFrame.TextRange
    props = {"Name": font_name, "Size": font_size, "Bold": bold, "ColorIndex": color}
    failed = com_set_batch(tr.Font, props)
    return {"slide": slide_index, "shape": shape_index, "failed": failed}


def insert_image(slide_index: int, image_path: str, left: int = 100, top: int = 100,
                 width: int = 400, height: int = 300) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    shp = slide.Shapes.AddPicture(image_path, 0, -1, left, top, width, height)
    return {"slide": slide_index, "shape_index": slide.Shapes.Count, "image": image_path}


def insert_table(slide_index: int, rows: int, cols: int, left: int = 50,
                 top: int = 150, width: int = 600, height: int = 300) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    shp = slide.Shapes.AddTable(rows, cols, left, top, width, height)
    return {"slide": slide_index, "shape_index": slide.Shapes.Count, "rows": rows, "cols": cols}


def fill_cell(slide_index: int, table_index: int, row: int, col: int, text: str) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    shp = slide.Shapes.Item(table_index)
    if not shp.HasTable:
        return {"error": "Shape is not a table"}
    shp.Table.Cell(row, col).Shape.TextFrame.TextRange.Text = text
    return {"slide": slide_index, "table": table_index, "row": row, "col": col, "text": text}


def apply_theme(theme_name: str) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    try:
        pres.ApplyTemplate(theme_name)
        return {"theme": theme_name}
    except Exception as e:
        return {"error": str(e)}


def add_notes(slide_index: int, text: str) -> Dict:
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    try:
        notes_page = slide.NotesPage
        for i in range(1, notes_page.Shapes.Count + 1):
            shp = notes_page.Shapes.Item(i)
            if shp.HasTextFrame:
                existing = com_property(shp.TextFrame.TextRange, "Text", "")
                shp.TextFrame.TextRange.Text = existing + "\n" + text if existing else text
                return {"slide": slide_index, "notes": text}
    except Exception:
        pass
    return {"error": "Could not set notes"}


def set_slide_format(slide_index: int, background_color: Optional[int] = None) -> Dict:
    slide = _ppt.active_presentation.Slides.Item(slide_index)
    if background_color is not None:
        try:
            slide.Background.Fill.ForeColor.RGB = background_color
            slide.Background.Fill.Visible = True
        except Exception as e:
            return {"error": str(e)}
    return {"slide": slide_index, "background_set": True}


def add_shape(slide_index: int, shape_type: str, left: float = 100, top: float = 100,
              width: float = 300, height: float = 200, fill_color: Optional[str] = None,
              line_color: Optional[str] = None, line_width: float = 1) -> Dict:
    """Add a shape to a slide. type: rectangle/oval/line/rounded_rectangle"""
    SHAPES = {"rectangle": 1, "oval": 9, "line": 10, "rounded_rectangle": 5}
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    st = SHAPES.get(shape_type, 1)
    slide = pres.Slides.Item(slide_index)
    shp = slide.Shapes.AddShape(st, left, top, width, height)
    if fill_color:
        try:
            shp.Fill.ForeColor.RGB = int(fill_color, 16)
            shp.Fill.Visible = True
        except Exception:
            pass
    if line_color:
        try:
            shp.Line.ForeColor.RGB = int(line_color, 16)
            shp.Line.Weight = line_width
            shp.Line.Visible = True
        except Exception:
            pass
    return {"slide": slide_index, "shape_index": slide.Shapes.Count, "type": shape_type}


def reorder_slides(slide_order: List[int]) -> Dict:
    """Reorder slides by their new index order. slide_order = [3, 1, 2] means slide 3 -> pos 1, etc."""
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    total = pres.Slides.Count
    if len(slide_order) != total:
        return {"error": f"slide_order length ({len(slide_order)}) must match total slides ({total})"}
    moved = 0
    for target_pos, slide_idx in enumerate(slide_order, 1):
        if slide_idx != target_pos:
            try:
                pres.Slides.Item(slide_idx).MoveTo(target_pos)
                moved += 1
            except Exception:
                continue
    return {"reordered": total, "moved": moved}


def set_slide_background(slide_index: int, color_hex: Optional[str] = None,
                         image_path: Optional[str] = None, transparency: int = 0) -> Dict:
    """Set slide background to solid color or image. color_hex: 6-char hex like '1E2761'"""
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    slide = pres.Slides.Item(slide_index)
    bg = slide.Background
    if color_hex:
        try:
            bg.Fill.ForeColor.RGB = int(color_hex, 16)
            if transparency:
                bg.Fill.ForeColor.Brightness = 1.0 - (transparency / 100.0)
            bg.Fill.Visible = True
        except Exception as e:
            return {"error": str(e)}
    elif image_path:
        try:
            bg.Fill.UserPicture(image_path)
            bg.Fill.Visible = True
        except Exception as e:
            return {"error": str(e)}
    return {"slide": slide_index, "background": "color" if color_hex else "image", "value": color_hex or image_path}


def add_chart_modern(slide_index: int, chart_type: str, categories: List[str], values: List[float],
                     series_name: str = "", title: str = "", left: float = 50, top: float = 100,
                     width: float = 600, height: float = 350) -> Dict:
    """Add a chart with modern styling. chart_type: bar/line/pie/column"""
    CHART_TYPES = {"bar": 2, "line": 4, "pie": 5, "column": 1}
    pres = _ppt.active_presentation
    if pres is None:
        return {"error": "No presentation open"}
    ct = CHART_TYPES.get(chart_type, 1)
    slide = pres.Slides.Item(slide_index)
    shp = slide.Shapes.AddChart(ct, left, top, width, height)
    chart = shp.Chart
    chart.ChartData.Activate()
    try:
        wb = chart.ChartData.Workbook
        ws = wb.Worksheets(1)
        # Clear default data
        ws.Cells.Clear()
        # Write categories (row 1, starting col 2)
        ws.Cells(1, 1).Value = ""
        for i, cat in enumerate(categories):
            ws.Cells(1, i + 2).Value = cat
        # Write values (row 2)
        ws.Cells(2, 1).Value = series_name or "Series 1"
        for i, val in enumerate(values):
            ws.Cells(2, i + 2).Value = val
        chart.SetSourceData(ws.Range(ws.Cells(1, 1), ws.Cells(2, len(categories) + 1)))
        if title:
            chart.HasTitle = True
            chart.ChartTitle.Text = title
        # Modern styling
        try:
            chart.ChartArea.RoundedCorners = True
        except Exception:
            pass
        try:
            chart.HasLegend = len(values) > 1
        except Exception:
            pass
    except Exception as e:
        return {"error": str(e), "chart_added": True}
    return {"slide": slide_index, "shape_index": slide.Shapes.Count, "chart_type": chart_type, "categories": len(categories), "values": len(values)}
