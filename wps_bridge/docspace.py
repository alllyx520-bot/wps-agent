# -*- coding: utf-8 -*-
"""
DocSpace: Unified document space for Word/Excel/PPT
Maps doc_id like "word:1", "excel:1" to actual COM objects
"""
from typing import Any, Dict, List, Optional, Tuple
from .app import get_app as get_word_app, list_documents as list_word_docs
from .excel_app import ExcelApplication
from .ppt_app import PPTApplication
from .utils import com_property


def _resolve(doc_id: str) -> Tuple[str, int]:
    parts = doc_id.split(":")
    if len(parts) != 2:
        raise ValueError(f"Invalid doc_id format: {doc_id}. Expected 'type:index' e.g. 'word:1'")
    return parts[0], int(parts[1])


def list_all() -> List[Dict]:
    docs = []
    # Word
    try:
        for d in list_word_docs():
            docs.append({
                "doc_id": f"word:{d['index']}",
                "type": "word",
                "name": d["name"],
                "full_name": d["full_name"],
                "saved": d["saved"],
            })
    except Exception:
        pass
    # Excel
    try:
        xl = ExcelApplication()
        for wb in xl.list_workbooks():
            docs.append({
                "doc_id": f"excel:{wb['index']}",
                "type": "excel",
                "name": wb["name"],
                "full_name": wb["full_name"],
                "saved": wb["saved"],
            })
    except Exception:
        pass
    # PPT
    try:
        ppt = PPTApplication()
        for pres in ppt.list_presentations():
            docs.append({
                "doc_id": f"ppt:{pres['index']}",
                "type": "ppt",
                "name": pres["name"],
                "full_name": pres["full_name"],
                "saved": pres["saved"],
            })
    except Exception:
        pass
    return docs


def doc_info(doc_id: str) -> Dict:
    app_type, index = _resolve(doc_id)
    if app_type == "word":
        from . import document
        return document.doc_info(index)
    elif app_type == "excel":
        xl = ExcelApplication()
        wb = xl.app.Workbooks.Item(index)
        return {
            "doc_id": doc_id,
            "type": "excel",
            "name": com_property(wb, "Name", ""),
            "full_name": com_property(wb, "FullName", ""),
            "sheets": com_property(wb.Worksheets, "Count", 0),
            "saved": com_property(wb, "Saved", False),
        }
    elif app_type == "ppt":
        ppt = PPTApplication()
        pres = ppt.app.Presentations.Item(index)
        return {
            "doc_id": doc_id,
            "type": "ppt",
            "name": com_property(pres, "Name", ""),
            "full_name": com_property(pres, "FullName", ""),
            "slides": com_property(pres.Slides, "Count", 0),
            "saved": com_property(pres, "Saved", False),
        }
    else:
        return {"error": f"Unsupported app type: {app_type}"}


def activate(doc_id: str) -> Dict:
    app_type, index = _resolve(doc_id)
    if app_type == "word":
        from . import document
        return document.doc_activate(index)
    elif app_type == "excel":
        xl = ExcelApplication()
        wb = xl.app.Workbooks.Item(index)
        wb.Activate()
        return {"active": wb.Name, "doc_id": doc_id}
    elif app_type == "ppt":
        ppt = PPTApplication()
        pres = ppt.app.Presentations.Item(index)
        pres.Activate()
        return {"active": pres.Name, "doc_id": doc_id}
    else:
        return {"error": f"Unsupported app type: {app_type}"}


def close_all() -> Dict:
    closed = []
    errors = []
    # Close Word docs
    try:
        app = get_word_app()
        count = app.Documents.Count
        for i in range(count, 0, -1):
            try:
                doc = app.Documents.Item(i)
                name = doc.Name
                doc.Close(False)
                closed.append(f"word:{name}")
            except Exception as e:
                errors.append(str(e))
    except Exception as e:
        errors.append(str(e))
    # Close Excel workbooks
    try:
        xl = ExcelApplication()
        count = xl.app.Workbooks.Count
        for i in range(count, 0, -1):
            try:
                wb = xl.app.Workbooks.Item(i)
                name = wb.Name
                wb.Close(False)
                closed.append(f"excel:{name}")
            except Exception as e:
                errors.append(str(e))
    except Exception as e:
        errors.append(str(e))
    # Close PPT presentations
    try:
        ppt = PPTApplication()
        count = ppt.app.Presentations.Count
        for i in range(count, 0, -1):
            try:
                pres = ppt.app.Presentations.Item(i)
                name = pres.Name
                pres.Close()
                closed.append(f"ppt:{name}")
            except Exception as e:
                errors.append(str(e))
    except Exception as e:
        errors.append(str(e))
    return {"closed": closed, "errors": errors}


def save_all() -> Dict:
    saved = []
    errors = []
    # Save Word docs
    try:
        app = get_word_app()
        count = app.Documents.Count
        for i in range(1, count + 1):
            try:
                doc = app.Documents.Item(i)
                name = doc.Name
                doc.Save()
                saved.append(f"word:{name}")
            except Exception as e:
                errors.append(str(e))
    except Exception as e:
        errors.append(str(e))
    # Save Excel workbooks
    try:
        xl = ExcelApplication()
        count = xl.app.Workbooks.Count
        for i in range(1, count + 1):
            try:
                wb = xl.app.Workbooks.Item(i)
                name = wb.Name
                wb.Save()
                saved.append(f"excel:{name}")
            except Exception as e:
                errors.append(str(e))
    except Exception as e:
        errors.append(str(e))
    # Save PPT presentations
    try:
        ppt = PPTApplication()
        count = ppt.app.Presentations.Count
        for i in range(1, count + 1):
            try:
                pres = ppt.app.Presentations.Item(i)
                name = pres.Name
                pres.Save()
                saved.append(f"ppt:{name}")
            except Exception as e:
                errors.append(str(e))
    except Exception as e:
        errors.append(str(e))
    return {"saved": saved, "errors": errors}


def get_word_doc_by_id(doc_id: str) -> Any:
    app_type, index = _resolve(doc_id)
    if app_type != "word":
        raise ValueError(f"Expected word doc_id, got: {doc_id}")
    from .app import get_doc
    return get_doc(index)


def get_excel_wb_by_id(doc_id: str) -> Any:
    app_type, index = _resolve(doc_id)
    if app_type != "excel":
        raise ValueError(f"Expected excel doc_id, got: {doc_id}")
    xl = ExcelApplication()
    return xl.app.Workbooks.Item(index)
