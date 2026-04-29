# -*- coding: utf-8 -*-
import win32com.client
import pythoncom
import logging
from typing import Any, Optional, List, Dict
from .utils import com_property, com_set

logger = logging.getLogger("wps-agent.bridge")

_app: Any = None
_initialized: bool = False
_visible: bool = True


def _init_com():
    global _initialized
    if not _initialized:
        pythoncom.CoInitialize()
        _initialized = True


def get_app(visible: bool = True, force_new: bool = False) -> Any:
    global _app, _visible
    _visible = visible
    _init_com()
    if _app is not None and not force_new:
        try:
            _app.Documents.Count
            return _app
        except Exception:
            logger.warning("COM connection lost, reconnecting...")
            _app = None
    if not force_new:
        try:
            _app = win32com.client.GetObject(None, "Kwps.Application")
        except Exception:
            _app = win32com.client.Dispatch("Kwps.Application")
    else:
        _app = win32com.client.Dispatch("Kwps.Application")
    com_set(_app, "Visible", visible)
    return _app


def get_doc(doc_index: Optional[int] = None) -> Any:
    app = get_app()
    if doc_index is not None:
        return app.Documents.Item(doc_index)
    return app.ActiveDocument


def list_documents() -> List[Dict]:
    app = get_app()
    docs = []
    try:
        count = app.Documents.Count
    except Exception:
        return docs
    for i in range(1, count + 1):
        try:
            doc = app.Documents.Item(i)
            docs.append({
                "index": i,
                "name": com_property(doc, "Name", ""),
                "full_name": com_property(doc, "FullName", ""),
                "paragraph_count": com_property(doc.Paragraphs, "Count", 0),
                "table_count": com_property(doc.Tables, "Count", 0),
                "saved": com_property(doc, "Saved", False),
            })
        except Exception:
            continue
    return docs
