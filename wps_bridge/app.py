# -*- coding: utf-8 -*-
"""
WPS Application Bridge — managed COM connection with auto-reconnect.
"""
import win32com.client
import pythoncom
import logging
from typing import Any, Optional, List, Dict
from .utils import com_property, com_set, COMError

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
    """Get WPS Application COM object with auto-reconnect."""
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

    # Try to connect to running instance
    for progid in ("Kwps.Application", "WPS.Application", "ET.Application"):
        try:
            _app = win32com.client.GetObject(None, progid)
            logger.info(f"Connected to WPS via GetObject({progid})")
            break
        except Exception:
            continue

    # Fallback: start a new instance
    if _app is None:
        for progid in ("Kwps.Application", "WPS.Application", "Word.Application"):
            try:
                _app = win32com.client.Dispatch(progid)
                logger.info(f"Started WPS via Dispatch({progid})")
                break
            except Exception:
                continue

    if _app is None:
        logger.error("No running WPS instance found and could not start one.")
        raise COMError("WPS is not running. Please open WPS Office and try again.", "WPS_NOT_RUNNING")

    try:
        com_set(_app, "Visible", visible)
    except Exception:
        pass

    return _app


def get_doc(doc_index: Optional[int] = None) -> Any:
    """Get a document by index or active document."""
    app = get_app()
    if doc_index is not None:
        if doc_index < 1 or doc_index > app.Documents.Count:
            raise COMError(f"Document index {doc_index} out of range (1-{app.Documents.Count})", "DOC_INDEX_OUT_OF_RANGE")
        return app.Documents.Item(doc_index)
    doc = app.ActiveDocument
    if doc is None:
        raise COMError("No document is open. Create or open a document first.", "NO_DOCUMENT")
    return doc


_WPS_CACHED_AVAILABLE: Optional[bool] = None


def check_wps_available() -> bool:
    global _WPS_CACHED_AVAILABLE, _app
    _WPS_CACHED_AVAILABLE = None  # Always check fresh
    try:
        _init_com()
        for progid in ("Kwps.Application", "WPS.Application", "ET.Application"):
            try:
                app = win32com.client.GetObject(None, progid)
                count = app.Documents.Count
                _WPS_CACHED_AVAILABLE = True
                _app = app
                try:
                    com_set(_app, "Visible", True)
                except Exception:
                    pass
                return True
            except Exception:
                continue
        # Fallback: try Dispatch to start a new instance
        for progid in ("Kwps.Application", "WPS.Application", "ET.Application"):
            try:
                app = win32com.client.Dispatch(progid)
                app.Documents.Count
                _WPS_CACHED_AVAILABLE = True
                _app = app
                try:
                    com_set(_app, "Visible", True)
                except Exception:
                    pass
                return True
            except Exception:
                continue
        _WPS_CACHED_AVAILABLE = False
        return False
    except Exception:
        _WPS_CACHED_AVAILABLE = False
        return False


# Tools that require WPS COM to function
WPS_REQUIRED_TOOLS = frozenset({
    "document", "content", "format", "style", "table", "search",
    "layout", "review", "reference", "docspace", "transfer",
    "migrate", "compare", "presentation", "excel",
    "content_control", "field_codes",
})


def list_documents() -> List[Dict]:
    """List all open Word documents."""
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