# -*- coding: utf-8 -*-
"""
WPS Bridge Utilities — COM helpers with structured error handling.
Provides both safe (fallback default) and strict (raise exception) modes.
Includes retry with exponential backoff, COM health check, and extended constants.
"""
import functools
import logging
import time
import pythoncom
import win32com.client
from typing import Any, Optional, Callable

_logger = logging.getLogger("wps-agent.com_utils")


class COMError(Exception):
    """Structured COM error with error code."""

    def __init__(self, message: str, code: str = "COM_ERROR", detail: Optional[dict] = None, recoverable: bool = False):
        super().__init__(message)
        self.code = code
        self.detail = detail or {}
        self.recoverable = recoverable

    def to_dict(self) -> dict:
        return {
            "success": False,
            "error_code": self.code,
            "error": str(self),
            "detail": self.detail,
            "recoverable": self.recoverable,
        }


def com_retry(max_retries: int = 3, base_delay: float = 0.5, backoff: float = 2.0):
    """Decorator: retry COM operations with exponential backoff."""
    def decorator(func: Callable):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            last_exc = None
            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except pythoncom.com_error as e:
                    last_exc = e
                    hresult = getattr(e, "hresult", -1)
                    recoverable = hresult in (
                        -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
                        -2147417848,  # RPC_E_DISCONNECTED
                        -2146959355,  # CO_E_OBJNOTCONNECTED
                    )
                    if attempt < max_retries and recoverable:
                        delay = base_delay * (backoff ** attempt)
                        _logger.warning(f"COM retry {attempt + 1}/{max_retries} after {delay:.1f}s (HR={hresult})")
                        time.sleep(delay)
                        pythoncom.CoInitialize()
                        continue
                    raise COMError(
                        f"COM call failed (HR={hresult}): {e}",
                        "COM_RETRY_EXHAUSTED",
                        {"hresult": hresult, "attempts": attempt + 1},
                        recoverable=False,
                    ) from e
                except Exception as e:
                    if attempt < max_retries:
                        delay = base_delay * (backoff ** attempt)
                        _logger.warning(f"General retry {attempt + 1}/{max_retries} after {delay:.1f}s: {e}")
                        time.sleep(delay)
                        continue
                    raise
            raise COMError("Retry exhausted", "COM_RETRY_EXHAUSTED") from last_exc
        return wrapper
    return decorator


def co_init():
    pythoncom.CoInitialize()


def co_uninit():
    pythoncom.CoUninitialize()


def com_health_check() -> dict:
    """Verify COM connection to WPS is alive. Returns health status dict."""
    try:
        app = win32com.client.GetActiveObject("Kwps.Application")
    except Exception:
        try:
            app = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return {"healthy": False, "error": "WPS/Word not running", "error_code": "WPS_NOT_RUNNING"}
    try:
        ver = getattr(app, "Version", "unknown")
        doc_count = getattr(app.Documents, "Count", -1)
        return {
            "healthy": True,
            "version": str(ver),
            "open_documents": doc_count,
            "application": getattr(app, "Name", "unknown"),
        }
    except Exception as e:
        return {"healthy": False, "error": str(e), "error_code": "COM_BROKEN"}


def com_property(obj: Any, name: str, default: Any = None, strict: bool = False) -> Any:
    """
    Get a COM property. If strict=True, raises COMError on failure.
    If strict=False, returns default (legacy safe mode).
    """
    try:
        return getattr(obj, name)
    except Exception as e:
        if strict:
            raise COMError(
                f"Failed to get property '{name}': {e}",
                "COM_PROPERTY_ERROR",
                {"property": name, "exception": str(e)},
            ) from e
        return default


def com_set(obj: Any, name: str, value: Any, strict: bool = False) -> bool:
    """
    Set a COM property. If strict=True, raises COMError on failure.
    """
    try:
        setattr(obj, name, value)
        return True
    except Exception as e:
        if strict:
            raise COMError(
                f"Failed to set property '{name}'={value}: {e}",
                "COM_SET_ERROR",
                {"property": name, "value": value, "exception": str(e)},
            ) from e
        return False


def com_set_batch(obj: Any, properties: dict, strict: bool = False, coerce_types: bool = True) -> list:
    """Set multiple properties. Returns list of failed property names.
    If coerce_types=True, automatically converts float→int for properties that need int.
    """
    INT_PROPERTIES = {"Alignment", "LineSpacingRule", "OutlineLevel", "WidowControl",
                      "KeepWithNext", "KeepTogether", "PageBreakBefore", "ColorIndex",
                      "Underline", "HighlightColorIndex", "Bold", "Italic",
                      "Superscript", "Subscript", "StrikeThrough", "AllCaps",
                      "SmallCaps", "Shadow", "Outline", "Emboss", "Hidden",
                      "Scaling"}
    failed = []
    for name, value in properties.items():
        if value is not None:
            if coerce_types and name in INT_PROPERTIES and isinstance(value, float):
                value = int(value)
            try:
                setattr(obj, name, value)
            except Exception as e:
                if strict:
                    raise COMError(
                        f"Failed to set property '{name}'={value}: {e}",
                        "COM_SET_ERROR",
                        {"property": name, "value": value, "exception": str(e)},
                    ) from e
                failed.append(name)
    return failed


def com_call(obj: Any, name: str, *args, strict: bool = False):
    """Call a COM method."""
    try:
        method = getattr(obj, name)
        return method(*args)
    except Exception as e:
        if strict:
            raise COMError(
                f"Failed to call method '{name}': {e}",
                "COM_CALL_ERROR",
                {"method": name, "args": args, "exception": str(e)},
            ) from e
        return e


def com_release(obj: Any):
    if obj is not None:
        try:
            obj._FlagAsMethod("Release")()
        except Exception:
            pass


def wd_constant(name: str) -> Optional[int]:
    try:
        return getattr(win32com.client.constants, name)
    except Exception:
        return None


# Alignment mappings
WDALIGNMENT = {
    "left": 0, "center": 1, "right": 2, "justify": 3,
    0: "left", 1: "center", 2: "right", 3: "justify",
}

WDLINESPACING = {
    "single": 0, "1.5lines": 1, "double": 2,
    "at_least": 3, "exactly": 4, "multiple": 5,
    0: "single", 1: "1.5lines", 2: "double",
    3: "at_least", 4: "exactly", 5: "multiple",
}

WDSTYLETYPE = {
    "paragraph": 1, "character": 2,
    1: "paragraph", 2: "character",
}

WDBUILTINSTYLE = {
    "Normal": -1,
    "Heading 1": -2, "Heading 2": -3, "Heading 3": -4,
    "Heading 4": -5, "Heading 5": -6, "Heading 6": -7,
    "Heading 7": -8, "Heading 8": -9, "Heading 9": -10,
    "Title": -63, "Subtitle": -64,
    "TOC 1": -73, "TOC 2": -74, "TOC 3": -75,
    "Header": -79, "Footer": -80,
    "Body Text": -91,
    "List Bullet": -111, "List Number": -112,
}

# ── Underline types (WDUnderline) ──
WDUNDERLINE = {
    "none": 0, "single": 1, "words": 2, "double": 3,
    "dotted": 4, "thick": 6, "dash": 7, "dot_dash": 9,
    "dot_dot_dash": 10, "wave": 11, "dash_heavy": 23,
    "dash_long": 39, "dash_long_heavy": 55,
    0: "none", 1: "single", 2: "words", 3: "double",
    4: "dotted", 6: "thick", 7: "dash", 9: "dot_dash",
    10: "dot_dot_dash", 11: "wave",
}

# ── Color index constants for COM (WAColorIndex) ──
WDCOLORINDEX = {
    "auto": 0, "black": 1, "blue": 2, "cyan": 3,
    "green": 4, "magenta": 5, "red": 6, "yellow": 7,
    "white": 8, "dark_blue": 9, "dark_cyan": 10,
    "dark_green": 11, "dark_magenta": 12, "dark_red": 13,
    "dark_yellow": 14, "dark_gray": 15, "light_gray": 16,
    0: "auto", 1: "black", 2: "blue", 3: "cyan",
    4: "green", 5: "magenta", 6: "red", 7: "yellow",
    8: "white", 9: "dark_blue", 10: "dark_cyan",
    11: "dark_green", 12: "dark_magenta", 13: "dark_red",
    14: "dark_yellow", 15: "dark_gray", 16: "light_gray",
}

# ── Text wrapping types (WdWrapType) ──
WDWRAPTYPE = {
    "square": 0, "tight": 1, "through": 2,
    "top_and_bottom": 3, "behind": 4, "in_front_of": 5,
    "inline": 7,
    0: "square", 1: "tight", 2: "through",
    3: "top_and_bottom", 4: "behind", 5: "in_front_of",
    7: "inline",
}

# ── Line numbering restart rule (WdNumberingRule) ──
WDNUMBERINGRULE = {
    "continuous": 0, "restart_page": 1, "restart_section": 2,
    0: "continuous", 1: "restart_page", 2: "restart_section",
}

# ── Page border application (WdBorderType) ──
WDBORDERTYPE = {
    "page_border_top": -1, "page_border_bottom": -2,
    "page_border_left": -3, "page_border_right": -4,
    "page_border_horizontal": -5, "page_border_vertical": -6,
    "page_border_diagonal_down": -7, "page_border_diagonal_up": -8,
}


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_cell(cell_ref: str):
    import re
    m = re.match(r'([A-Za-z]+)(\d+)', cell_ref)
    if not m:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    return m.group(1).upper(), int(m.group(2))


def hex_to_rgb(hex_color: str) -> tuple:
    """Convert hex color like 'FF0000' or '#FF0000' to (R, G, B)."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) == 6:
        return (int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))
    return (0, 0, 0)


def rgb_to_ole_int(hex_color: str) -> int:
    """Convert hex color to OLE COLORREF int (0x00BBGGRR)."""
    r, g, b = hex_to_rgb(hex_color)
    return b << 16 | g << 8 | r


COMPATIBLE_HIGHLIGHT_VALUES = {
    1, 0, 9999999,  # wdNoHighlight, wdAutoColor, wdBrightGreen → common valid highlight options
    2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16,
}