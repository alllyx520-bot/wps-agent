# -*- coding: utf-8 -*-
import pythoncom
import win32com.client
from typing import Any, Optional


def co_init():
    pythoncom.CoInitialize()


def co_uninit():
    pythoncom.CoUninitialize()


def com_property(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def com_set(obj: Any, name: str, value: Any) -> bool:
    try:
        setattr(obj, name, value)
        return True
    except Exception:
        return False


def com_set_batch(obj: Any, properties: dict) -> list:
    failed = []
    for name, value in properties.items():
        if value is not None:
            try:
                setattr(obj, name, value)
            except Exception:
                failed.append(name)
    return failed


def com_call(obj: Any, name: str, *args):
    try:
        method = getattr(obj, name)
        return method(*args)
    except Exception as e:
        return e


def com_release(obj: Any):
    if obj is not None:
        try:
            obj._FlagAsMethod("Release")()
        except Exception:
            pass


def wd_constant(name: str) -> int:
    try:
        return getattr(win32com.client.constants, name)
    except Exception:
        return -1


WDBUILTINSTYLE = {
    "Normal": -1,
    "Heading 1": -2,
    "Heading 2": -3,
    "Heading 3": -4,
    "Heading 4": -5,
    "Heading 5": -6,
    "Heading 6": -7,
    "Heading 7": -8,
    "Heading 8": -9,
    "Heading 9": -10,
    "Title": -63,
    "Subtitle": -64,
    "TOC 1": -73,
    "TOC 2": -74,
    "TOC 3": -75,
    "Header": -79,
    "Footer": -80,
    "Body Text": -91,
    "List Bullet": -111,
    "List Number": -112,
}


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


WDGOTO = {
    "line": 1, "page": 0, "section": 2, "bookmark": -1,
    "heading": 11,
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
