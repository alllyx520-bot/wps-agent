# -*- coding: utf-8 -*-
"""WPS Bridge — optimized COM interface with batch execution support."""
from .com_client import COMClient, BatchCommandQueue
from .commands import BatchCommandExecutor
from .app import get_app, get_doc, list_documents
from .utils import (
    COMError, com_property, com_set, com_set_batch,
    com_health_check, com_retry,
    WDALIGNMENT, WDLINESPACING, WDSTYLETYPE, WDBUILTINSTYLE,
    WDCOLORINDEX, WDUNDERLINE, WDWRAPTYPE,
    WDNUMBERINGRULE, WDBORDERTYPE,
    hex_to_rgb, rgb_to_ole_int,
)
from .transaction import EditTransaction
from .surgical_context import SurgicalContext
from .format_resolver import resolve_paragraph_format, resolve_run_format

__all__ = [
    "COMClient", "BatchCommandQueue", "BatchCommandExecutor",
    "get_app", "get_doc", "list_documents",
    "COMError", "com_property", "com_set", "com_set_batch",
    "com_health_check", "com_retry",
    "WDALIGNMENT", "WDLINESPACING", "WDSTYLETYPE", "WDBUILTINSTYLE",
    "WDCOLORINDEX", "WDUNDERLINE", "WDWRAPTYPE",
    "WDNUMBERINGRULE", "WDBORDERTYPE",
    "hex_to_rgb", "rgb_to_ole_int",
    "EditTransaction", "SurgicalContext",
    "resolve_paragraph_format", "resolve_run_format",
]