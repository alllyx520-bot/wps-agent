# -*- coding: utf-8 -*-
"""Structured error codes for the docx engine."""
from enum import Enum
from typing import Optional, Dict, Any


class ErrorCode(str, Enum):
    # Parse errors
    ZIP_INVALID = "ZIP_INVALID"
    XML_MALFORMED = "XML_MALFORMED"
    MISSING_PART = "MISSING_PART"
    RELATIONSHIP_BROKEN = "RELATIONSHIP_BROKEN"
    NAMESPACE_UNKNOWN = "NAMESPACE_UNKNOWN"

    # Serialize errors
    SERIALIZE_FAILED = "SERIALIZE_FAILED"
    ZIP_PACK_FAILED = "ZIP_PACK_FAILED"
    VALIDATION_FAILED = "VALIDATION_FAILED"

    # COM bridge errors
    COM_NOT_INITIALIZED = "COM_NOT_INITIALIZED"
    COM_DISCONNECTED = "COM_DISCONNECTED"
    WPS_NOT_RUNNING = "WPS_NOT_RUNNING"
    DOCUMENT_NOT_FOUND = "DOCUMENT_NOT_FOUND"
    OPERATION_TIMEOUT = "OPERATION_TIMEOUT"

    # Logic errors
    PARAGRAPH_NOT_FOUND = "PARAGRAPH_NOT_FOUND"
    RUN_NOT_FOUND = "RUN_NOT_FOUND"
    TABLE_NOT_FOUND = "TABLE_NOT_FOUND"
    STYLE_NOT_FOUND = "STYLE_NOT_FOUND"
    INVALID_RANGE = "INVALID_RANGE"
    INVALID_FORMAT = "INVALID_FORMAT"

    # General
    UNKNOWN = "UNKNOWN"


class DocxEngineError(Exception):
    """Base exception for docx engine."""

    def __init__(
        self,
        message: str,
        code: ErrorCode = ErrorCode.UNKNOWN,
        detail: Optional[Dict[str, Any]] = None,
    ):
        super().__init__(message)
        self.code = code
        self.detail = detail or {}

    def to_dict(self) -> Dict[str, Any]:
        return {
            "success": False,
            "error_code": self.code.value,
            "error": str(self),
            "detail": self.detail,
        }


class ParseError(DocxEngineError):
    """Error during ZIP unpack or XML parse."""
    pass


class SerializeError(DocxEngineError):
    """Error during XML serialize or ZIP pack."""
    pass


class ValidationError(DocxEngineError):
    """Error during document validation."""
    pass


class COMError(DocxEngineError):
    """Error during COM/WPS operations."""
    pass
