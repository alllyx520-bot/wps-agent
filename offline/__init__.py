# -*- coding: utf-8 -*-
"""Offline document builder — no WPS required."""
from .docx_builder import OfflineDocxBuilder, read_docx_model, write_docx_model
from .pptx_builder import build_pptx, extract_pptx_text
from .xlsx_builder import (
    build_xlsx, analyze_xlsx, convert_csv_to_xlsx,
    validate_formulas_offline, apply_financial_colors_offline,
)

__all__ = [
    "OfflineDocxBuilder", "read_docx_model", "write_docx_model",
    "build_pptx", "extract_pptx_text",
    "build_xlsx", "analyze_xlsx", "convert_csv_to_xlsx",
    "validate_formulas_offline", "apply_financial_colors_offline",
]