# -*- coding: utf-8 -*-
"""
Docx Engine — Native XML document processing engine.
Provides precise read/write capabilities for .docx files without requiring WPS.
"""
from .document_model import Document, Paragraph, Run, Table, Cell
from .xml_parser import parse_docx, unpack_docx, pack_docx
from .serializer import build_document_model, serialize_document_model, load_and_build
from .style_resolver import StyleResolver
from .intelligence import DocumentAnalyzer
from .formatter import Formatter
from .errors import DocxEngineError, ParseError, SerializeError, ValidationError, ErrorCode
from .semantic_model import SemanticParser, DocumentGraph, SemanticElement, SemanticRole
from .layout_model import LayoutAnalyzer, LayoutReport, LayoutIssue

__all__ = [
    "Document", "Paragraph", "Run", "Table", "Cell",
    "parse_docx", "unpack_docx", "pack_docx",
    "build_document_model", "serialize_document_model", "load_and_build",
    "StyleResolver",
    "DocumentAnalyzer",
    "Formatter",
    "SemanticParser", "DocumentGraph", "SemanticElement", "SemanticRole",
    "LayoutAnalyzer", "LayoutReport", "LayoutIssue",
    "DocxEngineError", "ParseError", "SerializeError", "ValidationError",
    "ErrorCode",
]