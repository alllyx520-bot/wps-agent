# -*- coding: utf-8 -*-
"""
Offline Docx Builder — full read/write capabilities without WPS.
Uses the docx_engine for XML-native document processing.
"""
from pathlib import Path
from typing import Dict, Any, Optional

from docx_engine import (
    parse_docx, unpack_docx, pack_docx,
    build_document_model, serialize_document_model,
    Document, DocumentAnalyzer, Formatter, StyleResolver,
    DocxEngineError, ParseError, SerializeError,
)


class OfflineDocxBuilder:
    """High-level offline builder for .docx files."""

    def __init__(self):
        self.document: Optional[Document] = None
        self._source_path: Optional[str] = None
        self._blank_template = str(Path(__file__).parents[1] / "blank_template.docx")

    def load(self, docx_path: str) -> Document:
        """Load a .docx file into the document model."""
        self._source_path = docx_path
        self.document = build_document_model(parse_docx(docx_path))
        return self.document

    def create(self) -> Document:
        """Create a new empty document, using blank template as source."""
        self.document = Document()
        self._source_path = self._blank_template
        return self.document

    def save(self, output_path: str) -> str:
        """Save the document model to a .docx file."""
        if self.document is None:
            raise SerializeError("No document loaded or created")
        return serialize_document_model(
            self.document,
            output_path,
            original_docx=self._source_path,
        )

    def analyze(self) -> Dict[str, Any]:
        """Analyze the current document."""
        if self.document is None:
            return {"error": "No document loaded"}
        analyzer = DocumentAnalyzer(self.document)
        return {
            "document_type": analyzer.detect_document_type(),
            "statistics": self.document.get_statistics(),
            "outline": analyzer.get_document_outline(),
            "headings": self.document.get_heading_structure(),
            "quality": analyzer.analyze_formatting_quality(),
        }

    def auto_format(self, document_type: Optional[str] = None) -> Dict[str, Any]:
        """Apply automatic formatting."""
        if self.document is None:
            return {"error": "No document loaded"}

        if document_type is None:
            analyzer = DocumentAnalyzer(self.document)
            document_type = analyzer.detect_document_type()

        resolver = None
        if self.document.styles:
            resolver = StyleResolver()
            resolver.styles = self.document.styles

        formatter = Formatter(self.document, resolver)
        return formatter.auto_format(document_type)

    def apply_template(self, template_name: str) -> Dict[str, Any]:
        """Apply a named template."""
        if self.document is None:
            return {"error": "No document loaded"}
        resolver = None
        if self.document.styles:
            resolver = StyleResolver()
            resolver.styles = self.document.styles
        formatter = Formatter(self.document, resolver)
        return formatter.apply_template(template_name)

    def add_numbering(self) -> Dict[str, Any]:
        """Add multi-level numbering to headings."""
        if self.document is None:
            return {"error": "No document loaded"}
        resolver = None
        if self.document.styles:
            resolver = StyleResolver()
            resolver.styles = self.document.styles
        formatter = Formatter(self.document, resolver)
        return formatter.add_multi_level_numbering()

    def replace_text(self, old: str, new: str, case_sensitive: bool = True) -> int:
        """Replace text across all runs."""
        if self.document is None:
            return 0
        return self.document.replace_text(old, new, case_sensitive)

    def get_text(self) -> str:
        """Get full document text."""
        if self.document is None:
            return ""
        return self.document.text

    def get_statistics(self) -> Dict[str, Any]:
        """Get document statistics."""
        if self.document is None:
            return {}
        return self.document.get_statistics()


# Convenience functions
def read_docx_model(docx_path: str) -> Document:
    """Read a .docx file and return the Document model."""
    return build_document_model(parse_docx(docx_path))


def write_docx_model(document: Document, output_path: str, original_docx: Optional[str] = None) -> str:
    """Write a Document model to a .docx file."""
    return serialize_document_model(document, output_path, original_docx)