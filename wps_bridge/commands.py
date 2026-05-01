# -*- coding: utf-8 -*-
"""
Batch Command Executor — high-level batch operations for WPS COM.
Builds command queues for common bulk operations like formatting paragraphs,
setting styles, etc.
"""
from typing import List, Dict, Any, Optional
import logging

from .com_client import COMClient, BatchCommandQueue, COMCommand
from .utils import WDALIGNMENT, WDLINESPACING

logger = logging.getLogger("wps-agent.commands")


class BatchCommandExecutor:
    """Builds and executes common batch command patterns."""

    def __init__(self, client: Optional[COMClient] = None):
        self.client = client or COMClient()
        self.client.init()

    def batch_set_paragraph_format(
        self,
        para_indices: List[int],
        alignment: Optional[str] = None,
        first_line_indent: Optional[float] = None,
        space_before: Optional[float] = None,
        space_after: Optional[float] = None,
        line_spacing: Optional[float] = None,
        line_rule: Optional[str] = None,
        doc_index: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Set paragraph format for multiple paragraphs in one batch."""
        queue = BatchCommandQueue()
        doc_ref = f"Documents.Item({doc_index})" if doc_index else "ActiveDocument"

        for idx in para_indices:
            para_ref = f"{doc_ref}.Paragraphs.Item({idx})"
            pf_ref = f"{para_ref}.Format"

            if alignment is not None:
                align_val = WDALIGNMENT.get(alignment, alignment)
                queue.add_set(pf_ref, "Alignment", align_val)

            if first_line_indent is not None:
                queue.add_set(pf_ref, "FirstLineIndent", first_line_indent)

            if space_before is not None:
                queue.add_set(pf_ref, "SpaceBefore", space_before)

            if space_after is not None:
                queue.add_set(pf_ref, "SpaceAfter", space_after)

            if line_spacing is not None:
                if line_rule == "auto":
                    queue.add_set(pf_ref, "LineSpacingRule", 5)  # wdLineSpaceMultiple
                    queue.add_set(pf_ref, "LineSpacing", line_spacing * 12)  # points to line spacing
                elif line_rule == "exact":
                    queue.add_set(pf_ref, "LineSpacingRule", 4)  # wdLineSpaceExactly
                    queue.add_set(pf_ref, "LineSpacing", line_spacing)
                elif line_rule == "at_least":
                    queue.add_set(pf_ref, "LineSpacingRule", 3)  # wdLineSpaceAtLeast
                    queue.add_set(pf_ref, "LineSpacing", line_spacing)
                else:
                    queue.add_set(pf_ref, "LineSpacingRule", 5)
                    queue.add_set(pf_ref, "LineSpacing", line_spacing * 12)

        results = self.client.execute_batch(queue)
        return {
            "total": len(para_indices),
            "commands_queued": len(queue),
            "results": results,
        }

    def batch_set_font(
        self,
        para_indices: List[int],
        font_name: Optional[str] = None,
        size: Optional[float] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color_index: Optional[int] = None,
        doc_index: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Set font properties for multiple paragraphs in one batch."""
        queue = BatchCommandQueue()
        doc_ref = f"Documents.Item({doc_index})" if doc_index else "ActiveDocument"

        for idx in para_indices:
            font_ref = f"{doc_ref}.Paragraphs.Item({idx}).Range.Font"

            if font_name is not None:
                queue.add_set(font_ref, "Name", font_name)
                queue.add_set(font_ref, "NameFarEast", font_name)

            if size is not None:
                queue.add_set(font_ref, "Size", size)

            if bold is not None:
                queue.add_set(font_ref, "Bold", bold)

            if italic is not None:
                queue.add_set(font_ref, "Italic", italic)

            if color_index is not None:
                queue.add_set(font_ref, "ColorIndex", color_index)

        results = self.client.execute_batch(queue)
        return {
            "total": len(para_indices),
            "commands_queued": len(queue),
            "results": results,
        }

    def batch_apply_style(
        self,
        para_indices: List[int],
        style_name: str,
        doc_index: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Apply a style to multiple paragraphs."""
        queue = BatchCommandQueue()
        doc_ref = f"Documents.Item({doc_index})" if doc_index else "ActiveDocument"

        for idx in para_indices:
            range_ref = f"{doc_ref}.Paragraphs.Item({idx}).Range"
            queue.add_set(range_ref, "Style", style_name)

        results = self.client.execute_batch(queue)
        return {
            "total": len(para_indices),
            "style": style_name,
            "results": results,
        }

    def batch_read_paragraphs(
        self,
        start: int,
        count: int,
        doc_index: Optional[int] = None,
    ) -> List[Dict[str, Any]]:
        """Read multiple paragraphs' text and basic format in one batch."""
        queue = BatchCommandQueue()
        doc_ref = f"Documents.Item({doc_index})" if doc_index else "ActiveDocument"
        results_data = []

        end = start + count
        for idx in range(start, end):
            para_ref = f"{doc_ref}.Paragraphs.Item({idx})"
            queue.add_get(para_ref, "Range.Text")
            queue.add_get(f"{para_ref}.Range.Font", "Name")
            queue.add_get(f"{para_ref}.Range.Font", "Size")
            queue.add_get(f"{para_ref}.Range.Font", "Bold")
            queue.add_get(f"{para_ref}.Format", "Alignment")

        results = self.client.execute_batch(queue)

        # Parse results (5 properties per paragraph)
        for i in range(count):
            base = i * 5
            if base < len(results):
                results_data.append({
                    "index": start + i,
                    "text": str(results[base])[:200] if base < len(results) else "",
                    "font": str(results[base + 1]) if base + 1 < len(results) else "",
                    "size": results[base + 2] if base + 2 < len(results) else 0,
                    "bold": bool(results[base + 3]) if base + 3 < len(results) else False,
                    "alignment": results[base + 4] if base + 4 < len(results) else 0,
                })

        return results_data

    def sync_document_model(self, doc_model, doc_index: Optional[int] = None) -> Dict[str, Any]:
        """Sync a Document model to WPS via batch commands."""
        queue = BatchCommandQueue()
        doc_ref = f"Documents.Item({doc_index})" if doc_index else "ActiveDocument"
        changes = 0

        # Sync paragraph formats
        for i, para in enumerate(doc_model.paragraphs, 1):
            pf_ref = f"{doc_ref}.Paragraphs.Item({i}).Format"

            if para.alignment is not None:
                align_val = WDALIGNMENT.get(para.alignment, 0)
                queue.add_set(pf_ref, "Alignment", align_val)
                changes += 1

            if para.first_line_indent is not None:
                queue.add_set(pf_ref, "FirstLineIndent", para.first_line_indent)
                changes += 1

            if para.space_before is not None:
                queue.add_set(pf_ref, "SpaceBefore", para.space_before)
                changes += 1

            if para.space_after is not None:
                queue.add_set(pf_ref, "SpaceAfter", para.space_after)
                changes += 1

            # Sync run formats
            for j, run in enumerate(para.runs):
                if j == 0:
                    # Apply to entire paragraph range for simplicity
                    font_ref = f"{doc_ref}.Paragraphs.Item({i}).Range.Font"
                    if run.font:
                        queue.add_set(font_ref, "Name", run.font)
                        queue.add_set(font_ref, "NameFarEast", run.font)
                    if run.size:
                        queue.add_set(font_ref, "Size", run.size)
                    if run.bold:
                        queue.add_set(font_ref, "Bold", True)
                    if run.italic:
                        queue.add_set(font_ref, "Italic", True)
                    if run.color and run.color != "auto":
                        # Color setting is complex in COM, skip for now
                        pass
                    changes += 1

        results = self.client.execute_batch(queue)
        return {
            "synced": True,
            "changes_applied": changes,
            "commands_queued": len(queue),
            "results": results,
        }
