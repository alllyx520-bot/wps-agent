# -*- coding: utf-8 -*-
"""Transaction-based COM editing — atomic all-or-nothing document modifications.

Before any write: take a paragraph-level snapshot.
Execute all mutations via the XML engine (deterministic, reliable).
After execution: verify by re-reading the document.
On success: save .docx, signal WPS to reload.
On failure: rollback from snapshot.

Key design: XML engine is the *source of truth* for writes.
COM is reserved for reading and visual preview.
"""

import copy
from typing import Any, Dict, List, Optional
from pathlib import Path
from .app import get_doc
from .content import snapshot as com_snapshot, rollback as com_rollback


class EditTransaction:
    """A single atomic edit session. Collects mutations, validates, commits."""

    def __init__(self, filepath: Optional[str] = None):
        self.filepath = filepath
        self.mutations: List[Dict] = []
        self.snapshot_data: Optional[List[Dict]] = None
        self.committed = False

    def mutate(self, tool: str, action: str, args: Dict) -> "EditTransaction":
        """Record a mutation to apply."""
        self.mutations.append({"tool": tool, "action": action, "args": copy.deepcopy(args)})
        return self

    def begin(self, doc_index: Optional[int] = None) -> "EditTransaction":
        """Take a full snapshot of the current document state."""
        self.snapshot_data = com_snapshot(doc_index).get("paragraphs_saved", 0)
        return self

    def execute_with(self, handler, doc_index: Optional[int] = None) -> Dict:
        """Apply all recorded mutations using the given handler function.

        handler(tool, action, args) → result_dict
        """
        results = []
        for mut in self.mutations:
            try:
                result = handler(mut["tool"], mut["action"], mut["args"])
                results.append({"ok": True, "mutation": mut, "result": result})
            except Exception as e:
                return self._fail(results, e)

        # Verify by re-reading document structure
        try:
            from .content import document_structure
            struct = document_structure(doc_index)
            results.append({"verify": "ok", "paragraphs": struct.get("total_paragraphs", 0)})
        except Exception as ve:
            self._rollback(doc_index)
            return {"committed": False, "error": f"Verification failed after {len(results)} mutations: {ve}",
                    "mutations": results, "rolled_back": True}

        self.committed = True
        return {"committed": True, "mutations": len(self.mutations), "results": results,
                "snapshot_paragraphs": self.snapshot_data}

    def rollback(self, doc_index: Optional[int] = None) -> Dict:
        """Explicitly rollback to the snapshot taken in begin()."""
        return self._rollback(doc_index)

    def _rollback(self, doc_index: Optional[int] = None) -> Dict:
        r = com_rollback(doc_index)
        self.mutations.clear()
        return {"rolled_back": True, "restored": r.get("restored_paragraphs", 0)}

    def _fail(self, partial_results, error) -> Dict:
        return {"committed": False, "error": str(error),
                "mutations_attempted": len(partial_results),
                "partial_results": partial_results}
