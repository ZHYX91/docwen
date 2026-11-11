"""Distinguish original template paragraphs from explicitly bound generation."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

from docx.oxml.ns import qn

from docwen_core._docx_numbering_package import _elements_equal


class TemplateParagraphInventory:
    """Keep template allowances attached to original nodes, then prove saved order."""

    def __init__(self, document: Any, style_ids: set[str], purpose: str, *, error_type: type[ValueError]) -> None:
        self._style_ids = frozenset(style_ids)
        self._purpose = purpose
        self._error_type = error_type
        self._original = tuple(self._paragraphs(document))
        self._bound: dict[tuple[int, int, str], Any] = {}
        self._indices: dict[tuple[int, int, str], int] | None = None
        self._original_snapshots: dict[int, Any] = {}

    def _paragraphs(self, document: Any) -> list[Any]:
        return [
            paragraph
            for paragraph in document.element.body.iter(qn("w:p"))
            if any(
                reference.get(qn("w:val")) in self._style_ids
                for reference in paragraph.findall(f"{qn('w:pPr')}/{qn('w:pStyle')}")
            )
        ]

    def bind(self, key: tuple[int, int, str], paragraph: Any) -> None:
        if self._indices is not None or key in self._bound:
            raise self._error_type(f"{self._purpose} inventory cannot bind this target")
        if paragraph in self._original or paragraph in self._bound.values():
            raise self._error_type(f"{self._purpose} target reuses an existing paragraph")
        self._bound[key] = paragraph

    def capture(self, document: Any) -> None:
        paragraphs = self._paragraphs(document)
        present = set(paragraphs)
        surviving_originals = present.intersection(self._original)
        expected = surviving_originals.union(self._bound.values())
        if len(paragraphs) != len(expected) or present != expected:
            raise self._error_type(f"{self._purpose} paragraph inventory differs from bound identities")
        indices = {paragraph: index for index, paragraph in enumerate(paragraphs)}
        bound_order = sorted(self._bound, key=lambda key: indices[self._bound[key]])
        if bound_order != sorted(self._bound):
            raise self._error_type(f"{self._purpose} paragraph order differs from source authority")
        self._indices = {key: indices[paragraph] for key, paragraph in self._bound.items()}
        # Template filling can consume anchors and edit metadata before capture.
        # Copies never inherit a removed original node's template allowance.
        self._original_snapshots = {indices[paragraph]: deepcopy(paragraph) for paragraph in surviving_originals}

    def prove(self, document: Any) -> dict[tuple[int, int, str], Any]:
        if self._indices is None:
            raise self._error_type(f"{self._purpose} paragraph inventory was not finalized")
        paragraphs = self._paragraphs(document)
        if len(paragraphs) != len(self._indices) + len(self._original_snapshots):
            raise self._error_type(f"{self._purpose} paragraph inventory differs after reopen")
        if any(
            not _elements_equal(paragraphs[index], original) for index, original in self._original_snapshots.items()
        ):
            raise self._error_type(f"template {self._purpose} paragraph changed after finalization")
        return {key: paragraphs[index] for key, index in self._indices.items()}
