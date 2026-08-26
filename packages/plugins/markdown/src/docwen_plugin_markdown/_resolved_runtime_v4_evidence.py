"""Small AST evidence helpers for the resolved-runtime v4 marker bridge."""

from __future__ import annotations

from typing import Any

from docwen_core.models.resolved_numbering import (
    ResolvedCitation,
    ResolvedDocumentTarget,
    ResolvedNumberingPort,
    ResolvedReference,
)

_TARGET_KEY = "_docwen_resolved_v4_target"
_CAPTION_CHILDREN_KEY = "_docwen_resolved_v4_caption_children"
_REFERENCE_KEY = "_docwen_resolved_v4_reference"
_CITATION_KEY = "_docwen_resolved_v4_citation"


def authored_inline_text(nodes: list[dict[str, Any]]) -> str:
    """Return authored visible inline text without interpreting a number prefix."""

    output: list[str] = []
    for node in nodes:
        node_type = node.get("type")
        if node_type == "text":
            output.append(str(node.get("raw", node.get("text", ""))))
            continue
        if node_type in {"semantic_cross_reference", "semantic_citation", "codespan", "inline_math"}:
            output.append(str(node.get("raw", node.get("text", ""))))
            continue
        if node_type in {"softbreak", "linebreak"}:
            output.append(" ")
            continue
        children = node.get("children")
        if isinstance(children, list):
            output.append(authored_inline_text(children))
    return "".join(output)


def prove_resolved_occurrence_order(
    nodes: list[dict[str, Any]],
    port: ResolvedNumberingPort,
) -> None:
    """Prove target and inline physical order against authenticated ranges."""

    expected_targets = sorted(item.occurrence_key for item in port.document.targets)
    if list(_iter_target_occurrences(nodes)) != expected_targets:
        raise ValueError("v4 target order differs from authenticated source order")

    expected_inline = sorted(
        [(item.source_start, item.source_end, "reference") for item in port.document.references]
        + [(item.source_start, item.source_end, "citation") for item in port.document.citations]
    )
    if list(_iter_inline_occurrences(nodes)) != expected_inline:
        raise ValueError("v4 inline occurrence order differs from authenticated source order")


def _iter_target_occurrences(nodes: list[dict[str, Any]]):
    for node in nodes:
        target = node.get(_TARGET_KEY)
        if isinstance(target, ResolvedDocumentTarget):
            yield target.occurrence_key
        caption_children = node.get(_CAPTION_CHILDREN_KEY)
        if isinstance(caption_children, list):
            yield from _iter_target_occurrences(caption_children)
        children = node.get("children")
        if isinstance(children, list):
            yield from _iter_target_occurrences(children)


def _iter_inline_occurrences(nodes: list[dict[str, Any]]):
    for node in nodes:
        reference = node.get(_REFERENCE_KEY)
        citation = node.get(_CITATION_KEY)
        if isinstance(reference, ResolvedReference):
            yield (reference.source_start, reference.source_end, "reference")
        elif isinstance(citation, ResolvedCitation):
            yield (citation.source_start, citation.source_end, "citation")
        caption_children = node.get(_CAPTION_CHILDREN_KEY)
        if isinstance(caption_children, list):
            yield from _iter_inline_occurrences(caption_children)
        children = node.get("children")
        if isinstance(children, list):
            yield from _iter_inline_occurrences(children)


__all__ = ["authored_inline_text", "prove_resolved_occurrence_order"]
