"""Render annotations derived from actual parsed Markdown siblings."""

from __future__ import annotations

from itertools import pairwise
from typing import Any

from docwen_core.text.heading_merge import HEADING_MERGE_PUNCTUATION_SET


def _heading_text(node: dict[str, Any]) -> str:
    """Read visible inline text without dropping an opaque trailing atom."""
    kind = node.get("type")
    if kind in {"image", "inline_math", "inline_latex"}:
        return "\ufffc"
    if "children" in node:
        return "".join(_heading_text(child) for child in node["children"])
    if kind in {"text", "codespan", "inline_html"}:
        return str(node.get("raw", ""))
    if kind in {"softbreak", "linebreak"}:
        return "\n"
    # A formula, image or semantic atom after punctuation is not punctuation.
    return "\ufffc"


def annotate_ast_with_merges(
    ast: list[dict[str, Any]],
    *,
    mode: str = "punct_required",
    punctuation: frozenset[str] | None = None,
) -> None:
    """Join only an actual heading and its immediately adjacent paragraph.

    Blank lines and other block constructs remain barriers. Setext headings
    and dialect-specific heading levels share this same parsed contract; code
    examples and nested headings cannot shift another heading's identity.
    """
    if mode not in {"punct_required", "always", "never"}:
        mode = "punct_required"
    if mode == "never":
        return
    punct = punctuation if punctuation is not None else HEADING_MERGE_PUNCTUATION_SET
    for node, following in pairwise(ast):
        if node.get("type") != "heading" or following.get("type") != "paragraph":
            continue
        children = following.get("children", [])
        # Keep the established exclusion for table-like and display-math text
        # even when incomplete syntax falls back to an ordinary paragraph.
        if (
            children
            and children[0].get("type") == "text"
            and str(children[0].get("raw", "")).lstrip().startswith(("|", "$$"))
        ):
            continue
        text = _heading_text(node).rstrip()
        if mode == "always" or (text and text[-1] in punct):
            node["_merge"] = True
            following["_merged_into_heading"] = True


def _protected_inline(node: dict[str, Any]) -> bool:
    return (
        any(key.startswith("_docwen_") for key in node)
        or str(node.get("type", "")).startswith(("_docwen_", "semantic_"))
        or node.get("type") in {"inline_html", "footnote_ref"}
        or any(_protected_inline(child) for child in node.get("children", []))
    )


def _split_inline_lines(children: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    parts: list[list[dict[str, Any]]] = [[]]
    for child in children:
        if child.get("type") == "softbreak":
            parts.append([])
            continue
        nested = child.get("children")
        fragments = (
            [[{**child, "children": part}] if part else [] for part in _split_inline_lines(nested)]
            if isinstance(nested, list)
            else [[child]]
        )
        parts[-1].extend(fragments[0])
        parts.extend(fragments[1:])
    return parts


def split_plain_paragraph_lines(ast: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Make ordinary Markdown soft-separated lines independent Word paragraphs.

    This is a request-owned document export policy, not a change to Markdown
    syntax or to explicit hard breaks. Only unowned top-level prose is split:
    semantic anchors, captions and all nested block structures retain their
    original ownership and paragraph boundaries.
    """
    output: list[dict[str, Any]] = []
    for node in ast:
        children = node.get("children")
        if node.get("type") != "paragraph" or not isinstance(children, list) or _protected_inline(node):
            output.append(node)
            continue
        parts = _split_inline_lines(children)
        if len(parts) == 1:
            output.append(node)
            continue
        # Do not invent empty Word paragraphs from an invalid/empty segment.
        if not all(
            part and any(item.get("type") != "text" or str(item.get("raw", "")).strip() for item in part)
            for part in parts
        ):
            output.append(node)
            continue
        output.extend({**node, "children": part} for part in parts)
    return output
