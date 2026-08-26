"""AST transformations applied between parsing and rendering.

The helpers here attach current renderer metadata that Mistune's normalized
AST cannot preserve on its own, such as heading/body merge markers and the
source marker used for thematic breaks.
"""

from __future__ import annotations

from typing import Any

_HR_LINE_STRIPPED_VALUES = frozenset({"---", "***", "___"})
_HR_MARKER_KEYS = {
    "---": "dash",
    "***": "asterisk",
    "___": "underscore",
}


def annotate_ast_with_merges(
    ast: list[dict[str, Any]],
    merge_indices: set[int],
) -> None:
    """Annotate heading AST nodes with merge information.

    For each heading whose 0‑based index is in *merge_indices*:
    - Sets ``node["_merge"] = True`` on the heading node
    - Sets ``next_node["_merged_into_heading"] = True`` on the following
      paragraph node (so the renderer can skip it)

    Args:
        ast: List of mistune AST token dicts.
        merge_indices: Set of 0‑based heading indexes to merge from
            :func:`~preprocessor.detect_heading_merges`.
    """
    heading_idx = 0
    for i, node in enumerate(ast):
        if node.get("type") == "heading":
            if heading_idx in merge_indices:
                node["_merge"] = True
                # Annotate the next sibling if it's a paragraph
                if i + 1 < len(ast) and ast[i + 1].get("type") == "paragraph":
                    ast[i + 1]["_merged_into_heading"] = True
            heading_idx += 1


def annotate_ast_with_hr_attachments(
    ast: list[dict[str, Any]],
    hr_attachments_line_set: set[int],
    md_body: str,
) -> None:
    """Annotate thematic_break AST nodes with attachment info.

    Converts line-based HR attachment detection (from
    :func:`~preprocessor.detect_hr_attachments`) to AST node annotation
    by aligning HR positions in the source text with the ordered
    ``thematic_break`` tokens in the AST.

    Args:
        ast: List of mistune AST token dicts.
        hr_attachments_line_set: Set of 0‑based source line indexes where
            an attached HR occurs.
        md_body: Raw markdown source text used to enumerate HR positions.
    """
    # Build list of source line indices and marker kinds that contain HR markers.
    # Mistune normalizes all of these to the same thematic_break token, so the
    # original marker has to be restored from source order before rendering.
    lines = md_body.split("\n")
    hr_source_markers: list[tuple[int, str]] = []
    for line_idx, line in enumerate(lines):
        marker = line.strip()
        if marker in _HR_LINE_STRIPPED_VALUES:
            hr_source_markers.append((line_idx, _HR_MARKER_KEYS[marker]))

    # Determine which HR positions (by order) should attach
    hr_attach_indices: set[int] = set()
    for hr_idx, (line_idx, _marker_key) in enumerate(hr_source_markers):
        if line_idx in hr_attachments_line_set:
            hr_attach_indices.add(hr_idx)

    # Annotate matching thematic_break nodes in the AST.
    hr_ast_idx = 0
    for node in ast:
        if node.get("type") == "thematic_break":
            if hr_ast_idx < len(hr_source_markers):
                node["_hr_marker"] = hr_source_markers[hr_ast_idx][1]
            if hr_ast_idx in hr_attach_indices:
                node["_attach_to_prev"] = True
            hr_ast_idx += 1
