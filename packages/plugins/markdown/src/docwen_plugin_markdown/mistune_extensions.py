"""Extended mistune Markdown parser with custom inline plugins.

Uses mistune v3's plugin API (``md.inline.register``). The built-in
formatting plugins from ``mistune.plugins.formatting`` provide:

- ``==highlight==``  — via plugin ``mark`` (``<mark>`` element)
- ``^^insert^^``    — via plugin ``insert`` (``<ins>`` element)
- ``^superscript^`` — via plugin ``superscript``
- ``~subscript~``   — via plugin ``subscript``

We also provide a custom ``plugin_underline`` for ``<u>text</u>`` inline
HTML support (rendered as underline in AST mode).

The module exposes ``parse_markdown_text()`` as the Markdown plugin's parser
entry point.
"""

from __future__ import annotations

import re
from typing import Any

import mistune
from mistune.plugins.formatting import (
    insert as _insert,
)
from mistune.plugins.formatting import (
    mark as _mark,
)
from mistune.plugins.formatting import (
    subscript as _subscript,
)
from mistune.plugins.formatting import (
    superscript as _superscript,
)

_EXTENDED_ATX_HEADING_TRIM = re.compile(r"(\s+|^)#+\s*$")
_STRUCTURAL_TABLE_DELIMITER_CELL = re.compile(r"^:?-{3,}:?$")
_STRUCTURAL_TABLE_BLOCK = (
    r"^ {0,3}(?P<structural_table_rows>"
    r"(?:\|[^\n]*\|[ \t]*(?:\n|$)){2,})"
)

# ═══════════════════════════════════════════════════════════════════════════
# Custom plugin: <u>underline</u>
# ═══════════════════════════════════════════════════════════════════════════


def plugin_underline(md: mistune.Markdown) -> None:
    """Mistune v3 plugin: inline ``<u>text</u>`` → underline token.

    Registers a parse rule that captures ``<u>...</u>`` and emits a
    ``underline`` token in AST mode.
    """

    PATTERN = r"<u>(.*?)</u>"

    def parse_underline(inline, m, state):
        text = m.group(1)
        new_state = state.copy()
        new_state.src = text
        children = inline.render(new_state)
        state.append_token({"type": "underline", "children": children})
        return m.end()

    md.inline.register(
        "underline",
        PATTERN,
        parse_underline,
        before="linebreak",
    )


# ═══════════════════════════════════════════════════════════════════════════
# Custom plugin: standalone single-line display math
# ═══════════════════════════════════════════════════════════════════════════


def plugin_single_line_block_math(md: mistune.Markdown) -> None:
    """Parse a standalone ``$$formula$$`` line as ``block_math``.

    Mistune's built-in math plugin only treats the three-line form as display
    math.  The maintained parser contract also accepts the common single-line form; if
    it falls through to the inline parser, the two outer dollar signs become
    literal text in Word.  Registering this as a block rule keeps fenced code
    untouched and gives both spellings the same AST contract.
    """

    pattern = r"^ {0,3}\$\$(?P<single_line_math_text>[^\n]+?)\$\$[ \t]*$"

    def parse_single_line_block_math(block, match, state):
        state.append_token(
            {
                "type": "block_math",
                "raw": match.group("single_line_math_text").strip(),
            }
        )
        return match.end() + 1

    md.block.register(
        "single_line_block_math",
        pattern,
        parse_single_line_block_math,
        before="block_math",
    )


def plugin_extended_atx_headings(md: mistune.Markdown) -> None:
    """Parse DocWen's legacy-compatible ATX heading levels 7 through 9.

    CommonMark and Mistune stop at level 6.  Word exposes nine outline
    levels, and classic DocWen represented the last three with seven to nine
    leading ``#`` markers.  Keep that extension narrow: ten or more markers
    remain ordinary paragraph text.
    """

    pattern = (
        r"^ {0,3}(?P<extended_atx_marks>#{7,9})(?!#)"
        r"(?P<extended_atx_text>[ \t]*|[ \t]+.*?)$"
    )

    def parse_extended_atx_heading(block, match, state):
        text = match.group("extended_atx_text").strip()
        if text:
            text = _EXTENDED_ATX_HEADING_TRIM.sub("", text)
        state.append_token(
            {
                "type": "heading",
                "text": text,
                "attrs": {"level": len(match.group("extended_atx_marks"))},
                "style": "atx",
            }
        )
        return match.end() + 1

    md.block.register(
        "extended_atx_heading",
        pattern,
        parse_extended_atx_heading,
        before="atx_heading",
    )


def _structural_pipe_row(line: str) -> list[str] | None:
    """Split one Structural Tables row without consuming escaped/code pipes."""

    if "|" not in line:
        return None
    segments: list[str] = []
    current: list[str] = []
    escaped = False
    code_ticks = 0
    index = 0
    while index < len(line):
        character = line[index]
        if escaped:
            current.append(character)
            escaped = False
            index += 1
            continue
        if character == "\\":
            current.append(character)
            escaped = True
            index += 1
            continue
        if character == "`":
            run = 1
            while index + run < len(line) and line[index + run] == "`":
                run += 1
            current.extend("`" * run)
            if code_ticks == 0:
                code_ticks = run
            elif code_ticks == run:
                code_ticks = 0
            index += run
            continue
        if character == "|" and code_ticks == 0:
            segments.append("".join(current))
            current = []
        else:
            current.append(character)
        index += 1
    segments.append("".join(current))
    if segments and not segments[0].strip():
        segments.pop(0)
    if segments and not segments[-1].strip():
        segments.pop()
    return segments or None


def _structural_delimiter(line: str) -> tuple[list[str | None], int] | None:
    cells = _structural_pipe_row(line)
    if cells is None:
        return None
    boundaries = [index for index, cell in enumerate(cells) if cell == ""]
    if any(not cell and cell != "" for cell in cells):  # pragma: no cover - defensive type boundary
        return None
    if any(cell != "" and not cell.strip() for cell in cells):
        return None
    if len(boundaries) > 1:
        return None
    boundary = boundaries[0] if boundaries else None
    if boundary is not None and (boundary == 0 or boundary == len(cells) - 1):
        return None
    tokens = [cell for index, cell in enumerate(cells) if index != boundary]
    if not tokens or any(_STRUCTURAL_TABLE_DELIMITER_CELL.fullmatch(token.strip()) is None for token in tokens):
        return None
    alignments: list[str | None] = []
    for token in tokens:
        stripped = token.strip()
        if stripped.startswith(":") and stripped.endswith(":"):
            alignments.append("center")
        elif stripped.startswith(":"):
            alignments.append("left")
        elif stripped.endswith(":"):
            alignments.append("right")
        else:
            alignments.append(None)
    return alignments, 0 if boundary is None else boundary


def _structural_table_cell(text: str, alignment: str | None, *, head: bool) -> dict[str, Any]:
    stripped = text.strip()
    return {
        "type": "table_cell",
        "text": stripped,
        "attrs": {
            "align": alignment,
            "head": head,
            "docwen_literal_merge_marker": stripped in {r"\<", r"\^"},
        },
    }


def plugin_structural_tables(md: mistune.Markdown) -> None:
    """Parse the Obsidian Structural Tables source dialect.

    The extension owns only structurally distinctive tables: multiple header
    rows, a delimiter ``||`` row-header boundary, or merge markers. Ordinary
    GFM tables continue through Mistune's built-in table plugin.
    """

    def parse_structural_table(block, match, state):
        raw = match.group("structural_table_rows")
        physical_lines = raw.rstrip("\n").splitlines()
        rows = [_structural_pipe_row(line) for line in physical_lines]
        delimiters = [
            (index, parsed)
            for index, line in enumerate(physical_lines)
            if (parsed := _structural_delimiter(line)) is not None
        ]
        if len(delimiters) != 1:
            return None
        delimiter_index, (alignments, header_columns) = delimiters[0]
        column_count = len(alignments)
        if delimiter_index < 1:
            return None
        content_rows = [*rows[:delimiter_index], *rows[delimiter_index + 1 :]]
        if any(row is None or len(row) != column_count for row in content_rows):
            return None
        concrete_rows = [row for row in content_rows if row is not None]
        marker_found = any(cell.strip() in {"<", "^"} for row in concrete_rows for cell in row)
        structural = marker_found or delimiter_index > 1 or header_columns > 0

        head_rows = [
            {
                "type": "table_row",
                "children": [
                    _structural_table_cell(cell, alignments[column], head=True) for column, cell in enumerate(row or [])
                ],
            }
            for row in rows[:delimiter_index]
        ]
        body_rows = [
            {
                "type": "table_row",
                "children": [
                    _structural_table_cell(cell, alignments[column], head=False)
                    for column, cell in enumerate(row or [])
                ],
            }
            for row in rows[delimiter_index + 1 :]
        ]
        token: dict[str, Any] = {
            "type": "table",
            "children": [
                {"type": "table_head", "children": head_rows},
                {"type": "table_body", "children": body_rows},
            ],
        }
        if structural:
            token["_structural_table"] = {
                "header_rows": delimiter_index,
                "header_columns": header_columns,
            }
        state.append_token(token)
        return match.end()

    md.block.register(
        "structural_table",
        _STRUCTURAL_TABLE_BLOCK,
        parse_structural_table,
        before="table",
    )


# ═══════════════════════════════════════════════════════════════════════════
# Factory
# ═══════════════════════════════════════════════════════════════════════════


def create_extended_markdown(*, auto_link_bare_url: bool = False) -> mistune.Markdown:
    """Create a mistune Markdown instance with the requested plugins.

    Includes:
    - Built-in: table, footnotes, strikethrough, task_lists, math
    - Formatting: mark (``==text==``), insert (``^^text^^``),
      superscript (``^text^``), subscript (``~text~``)
    - Custom: underline (``<u>text</u>``)

    Returns:
        A mistune ``Markdown`` instance configured with ``renderer="ast"``
        for structured output.
    """
    plugins: list[Any] = [
        "table",
        "footnotes",
        "strikethrough",
        "task_lists",
        "math",
        plugin_structural_tables,
        _mark,
        _insert,
        _superscript,
        _subscript,
        plugin_single_line_block_math,
        plugin_extended_atx_headings,
        plugin_underline,
    ]
    if auto_link_bare_url:
        plugins.append("url")

    return mistune.create_markdown(
        renderer="ast",
        plugins=plugins,
    )


# ═══════════════════════════════════════════════════════════════════════════
# Parsing entry point
# ═══════════════════════════════════════════════════════════════════════════


def parse_markdown_text(
    content: str,
    *,
    auto_link_bare_url: bool = False,
) -> list[dict[str, Any]]:
    """Parse Markdown text into an AST using the extended mistune parser.

    Args:
        content: Raw Markdown text.

    Returns:
        A list of token dicts from mistune's AST renderer.
    """
    parser = create_extended_markdown(auto_link_bare_url=auto_link_bare_url)
    return parser(content)  # pyright: ignore[reportReturnType]
