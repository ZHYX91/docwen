"""Extended mistune Markdown parser with custom inline plugins.

Uses mistune v3's plugin API (``md.inline.register``). The built-in
formatting plugins from ``mistune.plugins.formatting`` provide:

- ``==highlight==``  — via plugin ``mark`` (``<mark>`` element)
- ``^^insert^^``    — via plugin ``insert`` (``<ins>`` element)
- ``^superscript^`` — via plugin ``superscript``
- ``~subscript~``   — via plugin ``subscript``

The inline HTML formatting plugin accepts the finite set of tags emitted by
the DOCX Markdown writer; arbitrary HTML remains literal content.

The module exposes ``parse_markdown_text()`` as the Markdown plugin's parser
entry point.
"""

from __future__ import annotations

import re
from functools import partial
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
from mistune.plugins.table import table_in_list, table_in_quote

from docwen_core.links import (
    is_markdown_source_escaped,
    markdown_source_owners,
    markdown_table_cell_ranges,
    markdown_table_delimiter,
)
from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_core.markdown_math import DISPLAY_MATH_START, match_formula
from docwen_plugin_markdown.callouts import plugin_callout_headers

_EXTENDED_ATX_HEADING_TRIM = re.compile(r"(\s+|^)#+\s*$")
_STRUCTURAL_TABLE_BLOCK = (
    r"^ {0,3}(?P<structural_table_rows>"
    r"(?:[^\n]*\|[^\n]*(?:\n|$)){2,})"
)

# ═══════════════════════════════════════════════════════════════════════════
# Custom plugin: semantic inline HTML formatting
# ═══════════════════════════════════════════════════════════════════════════


_HTML_FORMAT_TYPES = {
    "u": "underline",
    "sub": "subscript",
    "sup": "superscript",
    "del": "strikethrough",
    "mark": "mark",
    "strong": "strong",
    "em": "emphasis",
}
_HTML_FORMAT_TAG = r"(?i:<(?:u|sub|sup|del|mark|strong|em)>)"
_HTML_FORMAT_BOUNDARY = re.compile(r"(?i:</?(?:u|sub|sup|del|mark|strong|em)>)")


def _html_format_end(source: str, start: int, tag: str, protected: list[tuple[int, int]]) -> tuple[int, int] | None:
    """Balance tags without consuming code, formulas or link metadata."""
    stack = [tag]
    position = start
    owner_index = 0
    while match := _HTML_FORMAT_BOUNDARY.search(source, position):
        token = match.group(0)
        position = match.end()
        while owner_index < len(protected) and protected[owner_index][1] <= match.start():
            owner_index += 1
        if owner_index < len(protected) and protected[owner_index][0] <= match.start():
            position = protected[owner_index][1]
            continue
        if is_markdown_source_escaped(source, match.start()):
            continue
        name = token.strip("</>").lower()
        if token.startswith("</"):
            if name != stack[-1]:
                return None
            stack.pop()
            if not stack:
                return match.start(), match.end()
        else:
            stack.append(name)
            # Bound recursive inline parsing just as Mistune bounds emphasis.
            if len(stack) > 20:
                return None
    return None


def plugin_inline_html_formatting(md: mistune.Markdown) -> None:
    """Parse supported HTML tags before generic ``inline_html`` consumes them.

    Named groups are required: Mistune combines plugin patterns into one
    scanner. A numeric group is not local to this rule. The explicit outer
    emphasis form also accepts ``中**<u>文</u>**中`` without treating the
    semantic HTML container as punctuation at the emphasis boundary.
    """

    def parse_html_format(inline, match, state):
        opening = match.group("docwen_html_format_open")
        tag = opening[1:-1].lower()
        # Keep source ownership in the existing Core lexer rather than adding
        # separate approximations for code, math and link destinations here.
        ownership = state.env.setdefault("docwen_html_format_owners", {})
        if state.src not in ownership:
            ownership[state.src] = [
                (owner.start, owner.end)
                for owner in markdown_source_owners(state.src)
                if not (owner.kind == "literal" and _HTML_FORMAT_BOUNDARY.fullmatch(state.src[owner.start : owner.end]))
            ]
        end = _html_format_end(state.src, match.end(), tag, ownership[state.src])
        if end is None:
            return None
        content_end, token_end = end
        marker = match.group("docwen_html_format_emphasis")
        if marker:
            if not state.src.startswith(marker, token_end):
                return None
            token_end += len(marker)
        child_state = state.copy()
        child_state.src = state.src[match.end() : content_end]
        token = {"type": _HTML_FORMAT_TYPES[tag], "children": inline.render(child_state)}
        if marker:
            if len(marker) >= 2:
                token = {"type": "strong", "children": [token]}
            if len(marker) % 2:
                token = {"type": "emphasis", "children": [token]}
        state.append_token(token)
        return token_end

    md.inline.register(
        "docwen_html_format",
        rf"(?P<docwen_html_format_emphasis>\*{{1,3}}|_{{1,3}})?(?P<docwen_html_format_open>{_HTML_FORMAT_TAG})",
        parse_html_format,
        before="emphasis",
    )


# ═══════════════════════════════════════════════════════════════════════════
# Custom plugin: shared dollar-math grammar
# ═══════════════════════════════════════════════════════════════════════════


def plugin_math(md: mistune.Markdown) -> None:
    """Use the source scanner's math boundaries in every Markdown container."""

    def append_formula(state, formula):
        state.append_token(
            {
                "type": "inline_math" if formula["is_inline"] else "block_math",
                "raw": formula["latex"],
                "_math_source": state.src[formula["start"] : formula["end"]],
            }
        )

    def parse_block_math(block, match, state):
        start = match.end() - 2
        formula = match_formula(state.src, start, standalone=True)
        if formula is None:
            return None
        append_formula(state, formula)
        end = formula["end"]
        line_end = state.src.find("\n", end)
        return len(state.src) if line_end < 0 else line_end + 1

    def parse_inline_math(inline, match, state):
        formula = match_formula(state.src, match.start())
        if formula is None:
            return None
        append_formula(state, formula)
        return formula["end"]

    md.block.register("block_math", DISPLAY_MATH_START.pattern, parse_block_math, before="list")
    md.inline.register("inline_math", r"\$", parse_inline_math, before="codespan")
    for rules_name in ("block_quote_rules", "list_rules"):
        rules = getattr(md.block, rules_name)
        if "block_math" not in rules:
            md.block.insert_rule(rules, "block_math", before="list")


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
    # Container parses have their own rule lists.  Register the extension in
    # each list at the same precedence as its top-level counterpart so source
    # heading ownership and the native AST agree at every supported depth.
    for rules_name in ("block_quote_rules", "list_rules"):
        rules = getattr(md.block, rules_name)
        if "extended_atx_heading" not in rules:
            md.block.insert_rule(rules, "extended_atx_heading", before="atx_heading")


def _structural_pipe_row(line: str) -> list[str] | None:
    """Use the same source cell ownership as request-scoped link processing."""

    if "|" not in line:
        return None
    ranges = markdown_table_cell_ranges(line, structural_tables=True)
    return [line[start:end] for start, end in ranges] or None


def _structural_delimiter(line: str) -> tuple[list[str | None], int] | None:
    return markdown_table_delimiter(line, structural_tables=True)


def _structural_table_cell(text: str, alignment: str | None, *, head: bool) -> dict[str, Any]:
    stripped = text.strip()
    attrs: dict[str, Any] = {
        "align": alignment,
        "head": head,
        "docwen_literal_merge_marker": stripped in {r"\<", r"\^"},
    }
    if stripped in {"<", "^"}:
        attrs["docwen_merge_marker"] = stripped
    return {
        "type": "table_cell",
        "text": stripped,
        "attrs": attrs,
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
        content_rows = [*rows[:delimiter_index], *rows[delimiter_index + 1 :]]
        if any(row is None or len(row) != column_count for row in content_rows):
            return None
        concrete_rows = [row for row in content_rows if row is not None]
        marker_found = any(cell.strip() in {"<", "^"} for row in concrete_rows for cell in row)
        structural = marker_found or delimiter_index != 1 or header_columns > 0

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
    for nested_rules_name in ("block_quote_rules", "list_rules"):
        nested_rules = getattr(md.block, nested_rules_name, None)
        if isinstance(nested_rules, list) and "structural_table" not in nested_rules:
            md.block.insert_rule(nested_rules, "structural_table", before="table")


def plugin_obsidian_comment_blocks(md: mistune.Markdown, *, hide_content: bool) -> None:
    """Protect standalone comment blocks before table parsing at every depth."""

    def parse_literal_comment(block, match, state):
        if hide_content:
            state.append_token({"type": "blank_line"})
        else:
            state.append_token(
                {"type": "paragraph", "children": [{"type": "text", "raw": match.group(0).rstrip("\n")}]}
            )
        return match.end()

    md.block.register(
        "obsidian_comment",
        r"^ {0,3}%%[ \t]*\n[\s\S]*?(?:^ {0,3}%%[ \t]*(?:\n|$)|\Z)",
        parse_literal_comment,
        before="table",
    )
    for rules_name in ("block_quote_rules", "list_rules"):
        rules = getattr(md.block, rules_name, None)
        if isinstance(rules, list) and "obsidian_comment" not in rules:
            rules.insert(0, "obsidian_comment")


# ═══════════════════════════════════════════════════════════════════════════
# Factory
# ═══════════════════════════════════════════════════════════════════════════


def plugin_source_breaks(md: mistune.Markdown) -> None:
    """Bind break metadata while the parser still owns the matched source.

    Literal code and Setext underlines never reach the thematic-break rule.
    An HTML break is inline content, so it must not invent a physical line
    that could split a heading, table cell, or link before block parsing.
    """

    def parse_thematic_break(block, match, state):
        end = block.parse_thematic_break(match, state)
        token = state.last_token()
        token["_hr_marker"] = {"-": "dash", "*": "asterisk", "_": "underscore"}[match.group(0).lstrip()[0]]
        previous_end = max(0, match.start() - 1)
        previous_start = state.src.rfind("\n", 0, previous_end) + 1
        adjacent = bool(state.src[previous_start:previous_end].strip())
        if adjacent:
            token["_attach_to_prev"] = True
        return end

    def bind_attachments(markdown, state):
        def visit(nodes):
            previous = None
            for node in nodes:
                # Lists can be inserted before the rule only after their child
                # parser returns. Inspect the completed sibling list here.
                if node.get("type") == "thematic_break" and (
                    previous is None or previous.get("type") not in {"paragraph", "list"}
                ):
                    node.pop("_attach_to_prev", None)
                visit(node.get("children", []))
                previous = node

        visit(state.tokens)

    def parse_html_linebreak(inline, match, state):
        state.append_token({"type": "linebreak"})
        return match.end()

    md.block.register("thematic_break", None, parse_thematic_break)
    md.before_render_hooks.append(bind_attachments)
    md.inline.register("html_linebreak", r"(?i:<br[ \t]*/?>)", parse_html_linebreak, before="inline_html")


def create_extended_markdown(
    *, auto_link_bare_url: bool = False, extensions: MarkdownExtensions | None = None
) -> mistune.Markdown:
    """Create a mistune Markdown instance with the requested plugins.

    Includes:
    - Built-in: table, footnotes, strikethrough, task_lists
    - Formatting: mark (``==text==``), insert (``^^text^^``),
      superscript (``^text^``), subscript (``~text~``)
    - Custom: inline HTML formats emitted by the DOCX Markdown writer

    Returns:
        A mistune ``Markdown`` instance configured with ``renderer="ast"``
        for structured output.
    """
    dialect = extensions or MarkdownExtensions.obsidian()
    plugins: list[Any] = [
        "table",
        table_in_quote,
        table_in_list,
        "footnotes",
        "strikethrough",
        "task_lists",
        plugin_math,
        _mark,
        _insert,
        _superscript,
        _subscript,
        plugin_inline_html_formatting,
        plugin_source_breaks,
        plugin_callout_headers,
        partial(plugin_obsidian_comment_blocks, hide_content=dialect.structural_tables),
    ]
    if dialect.structural_tables:
        plugins.append(plugin_structural_tables)
    if dialect.extended_headings:
        plugins.append(plugin_extended_atx_headings)
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
    extensions: MarkdownExtensions | None = None,
    consume_initial_bom: bool = True,
) -> list[dict[str, Any]]:
    """Parse Markdown text into an AST using the extended mistune parser.

    Args:
        content: Raw Markdown text.
        consume_initial_bom: Consume one document encoding signature. Body,
            label and note fragments must pass False after their document's
            initial signature has already been handled.

    Returns:
        A list of token dicts from mistune's AST renderer.
    """
    parser = create_extended_markdown(auto_link_bare_url=auto_link_bare_url, extensions=extensions)
    # Source ownership and numbering keep exact authored coordinates.  Only a
    # complete document can authorize consuming an initial encoding signature;
    # a fragment's leading U+FEFF is ordinary content.
    nodes = parser(content.removeprefix("\ufeff") if consume_initial_bom else content)
    if extensions is not None and not extensions.structural_tables:

        def literal_cells(items: list[dict[str, Any]]) -> None:
            for node in items:
                if node.get("type") == "table_cell":
                    node.setdefault("attrs", {})["docwen_literal_merge_marker"] = True
                literal_cells(node.get("children", []))

        literal_cells(nodes)  # pyright: ignore[reportArgumentType]
    return nodes  # pyright: ignore[reportReturnType]
