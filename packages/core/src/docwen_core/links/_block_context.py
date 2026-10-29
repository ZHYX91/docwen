"""Source-preserving CommonMark container prefixes shared by source consumers."""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from itertools import pairwise

from docwen_core.links._markdown_inline import _INLINE_HTML_RE, _is_backslash_escaped, matching_backtick_run_end

_QUOTE = re.compile(r" {0,3}>[ \t]?")
_LIST = re.compile(r" {0,3}(?P<marker>[-+*]|[0-9]{1,9}[.)])(?P<padding>[ \t]+|$)")
_FENCED_CODE_START = re.compile(r" {0,3}(`{3,}|~{3,})([^\r\n]*)")
# The production Mistune BlockParser stops descending after twenty containers.
_MAX_CONTAINER_DEPTH = 20
THEMATIC_BREAK_RE = re.compile(r" {0,3}(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,})$")
ATX_HEADING_RE = re.compile(r" {0,3}(#{1,9})(?!#)(?:[ \t]+|$)")
SETEXT_HEADING_RE = re.compile(r" {0,3}(=+|-+)[ \t]*$")
TABLE_DELIMITER_RE = re.compile(r" {0,3}\|?[ \t]*:?-+:?[ \t]*(?:\|[ \t]*:?-+:?[ \t]*)*\|?[ \t]*$")
_YAML_OPEN = re.compile(r"(?:\ufeff)?---[ \t]*(?:\r\n|\r|\n)")
_YAML_CLOSE = re.compile(r"---[ \t]*(?:\r\n|\r|\n|$)")


_HTML_PRE_TAGS = frozenset({"pre", "script", "style", "textarea"})
_HTML_BLOCK_TAGS = frozenset(
    {
        "address",
        "article",
        "aside",
        "base",
        "basefont",
        "blockquote",
        "body",
        "caption",
        "center",
        "col",
        "colgroup",
        "dd",
        "details",
        "dialog",
        "dir",
        "div",
        "dl",
        "dt",
        "fieldset",
        "figcaption",
        "figure",
        "footer",
        "form",
        "frame",
        "frameset",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "head",
        "header",
        "hr",
        "html",
        "iframe",
        "legend",
        "li",
        "link",
        "main",
        "menu",
        "menuitem",
        "meta",
        "nav",
        "noframes",
        "ol",
        "optgroup",
        "option",
        "p",
        "param",
        "section",
        "source",
        "summary",
        "table",
        "tbody",
        "td",
        "tfoot",
        "th",
        "thead",
        "title",
        "tr",
        "track",
        "ul",
    }
)


def html_block_mode(
    line: str,
    *,
    allow_type7: bool,
) -> tuple[str | None, bool] | None:
    """Return ``(end marker, until blank)`` for a Mistune HTML block start."""
    leading = re.match(r"^ {0,3}", line)
    assert leading is not None
    content = line[leading.end() :]
    if content.startswith("<!--"):
        return "-->", False
    if content.startswith("<?"):
        return "?>", False
    if content.startswith("<![CDATA["):
        return "]]>", False
    if re.match(r"<![A-Z]", content):
        return ">", False

    match = re.match(
        r"</?([A-Za-z][A-Za-z0-9-]*)\b",
        content,
        re.IGNORECASE,
    )
    if match is None:
        return None
    marker = match.group(0).lower()
    tag = (match.group(1) or "").lower()
    if tag in _HTML_PRE_TAGS and not marker.startswith("</"):
        return f"</{tag}>", False
    if tag in _HTML_BLOCK_TAGS or tag in _HTML_PRE_TAGS:
        return None, True
    candidate = line.lstrip(" ").rstrip(" \t\r\n")
    generic_self_closing = candidate.startswith("<") and not candidate.startswith("</") and candidate.endswith("/>")
    if allow_type7 and not generic_self_closing and _INLINE_HTML_RE.fullmatch(candidate) is not None:
        return None, True
    return None


def match_atx_heading(line: str, *, extended_headings: bool = False) -> re.Match[str] | None:
    """Match an actual ATX boundary in the frozen input heading dialect."""
    match = ATX_HEADING_RE.match(line)
    return match if match is not None and (extended_headings or len(match.group(1)) <= 6) else None


def yaml_front_matter_end(source: str) -> int:
    """Return a complete front matter prefix length, without normalizing EOLs."""
    opening = _YAML_OPEN.match(source)
    if opening is None:
        return 0
    offset = opening.end()
    for line in source[offset:].splitlines(keepends=True):
        if _YAML_CLOSE.fullmatch(line):
            return offset + len(line)
        offset += len(line)
    return 0


@dataclass(frozen=True, slots=True)
class _Container:
    kind: str
    indent: int
    identity: int
    marker: str = ""
    requires_marker: bool = False


@dataclass(frozen=True, slots=True)
class MarkdownBlockLine:
    content: str
    content_start: int
    containers: tuple[_Container, ...]
    lazy: bool = False
    block_kind: str = "paragraph"
    paragraph_continuation: bool = False


def _starts_fenced_code(line: str) -> bool:
    match = _FENCED_CODE_START.match(line)
    return bool(match and (match.group(1)[0] == "~" or "`" not in match.group(2)))


def _list_interrupts_paragraph(match: re.Match[str], line: str) -> bool:
    """A list needs a nonempty first item, and ordered lists must start at one."""
    marker = match.group("marker")
    return bool(line[match.end() :].strip() and (not marker[0].isdigit() or int(marker[:-1]) == 1))


def _indent_end(line: str, start: int, columns: int) -> int | None:
    consumed = 0
    offset = start
    while offset < len(line) and consumed < columns:
        if line[offset] == " ":
            consumed += 1
        elif line[offset] == "\t":
            consumed += 4 - consumed % 4
        else:
            return None
        offset += 1
    return offset if consumed >= columns else None


class MarkdownBlockContext:
    """Track actual list/quote owners, including exits of unclosed code fences.

    Container extraction and paragraph eligibility are distinct lifecycles.
    In particular, a list may own unindented literal payload until a blank or
    interrupting block ends the item. Such payload never opens a paragraph.
    Callers disable new containers inside literals so sample markers stay data.
    """

    def __init__(self, *, extended_headings: bool = False) -> None:
        self._extended_headings = extended_headings
        self._containers: tuple[_Container, ...] = ()
        self._next_identity = 0
        self._paragraph_open = False
        self._previous_content = ""
        self._table_columns = 0
        self._blank_containers: set[int] = set()

    def line(
        self,
        line: str,
        *,
        allow_new: bool = True,
        allow_lazy: bool = False,
        literal: bool = False,
    ) -> MarkdownBlockLine:
        offset = 0
        active: list[_Container] = []
        blank_containers: set[int] = set()
        lazy = False
        quote_lazy = False
        for container in self._containers:
            remainder = line[offset:].rstrip("\r\n")
            if container.kind == "quote":
                match = _QUOTE.match(line, offset)
                if match is not None:
                    offset = match.end()
                elif (
                    allow_lazy
                    and not container.requires_marker
                    and container.identity not in self._blank_containers
                    and remainder.strip()
                    and not self._container_interrupts(remainder, quote=True)
                ):
                    lazy = quote_lazy = True
                else:
                    break
            else:
                end = _indent_end(line, offset, container.indent)
                if end is None:
                    if not remainder.strip():
                        end = len(line.rstrip("\r\n"))
                    elif (
                        allow_lazy
                        and container.identity not in self._blank_containers
                        and not self._container_interrupts(remainder, quote=False)
                    ):
                        end = offset
                        lazy = True
                    else:
                        break
                offset = end
            active.append(container)
            if not line[offset:].strip():
                blank_containers.add(container.identity)

        paragraph_continuation = self._paragraph_open and tuple(active) == self._containers
        if allow_new and not literal and not lazy:
            while offset < len(line) and len(active) < _MAX_CONTAINER_DEPTH:
                quote = _QUOTE.match(line, offset)
                if quote is not None:
                    child = line[quote.end() :].rstrip("\r\n")
                    requires_marker = (
                        not child.strip() or _starts_fenced_code(child) or _indent_end(child, 0, 4) is not None
                    )
                    active.append(self._new_container("quote", 0, requires_marker=requires_marker))
                    offset = quote.end()
                    paragraph_continuation = False
                    continue
                remainder = line[offset:].rstrip("\r\n")
                if THEMATIC_BREAK_RE.fullmatch(remainder) or (
                    paragraph_continuation and SETEXT_HEADING_RE.fullmatch(remainder)
                ):
                    break
                item = _LIST.match(remainder)
                if item is None or (paragraph_continuation and not _list_interrupts_paragraph(item, remainder)):
                    break
                end = offset + item.end()
                whitespace = item.group("padding")
                # More than four spaces after a marker leave an indented code
                # prefix after the one mandatory separator.
                if len(whitespace.expandtabs(4)) > 4:
                    end = offset + item.start("padding") + 1
                indent = len(line[offset:end].expandtabs(4)) + int(not whitespace)
                marker = item.group("marker")
                active.append(self._new_container("list", indent, marker[-1]))
                offset = end
                paragraph_continuation = False

        containers = tuple(active)
        content = line[offset:].rstrip("\r\n")
        if containers != self._containers:
            self._table_columns = 0
            self._previous_content = ""
        if (
            TABLE_DELIMITER_RE.fullmatch(content)
            and "|" in self._previous_content
            and len(markdown_table_cell_ranges(content)) == len(markdown_table_cell_ranges(self._previous_content))
        ):
            self._table_columns = len(markdown_table_cell_ranges(content))
        elif self._table_columns and (
            "|" not in content or len(markdown_table_cell_ranges(content)) != self._table_columns
        ):
            self._table_columns = 0
        if literal:
            kind = "literal"
        elif not content.strip():
            kind = "blank"
        elif match_atx_heading(content, extended_headings=self._extended_headings):
            kind = "atx_heading"
        elif paragraph_continuation and not quote_lazy and SETEXT_HEADING_RE.fullmatch(content):
            kind = "setext_heading"
        elif THEMATIC_BREAK_RE.fullmatch(content):
            kind = "thematic_break"
        elif _starts_fenced_code(content):
            kind = "fenced_code"
        elif html_block_mode(content, allow_type7=not paragraph_continuation) is not None:
            kind = "html_block"
        elif not paragraph_continuation and _indent_end(content, 0, 4) is not None:
            kind = "indented_code"
        elif self._table_columns:
            kind = "table"
        else:
            kind = "paragraph"
        self._paragraph_open = kind == "paragraph"
        self._previous_content = content
        self._containers = containers
        self._blank_containers = blank_containers
        return MarkdownBlockLine(line[offset:], offset, containers, lazy, kind, paragraph_continuation)

    def end_paragraph(self) -> None:
        """Record a literal block handled by the caller, never its sample text."""
        self._paragraph_open = False
        self._previous_content = ""

    def _container_interrupts(self, line: str, *, quote: bool) -> bool:
        # These are the production container extraction rules, which differ
        # from paragraph interruption. A sibling ordered item may start at any
        # number; an unmarked quote ATX line remains in that quote's child parse.
        return bool(
            THEMATIC_BREAK_RE.fullmatch(line)
            or _starts_fenced_code(line)
            or _LIST.match(line)
            or html_block_mode(line, allow_type7=False) is not None
            or (
                not quote and (_QUOTE.match(line) or match_atx_heading(line, extended_headings=self._extended_headings))
            )
        )

    def _new_container(self, kind: str, indent: int, marker: str = "", *, requires_marker: bool = False) -> _Container:
        self._next_identity += 1
        return _Container(kind, indent, self._next_identity, marker, requires_marker)


def markdown_inline_regions(
    source: str, *, structural_tables: bool = False, extended_headings: bool = False
) -> list[tuple[int, int]]:
    """Inline parser scopes after opaque blocks have been masked in place.

    Blank lines, new container items, ATX headings, thematic breaks and table
    cells end an inline scope. A Setext title keeps its complete paragraph
    scope, then ends it at the underline rather than joining the next block.
    """
    boundaries = {0, len(source)}
    context = MarkdownBlockContext(extended_headings=extended_headings)
    previous_containers: tuple[object, ...] = ()
    offset = 0
    rows: list[tuple[int, int, MarkdownBlockLine]] = []
    for line in source.splitlines(keepends=True):
        row = context.line(line, allow_lazy=True)
        end = offset + len(line)
        if row.containers != previous_containers:
            boundaries.add(offset)
        if row.block_kind in {"blank", "atx_heading", "thematic_break"}:
            boundaries.update((offset, end))
        elif row.block_kind == "setext_heading":
            boundaries.add(end)
        rows.append((offset, end, row))
        previous_containers = row.containers
        offset = end
    for start, end in markdown_table_blocks(
        [row.content.rstrip("\r\n") for _start, _end, row in rows],
        [row.containers for _start, _end, row in rows],
        structural_tables=structural_tables,
        extended_headings=extended_headings,
    ):
        for offset, line_end, row in rows[start:end]:
            boundaries.update((offset, line_end))
            for cell_start, cell_end in markdown_table_cell_ranges(
                row.content.rstrip("\r\n"), structural_tables=structural_tables
            ):
                boundaries.update((offset + row.content_start + cell_start, offset + row.content_start + cell_end))
    ordered = sorted(boundaries)
    return list(pairwise(ordered))


def markdown_table_cell_ranges(line: str, *, structural_tables: bool = False) -> list[tuple[int, int]]:
    """Original half-open cell ranges in the ordinary pipe-table grammar."""
    start = 0 if structural_tables else len(line) - len(line.lstrip(" \t"))
    end = len(line) if structural_tables else len(line.rstrip(" \t\r\n"))
    if not structural_tables and start < end and line[start] == "|":
        start += 1
    if not structural_tables and start < end and line[end - 1] == "|" and not _is_backslash_escaped(line, end - 1):
        end -= 1
    ranges: list[tuple[int, int]] = []
    cursor = start
    index = start
    while index < end:
        if structural_tables and line[index] == "`" and not _is_backslash_escaped(line, index):
            code_end = matching_backtick_run_end(line[:end], index)
            if code_end is not None:
                index = code_end
                continue
        if line[index] == "|" and not _is_backslash_escaped(line, index):
            ranges.append((cursor, index))
            cursor = index + 1
        index += 1
    ranges.append((cursor, end))
    if structural_tables:
        if ranges and not line[ranges[0][0] : ranges[0][1]].strip():
            ranges.pop(0)
        if ranges and not line[ranges[-1][0] : ranges[-1][1]].strip():
            ranges.pop()
    return ranges


def markdown_table_delimiter(line: str, *, structural_tables: bool = False) -> tuple[list[str | None], int] | None:
    """Shared accepted delimiter shape, including the optional row header boundary."""
    if "|" not in line:
        return None
    cells = [line[start:end] for start, end in markdown_table_cell_ranges(line, structural_tables=structural_tables)]
    empty = [index for index, cell in enumerate(cells) if cell == ""]
    if empty and (not structural_tables or len(empty) > 1 or empty[0] in {0, len(cells) - 1}):
        return None
    boundary = empty[0] if empty else None
    tokens = [cell for index, cell in enumerate(cells) if index != boundary]
    if not tokens or any(re.fullmatch(r":?-+:?", cell.strip()) is None for cell in tokens):
        return None
    alignments: list[str | None] = []
    for cell in tokens:
        token = cell.strip()
        alignments.append(
            "center"
            if token.startswith(":") and token.endswith(":")
            else "left"
            if token.startswith(":")
            else "right"
            if token.endswith(":")
            else None
        )
    return alignments, 0 if boundary is None else boundary


def markdown_table_blocks(
    lines: Sequence[str],
    containers: Sequence[tuple[object, ...]],
    *,
    structural_tables: bool = False,
    extended_headings: bool = False,
) -> list[tuple[int, int]]:
    """Accepted table line intervals; rejected candidates remain paragraphs."""
    blocks: list[tuple[int, int]] = []
    index = 0
    while index < len(lines):
        if "|" not in lines[index] or match_atx_heading(lines[index], extended_headings=extended_headings):
            index += 1
            continue
        end = index + 1
        while (
            end < len(lines)
            and "|" in lines[end]
            and containers[end] == containers[index]
            and not match_atx_heading(lines[end], extended_headings=extended_headings)
        ):
            end += 1
        if structural_tables:
            delimiters = [
                (candidate, shape)
                for candidate in range(index, end)
                if (shape := markdown_table_delimiter(lines[candidate], structural_tables=True)) is not None
            ]
            if len(delimiters) == 1:
                delimiter, (alignments, _boundary) = delimiters[0]
                if end - index >= 2 and all(
                    len(markdown_table_cell_ranges(lines[candidate], structural_tables=True)) == len(alignments)
                    for candidate in range(index, end)
                    if candidate != delimiter
                ):
                    blocks.append((index, end))
        else:
            for header in range(index, end - 1):
                shape = markdown_table_delimiter(lines[header + 1])
                if shape is not None and all(
                    len(markdown_table_cell_ranges(lines[candidate])) == len(shape[0])
                    for candidate in range(header, end)
                ):
                    blocks.append((header, end))
                    break
        index = end
    return blocks
