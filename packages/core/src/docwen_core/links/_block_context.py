"""Source-preserving CommonMark container prefixes shared by source consumers."""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from itertools import pairwise

from docwen_core.links._markdown_inline import _is_backslash_escaped, matching_backtick_run_end

_QUOTE = re.compile(r" {0,3}>[ \t]?")
_LIST = re.compile(r" {0,3}(?P<marker>[-+*]|[0-9]{1,9}[.)])(?P<padding>[ \t]+)")
THEMATIC_BREAK_RE = re.compile(r" {0,3}(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,})$")
ATX_HEADING_RE = re.compile(r" {0,3}(#{1,9})(?!#)(?:[ \t]+|$)")
SETEXT_HEADING_RE = re.compile(r" {0,3}(=+|-+)[ \t]*$")
TABLE_DELIMITER_RE = re.compile(r" {0,3}\|?[ \t]*:?-+:?[ \t]*(?:\|[ \t]*:?-+:?[ \t]*)*\|?[ \t]*$")
_YAML_OPEN = re.compile(r"(?:\ufeff)?---[ \t]*(?:\r\n|\r|\n)")
_YAML_CLOSE = re.compile(r"---[ \t]*(?:\r\n|\r|\n|$)")


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


@dataclass(frozen=True, slots=True)
class MarkdownBlockLine:
    content: str
    content_start: int
    containers: tuple[_Container, ...]
    lazy: bool = False


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

    Fenced blocks cannot use paragraph lazy continuation. Callers disable new
    containers while reading a literal block so its sample markers stay data.
    """

    def __init__(self, *, extended_headings: bool = False) -> None:
        self._extended_headings = extended_headings
        self._containers: tuple[_Container, ...] = ()
        self._next_identity = 0
        self._paragraph_open = False
        self._previous_content = ""
        self._table_columns = 0

    def line(self, line: str, *, allow_new: bool = True, allow_lazy: bool = False) -> MarkdownBlockLine:
        offset = 0
        active: list[_Container] = []
        for container in self._containers:
            if container.kind == "quote":
                match = _QUOTE.match(line, offset)
                if match is None:
                    break
                offset = match.end()
            else:
                end = _indent_end(line, offset, container.indent)
                if end is None:
                    if line[offset:].strip():
                        break
                    end = len(line.rstrip("\r\n"))
                offset = end
            active.append(container)

        lazy = False
        remainder = line[offset:].rstrip("\r\n")
        if (
            allow_lazy
            and self._paragraph_open
            and len(active) < len(self._containers)
            and remainder.strip()
            and not match_atx_heading(remainder, extended_headings=self._extended_headings)
            and not _QUOTE.match(remainder)
            and not _LIST.match(remainder)
            and not THEMATIC_BREAK_RE.fullmatch(remainder)
            and not re.match(r" {0,3}(?:`{3,}|~{3,})", remainder)
        ):
            active = list(self._containers)
            lazy = True

        if allow_new:
            while offset < len(line):
                quote = _QUOTE.match(line, offset)
                if quote is not None:
                    active.append(self._new_container("quote", 0))
                    offset = quote.end()
                    continue
                if THEMATIC_BREAK_RE.fullmatch(line[offset:].rstrip("\r\n")):
                    break
                item = _LIST.match(line, offset)
                if item is None:
                    break
                end = item.end()
                whitespace = item.group("padding")
                # More than four spaces after a marker leave an indented code
                # prefix after the one mandatory separator.
                if len(whitespace.expandtabs(4)) > 4:
                    end = item.start("padding") + 1
                indent = len(line[offset:end].expandtabs(4))
                marker = item.group("marker")
                active.append(self._new_container("list", indent, marker[-1]))
                offset = end

        containers = tuple(active)
        content = line[offset:].rstrip("\r\n")
        if containers != self._containers:
            self._table_columns = 0
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
        self._paragraph_open = bool(
            content.strip()
            and not self._table_columns
            and not match_atx_heading(content, extended_headings=self._extended_headings)
            and not THEMATIC_BREAK_RE.fullmatch(content)
            and not re.match(r" {0,3}(?:`{3,}|~{3,})", content)
        )
        self._previous_content = content
        self._containers = containers
        return MarkdownBlockLine(line[offset:], offset, containers, lazy)

    def _new_container(self, kind: str, indent: int, marker: str = "") -> _Container:
        self._next_identity += 1
        return _Container(kind, indent, self._next_identity, marker)


def markdown_inline_regions(
    source: str, *, structural_tables: bool = False, extended_headings: bool = False
) -> list[tuple[int, int]]:
    """Inline parser scopes after opaque blocks have been masked in place.

    Blank lines, new container items, ATX headings, thematic breaks and table
    cells end an inline scope. Setext titles keep their paragraph scope.
    """
    boundaries = {0, len(source)}
    context = MarkdownBlockContext(extended_headings=extended_headings)
    previous_containers: tuple[object, ...] = ()
    offset = 0
    rows: list[tuple[int, int, MarkdownBlockLine]] = []
    for line in source.splitlines(keepends=True):
        row = context.line(line, allow_lazy=True)
        text = row.content.rstrip("\r\n")
        end = offset + len(line)
        if row.containers != previous_containers:
            boundaries.add(offset)
        if (
            not text.strip()
            or match_atx_heading(text, extended_headings=extended_headings)
            or THEMATIC_BREAK_RE.fullmatch(text)
        ):
            boundaries.update((offset, end))
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
