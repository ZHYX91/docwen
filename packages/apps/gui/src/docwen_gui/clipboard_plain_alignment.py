"""Align Office table text without using HTML's formatting whitespace as body."""

from __future__ import annotations

import csv
import io
import re
from array import array
from bisect import bisect_right
from collections.abc import Sequence
from dataclasses import dataclass, replace

from docwen_core.models.clipboard_document import (
    MAX_CLIPBOARD_TEXT_CODEPOINTS,
    ClipboardBlock,
    ClipboardHardBreak,
    ClipboardImageRef,
    ClipboardParagraph,
    ClipboardTable,
    ClipboardText,
)

_HSPACE = " \u00a0"
_MAX_CANDIDATES = 256


@dataclass(frozen=True, slots=True)
class CanonicalClipboardText:
    original: str
    text: str
    starts: Sequence[int]
    ends: Sequence[int]

    def left(self, index: int) -> int:
        return self.ends[index - 1] if index else 0

    def right(self, index: int) -> int:
        return self.starts[index] if index < len(self.starts) else len(self.original)


def canonical_clipboard_text(value: str) -> CanonicalClipboardText:
    """Comparison-only horizontal folding, retaining tab/newline and raw offsets."""

    characters: list[str] = []
    starts = array("I")
    ends = array("I")

    def field(start: int, end: int) -> None:
        while start < end and value[start] in _HSPACE:
            start += 1
        while end > start and value[end - 1] in _HSPACE:
            end -= 1
        while start < end:
            following = start + 1
            if value[start] in _HSPACE:
                while following < end and value[following] in _HSPACE:
                    following += 1
                characters.append(" ")
            else:
                characters.append(value[start])
            starts.append(start)
            ends.append(following)
            start = following

    start = 0
    index = 0
    while index < len(value):
        if value[index] not in "\t\r\n":
            index += 1
            continue
        field(start, index)
        following = index + 2 if value[index : index + 2] == "\r\n" else index + 1
        characters.append("\t" if value[index] == "\t" else "\n")
        starts.append(index)
        ends.append(following)
        start = following
        index = following
    field(start, len(value))
    return CanonicalClipboardText(value, "".join(characters), starts, ends)


def _paragraph_value(paragraph: ClipboardParagraph) -> str:
    parts: list[str] = []
    for inline in paragraph.inlines:
        if isinstance(inline, ClipboardText):
            # A source newline in normal HTML text is formatting whitespace;
            # an explicit HTML <br> is represented by ClipboardHardBreak.
            parts.append(re.sub(r"[ \t\r\n]+", " ", inline.value))
        elif isinstance(inline, ClipboardHardBreak):
            parts.append("\n")
    return "".join(parts)


@dataclass(frozen=True, slots=True)
class _ParagraphSpan:
    paragraph: ClipboardParagraph
    start: int
    end: int
    images: tuple[tuple[ClipboardImageRef, int], ...]


class _TablePlan:
    def __init__(self, *, grid: bool) -> None:
        self.grid = grid
        self.parts: list[str] = []
        self.length = 0
        self.paragraphs: list[_ParagraphSpan] = []
        self.empty_cells: dict[int, int] = {}

    def append(self, value: str) -> None:
        self.parts.append(value)
        self.length += len(value)

    def blocks(self, blocks: tuple[ClipboardBlock, ...]) -> None:
        for index, block in enumerate(blocks):
            if index:
                self.append("\n")
            if isinstance(block, ClipboardTable):
                self.table(block)
                continue
            value = canonical_clipboard_text(_paragraph_value(block)).text
            start = self.length
            images: list[tuple[ClipboardImageRef, int]] = []
            prefix: list = []
            for inline in block.inlines:
                if isinstance(inline, ClipboardImageRef):
                    offset = len(canonical_clipboard_text(_paragraph_value(ClipboardParagraph(tuple(prefix)))).text)
                    images.append((inline, start + offset))
                else:
                    prefix.append(inline)
            self.append(value)
            self.paragraphs.append(_ParagraphSpan(block, start, self.length, tuple(images)))

    def table(self, table: ClipboardTable) -> None:
        cells = {(cell.row, cell.column): cell for cell in table.cells}
        columns_by_row: dict[int, list[int]] = {}
        for cell in table.cells:
            columns_by_row.setdefault(cell.row, []).append(cell.column)
        for row in range(table.row_count):
            if row:
                self.append("\n")
            columns = range(table.column_count) if self.grid else sorted(columns_by_row.get(row, ()))
            for index, column in enumerate(columns):
                if index:
                    self.append("\t")
                cell = cells.get((row, column))
                if cell is not None:
                    if not cell.blocks:
                        self.empty_cells[id(cell)] = self.length
                    self.blocks(cell.blocks)

    def graft(self, table: ClipboardTable, source: CanonicalClipboardText, start: int) -> ClipboardTable:
        paragraphs: dict[int, ClipboardParagraph] = {}
        for span in self.paragraphs:
            cursor = source.left(start + span.start)
            end = source.right(start + span.end)
            inlines: list = []
            for image, position in span.images:
                position_raw = source.right(start + position)
                if position_raw > end:
                    raise ValueError("Clipboard table image position is invalid.")
                if position_raw > cursor:
                    inlines.append(ClipboardText(source.original[cursor:position_raw]))
                inlines.append(image)
                cursor = position_raw
            if end > cursor:
                inlines.append(ClipboardText(source.original[cursor:end]))
            paragraphs[id(span.paragraph)] = ClipboardParagraph(tuple(inlines))

        def walk(current: ClipboardTable) -> ClipboardTable:
            cells = []
            for cell in current.cells:
                blocks: list[ClipboardBlock] = []
                for block in cell.blocks:
                    blocks.append(walk(block) if isinstance(block, ClipboardTable) else paragraphs[id(block)])
                if not cell.blocks:
                    position = start + self.empty_cells[id(cell)]
                    value = source.original[source.left(position) : source.right(position)]
                    if value:
                        blocks.append(ClipboardParagraph((ClipboardText(value),)))
                cells.append(replace(cell, blocks=tuple(blocks)))
            return replace(current, cells=tuple(cells))

        return walk(table)


@dataclass(frozen=True, slots=True)
class PlainTableMatch:
    start: int
    end: int
    table: ClipboardTable


def _line_boundary(text: str, start: int, end: int) -> bool:
    return (start == 0 or text[start - 1] == "\n") and (end == len(text) or text[end] == "\n")


def _literal_matches(table: ClipboardTable, source: CanonicalClipboardText, minimum: int) -> list[PlainTableMatch]:
    matches: list[PlainTableMatch] = []
    for grid in (False, True):
        plan = _TablePlan(grid=grid)
        plan.table(table)
        signature = "".join(plan.parts)
        if not signature:
            if not source.text and minimum == 0:
                matches.append(PlainTableMatch(0, len(source.original), plan.graft(table, source, 0)))
            continue
        index = source.text.find(signature)
        while index >= 0:
            end = index + len(signature)
            if source.left(index) >= minimum and _line_boundary(source.text, index, end):
                matches.append(PlainTableMatch(source.left(index), source.right(end), plan.graft(table, source, index)))
                if len({(item.start, item.end) for item in matches}) > 1:
                    return matches
            index = source.text.find(signature, index + 1)
    return matches


def _cell_from_plain(blocks: tuple[ClipboardBlock, ...], value: str) -> tuple[ClipboardBlock, ...] | None:
    # A temporary one-cell plan also preserves nested tables and inline images.
    from docwen_core.models.clipboard_document import ClipboardTableCell

    wrapper = ClipboardTable(1, 1, (ClipboardTableCell(0, 0, 1, 1, blocks),))
    source = canonical_clipboard_text(value)
    for grid in (False, True):
        plan = _TablePlan(grid=grid)
        plan.table(wrapper)
        if "".join(plan.parts) == source.text:
            return plan.graft(wrapper, source, 0).cells[0].blocks
    return None


def _tsv_matches(table: ClipboardTable, plain_text: str, minimum: int) -> list[PlainTableMatch]:
    first_cells = sorted((cell for cell in table.cells if cell.row == 0), key=lambda cell: cell.column)
    key = ""
    for cell in first_cells:
        plan = _TablePlan(grid=False)
        plan.blocks(cell.blocks)
        words = "".join(plan.parts).split()
        if words:
            key = words[0]
            break
    candidates: set[int] = set()
    line_starts = array("I", [0])
    line_starts.extend(index + 1 for index, character in enumerate(plain_text) if character == "\n")
    if key:
        index = plain_text.find(key, minimum)
        while index >= 0:
            candidates.add(line_starts[bisect_right(line_starts, index) - 1])
            if len(candidates) > _MAX_CANDIDATES:
                return []
            index = plain_text.find(key, index + len(key))
    else:
        if len(line_starts) > _MAX_CANDIDATES:
            return []
        candidates = set(line_starts)
    matches = []
    for start in sorted(candidates):
        if start < minimum:
            continue
        stream = io.StringIO(plain_text[start:], newline="")
        reader = csv.reader(stream, delimiter="\t", strict=True)
        try:
            rows = [next(reader) for _row in range(table.row_count)]
        except (csv.Error, StopIteration):
            continue
        if any(len(row) != table.column_count for row in rows):
            continue
        anchors = {(cell.row, cell.column): cell for cell in table.cells}
        if any(
            canonical_clipboard_text(value).text
            for row_index, row in enumerate(rows)
            for column, value in enumerate(row)
            if (row_index, column) not in anchors
        ):
            continue
        cells = []
        for cell in table.cells:
            blocks = _cell_from_plain(cell.blocks, rows[cell.row][cell.column])
            if blocks is None:
                break
            cells.append(replace(cell, blocks=blocks))
        else:
            end = start + stream.tell()
            if plain_text[end - 2 : end] == "\r\n":
                end -= 2
            elif plain_text[end - 1 : end] in {"\n", "\r"}:
                end -= 1
            matches.append(PlainTableMatch(start, end, replace(table, cells=tuple(cells))))
            if len(matches) > 1:
                return matches
    return matches


def match_plain_table(table: ClipboardTable, plain_text: str, *, minimum: int = 0) -> PlainTableMatch | None:
    """Accept one unique anchor/full-grid or standard quoted-TSV representation."""

    if len(plain_text) > MAX_CLIPBOARD_TEXT_CODEPOINTS:
        return None
    source = canonical_clipboard_text(plain_text)
    matches = _literal_matches(table, source, minimum)
    if not matches:
        matches = _tsv_matches(table, plain_text, minimum)
    unique: dict[tuple[int, int], PlainTableMatch] = {}
    for match in matches:
        identity = (match.start, match.end)
        previous = unique.get(identity)
        if previous is not None and previous.table != match.table:
            return None
        unique[identity] = match
    return next(iter(unique.values())) if len(unique) == 1 else None


__all__ = ["canonical_clipboard_text", "match_plain_table"]
