"""Typed cached-cell facts and bounded, coordinate-aware ODF evidence."""

from __future__ import annotations

import re
from bisect import bisect_left
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation

from lxml import etree

_OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
_TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
_TEXT = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
_COORDINATE = re.compile(r"([A-Z]{1,3})([1-9][0-9]{0,6})\Z")


@dataclass(frozen=True, slots=True)
class CachedExternalCell:
    """One existing XLSX cached value; absence is never represented as empty."""

    sheet_name: str
    coordinate: str
    cell_type: str
    value: str

    @property
    def position(self) -> tuple[int, int]:
        match = _COORDINATE.fullmatch(self.coordinate)
        if match is None:
            raise ValueError("Invalid cached-cell coordinate")
        column = 0
        for letter in match[1]:
            column = column * 26 + ord(letter) - ord("A") + 1
        row = int(match[2])
        if column > 16384 or row > 1048576:
            raise ValueError("Cached-cell coordinate is outside the XLSX grid")
        return row, column


def numeric_equal(left: str, right: str) -> bool:
    try:
        a, b = Decimal(left), Decimal(right)
        return a.is_finite() and b.is_finite() and a == b
    except InvalidOperation:
        return False


def _repeat(element: etree._Element, attribute: str, maximum: int) -> int:
    value = element.get(_TABLE + attribute, "1")
    if len(value) > 7 or not value.isascii() or not value.isdecimal():
        raise ValueError("Invalid ODS repetition")
    result = int(value)
    if not 1 <= result <= maximum:
        raise ValueError("ODS repetition exceeds the worksheet grid")
    return result


def _text_content(element: etree._Element) -> str:
    parts = [element.text or ""]
    for child in element:
        if child.tag == _TEXT + "s":
            count = child.get(_TEXT + "c", "1")
            if len(count) > 5 or not count.isascii() or not count.isdecimal() or not 1 <= int(count) <= 32767:
                raise ValueError("Invalid ODS text-space repetition")
            parts.append(" " * int(count))
        elif child.tag == _TEXT + "tab":
            parts.append("\t")
        elif child.tag == _TEXT + "line-break":
            parts.append("\n")
        else:
            parts.append(_text_content(child))
        parts.append(child.tail or "")
    return "".join(parts)


def _matches(element: etree._Element, expected: CachedExternalCell) -> bool:
    value_type = element.get(_OFFICE + "value-type")
    if element.tag == _TABLE + "covered-table-cell":
        return False
    if expected.cell_type == "str":
        if value_type not in {None, "string"}:
            return False
        if any(
            element.get(_OFFICE + name) is not None for name in ("value", "boolean-value", "date-value", "time-value")
        ):
            return False
        if element.get(_TABLE + "formula") is not None:
            return False
        value = element.get(_OFFICE + "string-value")
        if value is None:
            value = "\n".join(_text_content(p) for p in element.findall(_TEXT + "p"))
        # A physically represented untyped blank is the ODF empty-string
        # mapping. An absent cell, covered cell or typed number is not.
        return value == expected.value and (value_type == "string" or not value)
    if expected.cell_type == "b":
        if value_type != "boolean":
            return False
        value = element.get(_OFFICE + "boolean-value")
        if value not in {"true", "false", "1", "0"}:
            return False
        return (value in {"true", "1"}) == (expected.value == "1")
    if expected.cell_type == "n":
        value = element.get(_OFFICE + "value")
        return (
            value_type in {"float", "currency", "percentage"}
            and value is not None
            and numeric_equal(value, expected.value)
        )
    return False


@dataclass(slots=True)
class _SheetCursor:
    name: str
    table_element: etree._Element
    row: int = 1
    row_repeat: int = 1
    column: int = 1
    row_element: etree._Element | None = None
    expected_columns: tuple[int, ...] = ()
    expected_cells: tuple[CachedExternalCell, ...] = ()


class OdsCachedCellVerifier:
    """Stream owned ODF rows/repeats without expanding grids or rescanning rows."""

    def __init__(self, expected: tuple[CachedExternalCell, ...]) -> None:
        self._by_sheet: dict[str, dict[int, list[tuple[int, CachedExternalCell]]]] = {}
        for cell in expected:
            row, column = cell.position
            self._by_sheet.setdefault(cell.sheet_name, {}).setdefault(row, []).append((column, cell))
        self._row_numbers: dict[str, list[int]] = {}
        for name, rows in self._by_sheet.items():
            self._row_numbers[name] = sorted(rows)
            for cells in rows.values():
                cells.sort(key=lambda item: item[0])
                columns = [column for column, _cell in cells]
                if len(set(columns)) != len(columns):
                    raise ValueError("Duplicate cached-cell identity")
        self._expected_count = len(expected)
        self._matched = 0
        self._seen_sheets: set[str] = set()
        self._cursor: _SheetCursor | None = None
        self._table_depth = 0
        self._cell_depth = 0
        self._spreadsheet_seen = False

    @property
    def inside_cell(self) -> bool:
        return self._cell_depth > 0

    @property
    def complete(self) -> bool:
        return self._spreadsheet_seen and self._matched == self._expected_count

    def _begin_row(self, element: etree._Element, cursor: _SheetCursor) -> None:
        parent = element.getparent()
        row_containers = {_TABLE + "table-row-group", _TABLE + "table-header-rows", _TABLE + "table-rows"}
        while parent is not None and parent.tag in row_containers:
            parent = parent.getparent()
        if parent is not cursor.table_element or cursor.row_element is not None:
            raise ValueError("ODS row has no worksheet ownership")
        cursor.row_element = element
        cursor.row_repeat = _repeat(element, "number-rows-repeated", 1048576)
        cursor.column = 1
        rows = self._by_sheet.get(cursor.name, {})
        row_numbers = self._row_numbers.get(cursor.name, [])
        lo = bisect_left(row_numbers, cursor.row)
        hi = bisect_left(row_numbers, cursor.row + cursor.row_repeat)
        # Only existing expected cells enter this row-span projection. Column
        # lookups then visit each owned fact once, even for a repeated row.
        facts = sorted(
            (item for row in row_numbers[lo:hi] for item in rows[row]),
            key=lambda item: item[0],
        )
        cursor.expected_columns = tuple(column for column, _cell in facts)
        cursor.expected_cells = tuple(cell for _column, cell in facts)

    def observe(self, event: str, element: etree._Element) -> None:
        tag = element.tag
        if event == "start" and element.getparent() is None and tag != _OFFICE + "document-content":
            raise ValueError("ODS content has no document-content root")
        if tag == _OFFICE + "spreadsheet" and event == "start":
            body = element.getparent()
            root = body.getparent() if body is not None else None
            if (
                body is None
                or body.tag != _OFFICE + "body"
                or root is None
                or root.tag != _OFFICE + "document-content"
                or root.getparent() is not None
                or self._spreadsheet_seen
            ):
                raise ValueError("ODS spreadsheet has no unique document body ownership")
            self._spreadsheet_seen = True
        if tag == _TABLE + "table":
            if event == "start":
                self._table_depth += 1
                parent = element.getparent()
                if self._table_depth == 1 and parent is not None and parent.tag == _OFFICE + "spreadsheet":
                    name = element.get(_TABLE + "name", "")
                    if name in self._seen_sheets and name in self._by_sheet:
                        raise ValueError("Ambiguous ODS sheet identity")
                    self._seen_sheets.add(name)
                    self._cursor = _SheetCursor(name, element)
            else:
                self._table_depth -= 1
                if self._table_depth == 0:
                    self._cursor = None
            return
        if tag in {_TABLE + "table-cell", _TABLE + "covered-table-cell"}:
            self._cell_depth += 1 if event == "start" else -1
        cursor = self._cursor
        if self._table_depth != 1 or cursor is None:
            return
        if tag == _TABLE + "table-row":
            if event == "start":
                self._begin_row(element, cursor)
            else:
                if element is not cursor.row_element:
                    raise ValueError("ODS row ended outside its worksheet ownership")
                cursor.row += cursor.row_repeat
                cursor.row_element = None
                cursor.expected_columns = ()
                cursor.expected_cells = ()
        elif tag in {_TABLE + "table-cell", _TABLE + "covered-table-cell"} and event == "end":
            if cursor.row_element is None or element.getparent() is not cursor.row_element:
                raise ValueError("ODS cell has no row ownership")
            repeat = _repeat(element, "number-columns-repeated", 16384)
            lo = bisect_left(cursor.expected_columns, cursor.column)
            hi = bisect_left(cursor.expected_columns, cursor.column + repeat)
            for expected in cursor.expected_cells[lo:hi]:
                if not _matches(element, expected):
                    raise ValueError(f"Cached value/type changed at {expected.sheet_name}!{expected.coordinate}")
                self._matched += 1
            cursor.column += repeat
