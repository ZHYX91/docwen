"""Immutable edit baselines for template-backed XLSX workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Any


def _value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool, date, datetime, time, timedelta)):
        return value
    if isinstance(value, (list, tuple)):
        return tuple(_value(item) for item in value)
    if isinstance(value, dict):
        return tuple(sorted((key, _value(item)) for key, item in value.items()))
    return (type(value).__name__, _value(vars(value)))


def _style(owner: Any) -> tuple[int, ...] | None:
    return tuple(owner._style) if owner._style is not None else None


def _comment(cell: Any) -> Any:
    comment = cell.comment
    if comment is None:
        return None
    return (comment.text, comment.author, comment.width, comment.height)


def _hyperlink(cell: Any) -> Any:
    link = cell.hyperlink
    if link is None:
        return None
    return (link.target, link.location, link.tooltip, link.display)


@dataclass(frozen=True)
class CellState:
    value: Any
    data_type: str
    style: tuple[int, ...] | None
    comment: Any
    hyperlink: Any

    @classmethod
    def capture(cls, cell: Any) -> CellState:
        return cls(_value(cell.value), cell.data_type, _style(cell), _comment(cell), _hyperlink(cell))


ROW_FIELDS = ("height", "hidden", "outlineLevel", "collapsed", "thickTop", "thickBot")
COLUMN_FIELDS = ("width", "hidden", "outlineLevel", "collapsed", "bestFit", "min", "max")


def dimension_state(dimension: Any, fields: tuple[str, ...]) -> dict[str, Any]:
    return {**{name: getattr(dimension, name, None) for name in fields}, "style": _style(dimension)}


@dataclass(frozen=True)
class SheetState:
    cells: dict[str, CellState]
    rows: dict[int, dict[str, Any]]
    columns: dict[str, dict[str, Any]]
    merges: frozenset[str]
    image_count: int
    title: str
    sheet_state: str
    properties: dict[str, bytes]

    @classmethod
    def capture(cls, sheet: Any) -> SheetState:
        from lxml import etree

        properties: dict[str, bytes] = {}
        for name in ("sheet_format", "sheet_properties", "views", "page_margins", "page_setup", "print_options"):
            value = getattr(sheet, name)
            if value is not None:
                properties[name] = etree.tostring(value.to_tree())
        return cls(
            {cell.coordinate: CellState.capture(cell) for cell in sheet._cells.values()},
            {key: dimension_state(value, ROW_FIELDS) for key, value in sheet.row_dimensions.items()},
            {key: dimension_state(value, COLUMN_FIELDS) for key, value in sheet.column_dimensions.items()},
            frozenset(str(item) for item in sheet.merged_cells.ranges),
            len(sheet._images),
            sheet.title,
            sheet.sheet_state,
            properties,
        )


@dataclass(frozen=True)
class SheetOrigin:
    part: str
    baseline: SheetState
    clone: bool = False


@dataclass
class TemplateSession:
    parts: dict[str, bytes]
    source_path: str
    origins: dict[Any, SheetOrigin]
    styles: Any
    styles_xml: bytes
    original_sheets: tuple[Any, ...]
    original_sheet_names: dict[Any, str]
