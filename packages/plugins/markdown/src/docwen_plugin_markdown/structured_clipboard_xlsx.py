"""XLSX workbook projection helpers for recursive clipboard documents."""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import range_boundaries


@dataclass(slots=True)
class StructuredProjectionWorkbook:
    workbook: Any
    prototype: Any
    order_sheet: Any
    semantics_sheet: Any
    template_backed: bool
    table_titles: dict[str, str] = field(default_factory=dict)


def unique_sheet_name(workbook: Any, requested: str) -> str:
    base = requested[:31] or "Sheet"
    name = base
    index = 2
    while name in workbook.sheetnames:
        suffix = f" {index}"
        name = base[: 31 - len(suffix)] + suffix
        index += 1
    return name


def prepare_projection_workbook(template_path: str | None) -> StructuredProjectionWorkbook:
    """Load a user-selected prototype without deleting its workbook structure."""

    template_backed = bool(template_path)
    workbook = load_workbook(Path(template_path)) if template_backed else Workbook()
    prototype = workbook.worksheets[0]
    if not template_backed:
        prototype.title = "Table Prototype"
    order_sheet = workbook.create_sheet(unique_sheet_name(workbook, "Document Order"), 0)
    semantics_sheet = workbook.create_sheet(unique_sheet_name(workbook, "Table Semantics"), 1)
    if not template_backed:
        prototype.sheet_state = "hidden"
    return StructuredProjectionWorkbook(workbook, prototype, order_sheet, semantics_sheet, template_backed)


def _copy_prototype_runtime_properties(source: Any, target: Any) -> None:
    """Copy worksheet properties that openpyxl 3.1.x copy_worksheet omits."""

    for attribute in (
        "sheet_properties",
        "sheet_format",
        "page_margins",
        "page_setup",
        "print_options",
        "views",
        "HeaderFooter",
        "auto_filter",
    ):
        setattr(target, attribute, copy(getattr(source, attribute)))
    target.print_title_rows = source.print_title_rows
    target.print_title_cols = source.print_title_cols
    target.print_area = source.print_area


def _ranges_intersect_data_area(range_string: str, row_count: int, column_count: int) -> bool:
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise ValueError("prototype merged range must have finite row and column bounds")
    return not (max_row < 1 or min_row > row_count or max_col < 1 or min_col > column_count)


def _clear_projection_data_area(sheet: Any, row_count: int, column_count: int) -> None:
    """Replace values/merges in A1:table-bounds while retaining copied formatting."""

    for merged in tuple(sheet.merged_cells.ranges):
        if _ranges_intersect_data_area(str(merged), row_count, column_count):
            sheet.unmerge_cells(str(merged))
    for row in range(1, row_count + 1):
        for column in range(1, column_count + 1):
            cell = sheet.cell(row, column)
            cell.value = None
            cell.comment = None
            cell.hyperlink = None


def create_table_sheet(
    projection: StructuredProjectionWorkbook,
    *,
    table_id: str,
    requested_name: str,
    row_count: int,
    column_count: int,
) -> Any:
    """Clone the selected prototype and reserve one authoritative table data area."""

    sheet = projection.workbook.copy_worksheet(projection.prototype)
    sheet.title = unique_sheet_name(projection.workbook, requested_name)
    _copy_prototype_runtime_properties(projection.prototype, sheet)
    _clear_projection_data_area(sheet, row_count, column_count)
    projection.table_titles[table_id] = sheet.title
    return sheet


def materialize_order_rows(
    rows: list[tuple[str, str, str, str, str, str, str]],
    table_titles: dict[str, str],
) -> list[tuple[str, str, str, str, str, str, str]]:
    """Replace provisional table labels with actual collision-safe sheet titles."""

    output: list[tuple[str, str, str, str, str, str, str]] = []
    for sequence, kind, text, sheet, parent, anchor, child in rows:
        actual_sheet = table_titles.get(child, sheet) if child else sheet
        output.append((sequence, kind, text, actual_sheet, parent, anchor, child))
    return output


def write_string_matrix(sheet: Any, headers: tuple[str, ...], rows: list[tuple[str, ...]]) -> None:
    for column, value in enumerate(headers, 1):
        cell = sheet.cell(1, column, value)
        cell.data_type = "s"
        cell.number_format = "@"
    for row_index, values in enumerate(rows, 2):
        for column, value in enumerate(values, 1):
            cell = sheet.cell(row_index, column, value)
            cell.data_type = "s"
            cell.number_format = "@"


__all__ = [
    "StructuredProjectionWorkbook",
    "create_table_sheet",
    "materialize_order_rows",
    "prepare_projection_workbook",
    "unique_sheet_name",
    "write_string_matrix",
]
