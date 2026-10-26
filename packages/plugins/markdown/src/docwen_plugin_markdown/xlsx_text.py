"""Reject authored text before openpyxl silently truncates an XLSX cell."""

from __future__ import annotations

from typing import Any

XLSX_CELL_TEXT_LIMIT = 32767


class XlsxCellTextTooLongError(ValueError):
    """The final composed string exceeds Excel's UTF-16 capacity."""

    def __init__(self, *, sheet: str, row: int, column: int, length: int) -> None:
        self.sheet = sheet
        self.row = row
        self.column = column
        self.length = length
        super().__init__(
            f"Cell in sheet {sheet!r}, row {row}, column {column} contains "
            f"{length} UTF-16 code units; Excel allows at most {XLSX_CELL_TEXT_LIMIT}."
        )


def check_xlsx_cell_text(cell: Any, value: Any) -> None:
    """Check a text scalar without assigning or changing its cell type."""
    if isinstance(value, str):
        units = len(value.encode("utf-16-le")) // 2
        if units > XLSX_CELL_TEXT_LIMIT:
            raise XlsxCellTextTooLongError(sheet=cell.parent.title, row=cell.row, column=cell.column, length=units)


def check_workbook_text_capacity(workbook: Any) -> None:
    """Include retained template cells without expanding sparse worksheet ranges."""
    for sheet in workbook.worksheets:
        for cell in sheet._cells.values():
            if cell.data_type != "f":
                check_xlsx_cell_text(cell, cell.value)


def existing_xlsx_cells(worksheet: Any) -> list[Any]:
    """Snapshot populated cells in row order without expanding sparse ranges."""
    return [cell for _position, cell in sorted(worksheet._cells.items())]


def assign_bounded_xlsx_value(cell: Any, value: Any) -> None:
    """Check the final value, before any assignment to openpyxl."""
    check_xlsx_cell_text(cell, value)
    cell.value = value
