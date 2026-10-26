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


def assign_bounded_xlsx_value(cell: Any, value: Any) -> None:
    """Check the final value, before any assignment to openpyxl."""
    if isinstance(value, str):
        units = len(value.encode("utf-16-le")) // 2
        if units > XLSX_CELL_TEXT_LIMIT:
            raise XlsxCellTextTooLongError(sheet=cell.parent.title, row=cell.row, column=cell.column, length=units)
    cell.value = value
