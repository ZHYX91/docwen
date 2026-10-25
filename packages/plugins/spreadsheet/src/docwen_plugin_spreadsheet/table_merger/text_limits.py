"""Reject merged text before openpyxl can silently truncate a cell."""

from __future__ import annotations

from docwen_plugin_spreadsheet.xlsx_text import XLSX_CELL_TEXT_LIMIT, xlsx_cell_text_length


class MergeCellTextTooLongError(ValueError):
    """The final scalar cannot be represented by an Excel cell."""

    def __init__(self, *, sheet: str, coordinate: str, length: int) -> None:
        self.location = f"{sheet}!{coordinate}"
        self.length = length
        self.limit = XLSX_CELL_TEXT_LIMIT
        super().__init__(
            f"Merged cell text at {self.location} has {length} UTF-16 code units; "
            f"Excel supports at most {self.limit} UTF-16 code units in one cell."
        )


def check_cell_text(value: str, *, sheet: str, coordinate: str) -> None:
    """Use the same UTF-16 boundary as the delimited-input XLSX builder."""
    length = xlsx_cell_text_length(value)
    if length > XLSX_CELL_TEXT_LIMIT:
        raise MergeCellTextTooLongError(sheet=sheet, coordinate=coordinate, length=length)
