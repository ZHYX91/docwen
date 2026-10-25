"""Shared native Excel cell text boundary for spreadsheet writers."""

from __future__ import annotations

XLSX_CELL_TEXT_LIMIT = 32767


def xlsx_cell_text_length(value: str) -> int:
    """Count UTF-16 code units, including both units of supplementary characters."""
    return len(value.encode("utf-16-le")) // 2
