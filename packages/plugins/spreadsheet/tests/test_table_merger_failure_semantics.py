"""Failure semantics for spreadsheet table merging."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from docwen_plugin_spreadsheet.table_merger.converter import TableMergerConverter

pytestmark = pytest.mark.unit


def test_unmerge_failure_is_not_silently_accepted() -> None:
    merged_range = SimpleNamespace(min_row=1, min_col=1, max_row=2, max_col=2)

    class BrokenWorksheet:
        merged_cells = SimpleNamespace(ranges=[merged_range])

        @staticmethod
        def cell(*, row: int, column: int) -> SimpleNamespace:
            del row, column
            return SimpleNamespace(value="merged")

        @staticmethod
        def unmerge_cells(_range: str) -> None:
            raise RuntimeError("corrupt merge metadata")

    with pytest.raises(RuntimeError, match="corrupt merge metadata"):
        TableMergerConverter._unmerge_all_cells(BrokenWorksheet())
