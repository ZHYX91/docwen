"""Bounded spreadsheet value projections, independent of workbook libraries."""

from __future__ import annotations

import csv
import os
import tempfile
from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any, Protocol

MAX_SPREADSHEET_PROJECTION_CELLS = 1_000_000
_CANCEL_INTERVAL = 1000


class SpreadsheetProjectionLimitError(ValueError):
    """A request would allocate more native or projected cells than allowed."""

    diagnostic_code = "SPREADSHEET-PROJECTION-LIMIT"


class NoOrdinaryWorksheetsError(ValueError):
    """A chart-only workbook has no worksheet output to publish."""

    diagnostic_code = "SPREADSHEET-NO-WORKSHEETS"


@dataclass
class ProjectionBudget:
    limit: int = MAX_SPREADSHEET_PROJECTION_CELLS
    used: int = 0

    def reserve(self, count: int, *, location: str) -> None:
        if count < 0 or self.used + count > self.limit:
            raise SpreadsheetProjectionLimitError(f"Spreadsheet projection exceeds {self.limit} cells at {location}.")
        self.used += count


def ordinary_worksheets[T](worksheets: Iterable[T]) -> tuple[T, ...]:
    """Require the library's typed worksheet collection, retaining its order.

    Callers pass ``workbook.worksheets``, never the workbook's mixed sheet list.
    Core does not import a particular library's worksheet or chartsheet classes.
    """
    result = tuple(worksheets)
    if not result:
        raise NoOrdinaryWorksheetsError("The workbook contains no ordinary worksheets to export.")
    return result


class CellValue(Protocol):
    @property
    def value(self) -> Any: ...


@dataclass(frozen=True)
class WorksheetValueView:
    """A sparse, read-only value map and its intended coordinate rectangle."""

    title: str
    row_count: int
    column_count: int
    cells: Mapping[tuple[int, int], CellValue]


def plan_delimited_projection(
    views: Iterable[WorksheetValueView],
    *,
    cancel_check: Callable[[], None],
    limit: int = MAX_SPREADSHEET_PROJECTION_CELLS,
) -> tuple[WorksheetValueView, ...]:
    """Reserve all output rectangles before any output file is created."""
    # Workbook libraries report a default A1 extent even when no cell exists.
    # Keep the worksheet's output identity without inventing a blank CSV row.
    selected = ordinary_worksheets(view if view.cells else replace(view, row_count=0, column_count=0) for view in views)
    budget = ProjectionBudget(limit=limit)
    for view in selected:
        cancel_check()
        if view.row_count < 0 or view.column_count < 0:
            raise ValueError("Worksheet projection dimensions must be nonnegative.")
        budget.reserve(view.row_count * view.column_count, location=view.title)
    return selected


def stored_value_rows(view: WorksheetValueView, *, cancel_check: Callable[[], None]) -> Iterator[list[Any]]:
    """Read bounded row buffers without allocating absent workbook cells."""
    visited = 0
    for row in range(1, view.row_count + 1):
        values: list[Any] = []
        for column in range(1, view.column_count + 1):
            if visited % _CANCEL_INTERVAL == 0:
                cancel_check()
            visited += 1
            cell = view.cells.get((row, column))
            values.append(None if cell is None else cell.value)
        yield values
    cancel_check()


def write_delimited_projection(
    view: WorksheetValueView,
    output_path: str | Path,
    *,
    cancel_check: Callable[[], None],
    delimiter: str = ",",
) -> int:
    """Commit one complete UTF-8/BOM projection, cleaning an interrupted write."""
    cancel_check()
    output = Path(output_path)
    if os.path.lexists(output):
        raise FileExistsError(f"Spreadsheet projection output already exists: {output}")
    descriptor, name = tempfile.mkstemp(prefix="delimited-", suffix=output.suffix, dir=output.parent)
    temporary = Path(name)
    row_count = 0
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=delimiter)
            for row in stored_value_rows(view, cancel_check=cancel_check):
                # csv.writer represents None as an empty field while preserving
                # booleans, numbers and formula-looking strings from this view.
                writer.writerow(row)
                row_count += 1
        cancel_check()
        # A same-directory hard link installs the fully closed file atomically
        # and refuses a target that appeared after the initial collision check.
        os.link(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)
    return row_count
