"""Bounded spreadsheet value projections, independent of workbook libraries."""

from __future__ import annotations

import csv
import os
import stat
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


def _file_identity(value: os.stat_result) -> tuple[int, int, int, int, int, int]:
    return value.st_dev, value.st_ino, value.st_mode, value.st_size, value.st_mtime_ns, value.st_ctime_ns


@dataclass(frozen=True)
class DelimitedWrite:
    """Receipt for exactly the regular file created by one projection write."""

    path: Path
    row_count: int
    identity: tuple[int, int, int, int, int, int]

    def validate(self) -> None:
        current = self.path.lstat()
        if not stat.S_ISREG(current.st_mode) or _file_identity(current) != self.identity:
            raise RuntimeError("Spreadsheet staging output changed after its creation; ownership is unconfirmed.")

    def cleanup(self) -> None:
        """Remove this write only; preserve any replacement during rollback."""
        try:
            current = self.path.lstat()
        except FileNotFoundError:
            return
        if not stat.S_ISREG(current.st_mode) or _file_identity(current) != self.identity:
            return
        self.path.unlink()


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


def _remove_created_file(path: Path, created: os.stat_result, *, complete: bool) -> None:
    """Keep any path whose identity no longer belongs to this writer."""
    try:
        current = path.lstat()
    except FileNotFoundError:
        return
    same_inode = (current.st_dev, current.st_ino) == (created.st_dev, created.st_ino)
    # Linking/unlinking changes ctime even when the contents remain our own.
    same_contents = _file_identity(current)[:-1] == _file_identity(created)[:-1]
    if not stat.S_ISREG(current.st_mode) or not same_inode or (complete and not same_contents):
        raise RuntimeError(f"Spreadsheet file changed before cleanup; ownership is unconfirmed: {path}")
    path.unlink()


def write_delimited_projection(
    view: WorksheetValueView,
    output_path: str | Path,
    *,
    cancel_check: Callable[[], None],
    delimiter: str = ",",
) -> DelimitedWrite:
    """Commit one complete UTF-8/BOM projection, cleaning an interrupted write."""
    cancel_check()
    output = Path(output_path)
    if os.path.lexists(output):
        raise FileExistsError(f"Spreadsheet projection output already exists: {output}")
    descriptor, name = tempfile.mkstemp(prefix="delimited-", suffix=output.suffix, dir=output.parent)
    temporary = Path(name)
    created = os.fstat(descriptor)
    row_count = 0
    installed = False
    written: os.stat_result | None = None
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=delimiter)
            for row in stored_value_rows(view, cancel_check=cancel_check):
                # csv.writer represents None as an empty field while preserving
                # booleans, numbers and formula-looking strings from this view.
                writer.writerow(row)
                row_count += 1
            handle.flush()
            written = os.fstat(handle.fileno())
        cancel_check()
        # A same-directory hard link installs the fully closed file atomically
        # and refuses a target that appeared after the initial collision check.
        os.link(temporary, output)
        installed = True
        _remove_created_file(temporary, written, complete=True)
        observed = output.lstat()
        if _file_identity(observed)[:-1] != _file_identity(written)[:-1]:
            raise RuntimeError("Spreadsheet staging output changed during installation; ownership is unconfirmed.")
        return DelimitedWrite(output, row_count, _file_identity(observed))
    except BaseException as error:
        # Ownership transfers to the caller only when a receipt is returned.
        # Until then, this writer must also unwind a successful installation.
        cleanup = [(temporary, created, False)]
        if installed and written is not None:
            cleanup.insert(0, (output, written, True))
        for path, identity, complete in cleanup:
            try:
                _remove_created_file(path, identity, complete=complete)
            except BaseException as cleanup_error:
                error.add_note(
                    f"Spreadsheet cleanup failed for {path}: {type(cleanup_error).__name__}: {cleanup_error}"
                )
        raise
