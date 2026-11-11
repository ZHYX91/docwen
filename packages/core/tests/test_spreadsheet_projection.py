"""Sparse delimited output has bounded geometry and exclusive file ownership."""

from __future__ import annotations

import csv
import errno
from pathlib import Path
from types import SimpleNamespace

import pytest

from docwen_core import spreadsheet_projection
from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested
from docwen_core.spreadsheet_projection import (
    NoOrdinaryWorksheetsError,
    SpreadsheetProjectionLimitError,
    WorksheetValueView,
    plan_delimited_projection,
    write_delimited_projection,
)

pytestmark = pytest.mark.contract


def test_rectangular_budget_is_cumulative_and_accepts_its_exact_boundary() -> None:
    views = [
        WorksheetValueView("First", 2, 3, {(2, 3): SimpleNamespace(value="last")}),
        WorksheetValueView("Second", 2, 2, {(2, 2): SimpleNamespace(value="last")}),
    ]
    assert plan_delimited_projection(views, cancel_check=lambda: None, limit=10) == tuple(views)
    with pytest.raises(SpreadsheetProjectionLimitError, match="Second"):
        plan_delimited_projection(views, cancel_check=lambda: None, limit=9)
    with pytest.raises(NoOrdinaryWorksheetsError):
        plan_delimited_projection([], cancel_check=lambda: None)


def test_empty_ordinary_sheet_keeps_an_empty_artifact_without_an_invented_row(tmp_path: Path) -> None:
    projections = plan_delimited_projection([WorksheetValueView("Empty", 1, 1, {})], cancel_check=lambda: None, limit=0)
    assert len(projections) == 1 and projections[0].title == "Empty"
    output = tmp_path / "empty.csv"
    assert write_delimited_projection(projections[0], output, cancel_check=lambda: None).row_count == 0
    assert output.read_bytes() == b""


@pytest.mark.parametrize("delimiter", [",", "\t"])
def test_saved_sparse_rectangle_retains_values_without_materializing_cells(tmp_path: Path, delimiter: str) -> None:
    cells = {
        (1, 1): SimpleNamespace(value=0),
        (1, 2): SimpleNamespace(value=False),
        (1, 3): SimpleNamespace(value="=1+2"),
        (101, 100): SimpleNamespace(value="last\nvalue"),
    }
    before = dict(cells)
    view = plan_delimited_projection([WorksheetValueView("Sparse", 101, 100, cells)], cancel_check=lambda: None)[0]
    output = tmp_path / "projection.csv"
    assert write_delimited_projection(view, output, cancel_check=lambda: None, delimiter=delimiter).row_count == 101
    assert output.read_bytes().startswith(b"\xef\xbb\xbf")
    with output.open(encoding="utf-8-sig", newline="") as handle:
        rows = list(csv.reader(handle, delimiter=delimiter))
    assert len(rows) == 101 and all(len(row) == 100 for row in rows)
    assert rows[0][:4] == ["0", "False", "=1+2", ""]
    assert rows[-1][-1] == "last\nvalue"
    assert sum(value != "" for row in rows for value in row) == 4
    assert cells == before
    assert list(tmp_path.iterdir()) == [output]


@pytest.mark.parametrize(("rows", "columns"), [(3001, 1), (1, 3001)])
def test_cancellation_within_rows_or_one_wide_row_leaves_no_partial_file(
    tmp_path: Path, rows: int, columns: int
) -> None:
    token = CancellationToken()

    class Values(dict):
        visits = 0

        def get(self, key, default=None):
            self.visits += 1
            if self.visits == 17:
                token.cancel("during values")
            return super().get(key, default)

    cells = Values({(rows, columns): SimpleNamespace(value="last")})
    view = plan_delimited_projection([WorksheetValueView("Sparse", rows, columns, cells)], cancel_check=token.check)[0]
    with pytest.raises(CancellationRequested):
        write_delimited_projection(view, tmp_path / "output.csv", cancel_check=token.check)
    assert 17 <= cells.visits <= 1000
    assert len(cells) == 1
    assert list(tmp_path.iterdir()) == []


def test_existing_output_is_refused_before_reading_values(tmp_path: Path) -> None:
    output = tmp_path / "existing.csv"
    output.write_bytes(b"owned by an earlier conversion")
    view = WorksheetValueView("Data", 1, 1, {(1, 1): SimpleNamespace(value="new")})
    with pytest.raises(FileExistsError):
        write_delimited_projection(view, output, cancel_check=lambda: None)
    assert output.read_bytes() == b"owned by an earlier conversion"
    assert list(tmp_path.iterdir()) == [output]


def test_late_destination_collision_never_replaces_or_deletes_the_other_file(tmp_path: Path) -> None:
    output = tmp_path / "late.csv"

    class Values(dict):
        def get(self, key, default=None):
            output.write_bytes(b"other owner won the destination")
            return super().get(key, default)

    view = WorksheetValueView("Data", 1, 1, Values({(1, 1): SimpleNamespace(value="new")}))
    with pytest.raises(FileExistsError):
        write_delimited_projection(view, output, cancel_check=lambda: None)
    assert output.read_bytes() == b"other owner won the destination"
    assert list(tmp_path.iterdir()) == [output]


@pytest.mark.parametrize("replacement", ["new_inode", "changed_contents", "symlink"])
def test_write_receipt_does_not_adopt_or_delete_changed_output(tmp_path: Path, replacement: str) -> None:
    output = tmp_path / "output.csv"
    view = WorksheetValueView("Data", 1, 1, {(1, 1): SimpleNamespace(value="original")})
    receipt = write_delimited_projection(view, output, cancel_check=lambda: None)
    receipt.validate()
    other = tmp_path / "other.csv"
    if replacement == "changed_contents":
        output.write_bytes(b"another writer modified this inode")
    else:
        other.write_bytes(b"a different owner")
        if replacement == "new_inode":
            other.replace(output)
        else:
            output.unlink()
            try:
                output.symlink_to(other)
            except OSError as exc:
                pytest.skip(f"Host does not allow symbolic-link creation: {exc}")
    before = output.read_bytes()
    with pytest.raises(RuntimeError, match="ownership is unconfirmed"):
        receipt.validate()
    receipt.cleanup()
    assert output.read_bytes() == before
    assert output.is_symlink() is (replacement == "symlink")
    if other.exists():
        assert other.read_bytes() == b"a different owner"


def test_unsupported_hardlink_volume_fails_without_partial_output(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    output = tmp_path / "output.csv"
    view = WorksheetValueView("Data", 1, 1, {(1, 1): SimpleNamespace(value="original")})

    def unsupported(_source, _destination):
        raise OSError(errno.ENOTSUP, "Hard links unsupported")

    monkeypatch.setattr(spreadsheet_projection.os, "link", unsupported)
    with pytest.raises(OSError, match="Hard links unsupported"):
        write_delimited_projection(view, output, cancel_check=lambda: None)
    assert list(tmp_path.iterdir()) == []
