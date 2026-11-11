"""Sparse delimited output has bounded geometry and exclusive file ownership."""

from __future__ import annotations

import csv
import errno
from dataclasses import dataclass
from pathlib import Path
from typing import Any

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


@dataclass(frozen=True)
class _Cell:
    value: Any


def test_rectangular_budget_is_cumulative_and_accepts_its_exact_boundary() -> None:
    views = [
        WorksheetValueView("First", 2, 3, {(2, 3): _Cell(value="last")}),
        WorksheetValueView("Second", 2, 2, {(2, 2): _Cell(value="last")}),
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
        (1, 1): _Cell(value=0),
        (1, 2): _Cell(value=False),
        (1, 3): _Cell(value="=1+2"),
        (101, 100): _Cell(value="last\nvalue"),
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

    cells = Values({(rows, columns): _Cell(value="last")})
    view = plan_delimited_projection([WorksheetValueView("Sparse", rows, columns, cells)], cancel_check=token.check)[0]
    with pytest.raises(CancellationRequested):
        write_delimited_projection(view, tmp_path / "output.csv", cancel_check=token.check)
    assert 17 <= cells.visits <= 1000
    assert len(cells) == 1
    assert list(tmp_path.iterdir()) == []


def test_existing_output_is_refused_before_reading_values(tmp_path: Path) -> None:
    output = tmp_path / "existing.csv"
    output.write_bytes(b"owned by an earlier conversion")
    view = WorksheetValueView("Data", 1, 1, {(1, 1): _Cell(value="new")})
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

    view = WorksheetValueView("Data", 1, 1, Values({(1, 1): _Cell(value="new")}))
    with pytest.raises(FileExistsError):
        write_delimited_projection(view, output, cancel_check=lambda: None)
    assert output.read_bytes() == b"other owner won the destination"
    assert list(tmp_path.iterdir()) == [output]


@pytest.mark.parametrize("replacement", ["new_inode", "changed_contents", "symlink"])
def test_write_receipt_does_not_adopt_or_delete_changed_output(tmp_path: Path, replacement: str) -> None:
    output = tmp_path / "output.csv"
    view = WorksheetValueView("Data", 1, 1, {(1, 1): _Cell(value="original")})
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
    view = WorksheetValueView("Data", 1, 1, {(1, 1): _Cell(value="original")})

    def unsupported(_source, _destination):
        raise OSError(errno.ENOTSUP, "Hard links unsupported")

    monkeypatch.setattr(spreadsheet_projection.os, "link", unsupported)
    with pytest.raises(OSError, match="Hard links unsupported"):
        write_delimited_projection(view, output, cancel_check=lambda: None)
    assert list(tmp_path.iterdir()) == []


@pytest.mark.parametrize("failure", ["initial_fstat", "fdopen"])
def test_failure_before_stream_ownership_closes_the_real_raw_descriptor_and_cleans_its_file(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, failure: str
) -> None:
    view = WorksheetValueView("Data", 1, 1, {(1, 1): _Cell(value="original")})
    error = OSError("injected before stream ownership")
    descriptors: list[int] = []
    injected = False
    mkstemp, fstat, fdopen = (
        spreadsheet_projection.tempfile.mkstemp,
        spreadsheet_projection.os.fstat,
        spreadsheet_projection.os.fdopen,
    )

    def create(*args, **kwargs):
        descriptor, path = mkstemp(*args, **kwargs)
        descriptors.append(descriptor)
        return descriptor, path

    def fail_stat(descriptor):
        nonlocal injected
        if not injected and descriptor in descriptors:
            injected = True
            raise error
        return fstat(descriptor)

    def fail_open(descriptor, *args, **kwargs):
        nonlocal injected
        if descriptor in descriptors:
            injected = True
            raise error
        return fdopen(descriptor, *args, **kwargs)

    monkeypatch.setattr(spreadsheet_projection.tempfile, "mkstemp", create)
    if failure == "initial_fstat":
        monkeypatch.setattr(spreadsheet_projection.os, "fstat", fail_stat)
    else:
        monkeypatch.setattr(spreadsheet_projection.os, "fdopen", fail_open)
    with pytest.raises(OSError) as raised:
        write_delimited_projection(view, tmp_path / "output.csv", cancel_check=lambda: None)
    assert injected and raised.value is error and len(descriptors) == 1
    with pytest.raises(OSError) as closed:
        fstat(descriptors[0])
    assert closed.value.errno == errno.EBADF
    assert list(tmp_path.iterdir()) == []


@pytest.mark.parametrize("failure", ["temporary_unlink", "receipt_stat", "receipt_construction"])
def test_failure_before_receipt_returns_removes_the_installed_output(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, failure: str
) -> None:
    output = tmp_path / "output.csv"
    view = WorksheetValueView("Data", 1, 1, {(1, 1): _Cell(value="original")})
    error = PermissionError("injected after installation")
    injected = False
    unlink, lstat = Path.unlink, Path.lstat

    def fail_once():
        nonlocal injected
        assert output.read_bytes() == b"\xef\xbb\xbforiginal\r\n"
        injected = True
        raise error

    def fail_unlink(path, *args, **kwargs):
        if not injected and path.name.startswith("delimited-") and output.exists():
            fail_once()
        return unlink(path, *args, **kwargs)

    def fail_stat(path, *args, **kwargs):
        if not injected and path == output:
            fail_once()
        return lstat(path, *args, **kwargs)

    def fail_receipt(*_args, **_kwargs):
        fail_once()

    if failure == "temporary_unlink":
        monkeypatch.setattr(Path, "unlink", fail_unlink)
    elif failure == "receipt_stat":
        monkeypatch.setattr(Path, "lstat", fail_stat)
    else:
        monkeypatch.setattr(spreadsheet_projection, "DelimitedWrite", fail_receipt)
    with pytest.raises(PermissionError) as raised:
        write_delimited_projection(view, output, cancel_check=lambda: None)
    assert injected and raised.value is error
    assert list(tmp_path.iterdir()) == []


@pytest.mark.parametrize("cancel", [False, True])
def test_replaced_temporary_file_is_not_installed_and_cleanup_preserves_the_original_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, cancel: bool
) -> None:
    output = tmp_path / "output.csv"
    error = CancellationRequested("cancel after closing values")
    replacements: list[Path] = []
    handles = []
    fdopen = spreadsheet_projection.os.fdopen

    def capture_handle(*args, **kwargs):
        handle = fdopen(*args, **kwargs)
        handles.append(handle)
        return handle

    def cancel_check():
        if handles and handles[0].closed and not replacements:
            # The handle is really closed, so this replacement is possible on
            # Windows as well as Unix; the writer has not installed it yet.
            [replacement] = tmp_path.glob("delimited-*.csv")
            replacements.append(replacement)
            foreign = tmp_path / "foreign.csv"
            foreign.write_bytes(b"a different temporary owner")
            foreign.replace(replacement)
            if cancel:
                raise error

    monkeypatch.setattr(spreadsheet_projection.os, "fdopen", capture_handle)
    view = WorksheetValueView("Data", 1, 1, {(1, 1): _Cell(value="original")})
    if cancel:
        with pytest.raises(CancellationRequested) as raised:
            write_delimited_projection(view, output, cancel_check=cancel_check)
        assert raised.value is error
        assert any("ownership is unconfirmed" in note for note in error.__notes__)
    else:
        with pytest.raises(RuntimeError, match="temporary file changed before installation"):
            write_delimited_projection(view, output, cancel_check=cancel_check)
    assert len(replacements) == 1 and replacements[0].read_bytes() == b"a different temporary owner"
    assert list(tmp_path.iterdir()) == replacements
