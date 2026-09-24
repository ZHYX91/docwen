from __future__ import annotations

import os
from pathlib import Path

import pytest
from tools import recycle

pytestmark = [pytest.mark.unit, pytest.mark.skipif(os.name != "nt", reason="Windows Recycle Bin")]


def test_permanent_delete_guard_raises_the_com_abort_error() -> None:
    import winerror
    from win32com.server.exception import COMException

    guard = recycle._recycle_guard()
    with pytest.raises(COMException, match="permanent_deletion_refused") as error:
        guard.PreDeleteItem(0, None)
    assert error.value.hresult == winerror.E_ABORT
    assert guard.refused


@pytest.mark.parametrize("directory", [False, True])
def test_real_recycle_returns_recoverable_payload(tmp_path: Path, directory: bool) -> None:
    target = tmp_path / "recycle-proof"
    if directory:
        target.mkdir()
        (target / "sample.txt").write_text("recoverable")
    else:
        target.write_text("recoverable")
    entries = recycle.recycle_path(target)
    assert not target.exists()
    assert entries
    payload = Path(entries[0]["payload"])
    assert (payload / "sample.txt" if directory else payload).read_text() == "recoverable"


def test_recycle_failure_keeps_target(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    target = tmp_path / "keep.txt"
    target.write_text("keep")

    def fail(path: Path) -> None:
        raise OSError("recycle unavailable")

    monkeypatch.setattr(recycle, "_perform_recycle", fail)
    with pytest.raises(OSError, match="recycle unavailable"):
        recycle.recycle_path(target)
    assert target.read_text() == "keep"


def test_shell_permanent_delete_is_vetoed_before_mutation(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    target = tmp_path / "must-survive.txt"
    target.write_text("keep")
    # Deliberately request permanent deletion of this disposable fixture. The real
    # COM callback must abort it; a Python HRESULT return alone is not sufficient.
    monkeypatch.setattr(recycle, "_RECYCLE_FLAGS", 0x100000 | 0x400 | 0x10 | 0x4)
    with pytest.raises(OSError, match="permanent_deletion_refused"):
        recycle.recycle_path(target)
    assert target.read_text() == "keep"


def test_recycle_waits_for_new_metadata(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    target = tmp_path / "source"
    target.write_text("keep")
    old = {"metadata": "old", "payload": "old-payload"}
    new = {"metadata": "new", "payload": "new-payload"}
    observations = iter([[old], [old], [old, new]])
    monkeypatch.setattr(recycle, "_recovery_entries", lambda path: next(observations))
    monkeypatch.setattr(recycle, "_perform_recycle", lambda path: path.rename(tmp_path / "moved"))
    monkeypatch.setattr(recycle.time, "sleep", lambda seconds: None)
    assert recycle.recycle_path(target) == [new]
