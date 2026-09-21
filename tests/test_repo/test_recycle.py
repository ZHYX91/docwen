from __future__ import annotations

import os
from pathlib import Path

import pytest
from tools import recycle

pytestmark = [pytest.mark.unit, pytest.mark.skipif(os.name != "nt", reason="Windows Recycle Bin")]


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
