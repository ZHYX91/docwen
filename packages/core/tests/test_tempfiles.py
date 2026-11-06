"""Shared temporary directory creation and cleanup contracts."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from docwen_core.tempfiles import TemporaryDirectory, TemporaryDirectoryError, make_temp_dir

pytestmark = pytest.mark.unit


def test_permission_denial_is_not_retried(monkeypatch, tmp_path: Path) -> None:
    calls = 0

    def denied(path, mode=0o777):
        nonlocal calls
        calls += 1
        raise PermissionError(13, "access denied", str(path))

    monkeypatch.setattr(os, "mkdir", denied)
    with pytest.raises(TemporaryDirectoryError):
        make_temp_dir(prefix="dw-", dir=tmp_path)
    assert calls == 1


def test_name_collision_is_bounded_and_checks_cancellation(monkeypatch, tmp_path: Path) -> None:
    from docwen_core.tempfiles import make_temp_dir

    calls = 0
    cancelled = False

    def conflict(path, mode=0o777):
        nonlocal calls, cancelled
        calls += 1
        if calls == 3:
            cancelled = True
        raise FileExistsError(17, "occupied", str(path))

    def check():
        if cancelled:
            from docwen_core.errors import CancellationRequested
            raise CancellationRequested("cancelled")

    monkeypatch.setattr(os, "mkdir", conflict)
    from docwen_core.errors import CancellationRequested
    with pytest.raises(CancellationRequested, match="cancelled"):
        make_temp_dir(prefix="dw-", dir=tmp_path, cancellation_check=check)
    assert calls == 3


def test_cleanup_failure_does_not_mask_primary_error(monkeypatch, tmp_path: Path) -> None:
    owner = TemporaryDirectory(prefix="dw-", dir=tmp_path)
    original = owner.name

    def denied(_path):
        raise PermissionError(13, "cleanup denied")

    monkeypatch.setattr("docwen_core.tempfiles.shutil.rmtree", denied)
    with pytest.raises(ValueError, match="conversion failed") as caught:
        with owner:
            raise ValueError("conversion failed")
    assert Path(original).is_dir()
    assert any("cleanup failed" in note.lower() for note in getattr(caught.value, "__notes__", []))


def test_cleanup_failure_is_reported_without_primary_error(monkeypatch, tmp_path: Path) -> None:
    owner = TemporaryDirectory(prefix="dw-", dir=tmp_path)
    monkeypatch.setattr("docwen_core.tempfiles.shutil.rmtree", lambda _path: (_ for _ in ()).throw(PermissionError()))
    with pytest.raises(PermissionError):
        owner.cleanup()
