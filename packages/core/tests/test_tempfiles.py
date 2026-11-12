"""Shared temporary directory creation and cleanup contracts."""

from __future__ import annotations

import os
import stat
from pathlib import Path

import pytest

from docwen_core.errors import CancellationRequested
from docwen_core.tempfiles import TemporaryDirectory, TemporaryDirectoryError, make_publication_dir, make_temp_dir

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("allocate", [make_temp_dir, make_publication_dir])
def test_permission_denial_is_not_retried(monkeypatch, tmp_path: Path, allocate) -> None:
    calls = 0

    def denied(path, mode=0o777):
        nonlocal calls
        calls += 1
        raise PermissionError(13, "access denied", str(path))

    monkeypatch.setattr(os, "mkdir", denied)
    with pytest.raises(TemporaryDirectoryError):
        allocate(prefix="dw-", dir=tmp_path)
    assert calls == 1


@pytest.mark.parametrize("allocate", [make_temp_dir, make_publication_dir])
def test_name_collision_is_bounded_and_checks_cancellation(monkeypatch, tmp_path: Path, allocate) -> None:
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
            raise CancellationRequested("cancelled")

    monkeypatch.setattr(os, "mkdir", conflict)
    with pytest.raises(CancellationRequested, match="cancelled"):
        allocate(prefix="dw-", dir=tmp_path, cancellation_check=check)
    assert calls == 3


@pytest.mark.parametrize("allocate", [make_temp_dir, make_publication_dir])
def test_allocation_attempts_are_bounded_without_changing_an_existing_entry(
    monkeypatch, tmp_path: Path, allocate
) -> None:
    monkeypatch.setattr("docwen_core.tempfiles.secrets.token_hex", lambda _count: "occupied")
    occupied = tmp_path / "dw-occupied"
    occupied.mkdir()
    marker = occupied / "foreign"
    marker.write_bytes(b"preserve existing bytes")
    with pytest.raises(TemporaryDirectoryError, match="unique"):
        allocate(prefix="dw-", dir=tmp_path, attempts=2)
    assert marker.read_bytes() == b"preserve existing bytes"
    assert list(tmp_path.iterdir()) == [occupied]
    with pytest.raises(ValueError, match="positive"):
        allocate(dir=tmp_path, attempts=0)


@pytest.mark.skipif(os.name == "nt", reason="POSIX mode and umask semantics")
def test_publication_directory_uses_normal_umask_and_private_work_remains_private(tmp_path: Path) -> None:
    normal = tmp_path / "normal"
    normal.mkdir()
    published = Path(make_publication_dir(prefix="output-", dir=tmp_path))
    private = Path(make_temp_dir(prefix="work-", dir=tmp_path))
    assert stat.S_IMODE(published.stat().st_mode) == stat.S_IMODE(normal.stat().st_mode)
    assert stat.S_IMODE(private.stat().st_mode) == 0o700


def test_cleanup_failure_does_not_mask_primary_error(monkeypatch, tmp_path: Path) -> None:
    owner = TemporaryDirectory(prefix="dw-", dir=tmp_path)
    original = owner.name

    def denied(_path):
        raise PermissionError(13, "cleanup denied")

    primary = ValueError("conversion failed")
    with monkeypatch.context() as patcher:
        patcher.setattr("docwen_core.tempfiles.shutil.rmtree", denied)
        with pytest.raises(ValueError, match="conversion failed") as caught, owner:
            raise primary
    assert caught.value is primary
    assert Path(original).is_dir()
    assert any("cleanup failed" in note.lower() for note in getattr(caught.value, "__notes__", []))
    owner.cleanup()
    assert not Path(original).exists()


def test_cleanup_failure_is_reported_without_primary_error(monkeypatch, tmp_path: Path) -> None:
    owner = TemporaryDirectory(prefix="dw-", dir=tmp_path)
    failure = PermissionError(13, "cleanup denied")

    def denied(_path):
        raise failure

    with monkeypatch.context() as patcher:
        patcher.setattr("docwen_core.tempfiles.shutil.rmtree", denied)
        with pytest.raises(PermissionError) as caught:
            owner.cleanup()
    assert caught.value is failure
    assert Path(owner.name).is_dir()
    owner.cleanup()
    assert not Path(owner.name).exists()
    owner.cleanup()


def test_ignored_cleanup_failure_can_still_be_retried(monkeypatch, tmp_path: Path) -> None:
    owner = TemporaryDirectory(prefix="dw-", dir=tmp_path, ignore_cleanup_errors=True)

    def denied(_path):
        raise PermissionError(13, "cleanup denied")

    with monkeypatch.context() as patcher:
        patcher.setattr("docwen_core.tempfiles.shutil.rmtree", denied)
        with owner:
            (Path(owner.name) / "result").write_text("keep until cleanup succeeds", encoding="utf-8")
    assert Path(owner.name).is_dir()
    owner.cleanup()
    assert not Path(owner.name).exists()


def test_cleanup_inside_a_handled_exception_is_not_suppressed(monkeypatch, tmp_path: Path) -> None:
    owner = TemporaryDirectory(prefix="dw-", dir=tmp_path)

    def denied(_path):
        raise PermissionError(13, "cleanup denied")

    with monkeypatch.context() as patcher:
        patcher.setattr("docwen_core.tempfiles.shutil.rmtree", denied)
        try:
            raise ValueError("already handled")
        except ValueError:
            with pytest.raises(PermissionError, match="cleanup denied"), owner:
                pass
    owner.cleanup()


def test_missing_child_during_cleanup_keeps_remaining_directory_retryable(monkeypatch, tmp_path: Path) -> None:
    owner = TemporaryDirectory(prefix="dw-", dir=tmp_path)
    remaining = Path(owner.name) / "result"
    remaining.write_text("not removed", encoding="utf-8")

    def missing_child(_path):
        raise FileNotFoundError(2, "child disappeared", str(Path(owner.name) / "missing"))

    with monkeypatch.context() as patcher:
        patcher.setattr("docwen_core.tempfiles.shutil.rmtree", missing_child)
        with pytest.raises(FileNotFoundError):
            owner.cleanup()
    assert remaining.read_text(encoding="utf-8") == "not removed"
    owner.cleanup()
    assert not Path(owner.name).exists()


def test_cleanup_accepts_an_already_removed_owned_directory(tmp_path: Path) -> None:
    owner = TemporaryDirectory(prefix="dw-", dir=tmp_path)
    Path(owner.name).rmdir()
    owner.cleanup()
    owner.cleanup()
