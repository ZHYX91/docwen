"""Prove Windows-only bounded retry and exact failure preservation."""

from __future__ import annotations

import errno
from pathlib import Path

import pytest

from docwen_core import file_io

pytestmark = pytest.mark.unit


class _WindowsPermissionError(PermissionError):
    winerror: int


def _winerror(code: int) -> OSError:
    exc = _WindowsPermissionError(errno.EACCES, "temporary sharing or access denial")
    exc.winerror = code
    return exc


@pytest.mark.parametrize("code", [5, 32])
def test_retry_then_commit(monkeypatch, tmp_path: Path, code: int) -> None:
    source = tmp_path / "prepared"
    target = tmp_path / "destination"
    source.write_bytes(b"new")
    target.write_bytes(b"old")
    original = file_io.os.replace
    calls = []
    checks = []
    monkeypatch.setattr(file_io, "_is_windows", lambda: True)
    monkeypatch.setattr(file_io.time, "sleep", lambda duration: calls.append(("sleep", duration)))

    def flaky(src, dst):
        calls.append(("replace", str(src)))
        if sum(item[0] == "replace" for item in calls) < 3:
            raise _winerror(code)
        original(src, dst)

    monkeypatch.setattr(file_io.os, "replace", flaky)
    file_io.replace_authorized_file(source, target, before_attempt=lambda: checks.append("checked"))
    assert target.read_bytes() == b"new"
    assert not source.exists()
    assert len(checks) == 3
    assert sum(item[0] == "sleep" for item in calls) == 2


def test_exhaustion_preserves_destination_and_original_error(monkeypatch, tmp_path: Path) -> None:
    source, target = tmp_path / "source", tmp_path / "target"
    source.write_bytes(b"new")
    target.write_bytes(b"old")
    tries = 0
    monkeypatch.setattr(file_io, "_is_windows", lambda: True)
    monkeypatch.setattr(file_io.time, "sleep", lambda delay: None)

    def denied(*args):
        nonlocal tries
        tries += 1
        raise _winerror(5)

    monkeypatch.setattr(file_io.os, "replace", denied)
    with pytest.raises(PermissionError):
        file_io.replace_authorized_file(source, target)
    assert tries == 5
    assert source.read_bytes() == b"new" and target.read_bytes() == b"old"


def test_nonwindows_never_retries(monkeypatch, tmp_path: Path) -> None:
    tries = 0
    monkeypatch.setattr(file_io, "_is_windows", lambda: False)

    def denied(*args):
        nonlocal tries
        tries += 1
        raise _winerror(32)

    monkeypatch.setattr(file_io.os, "replace", denied)
    with pytest.raises(PermissionError):
        file_io.replace_authorized_file(tmp_path / "a", tmp_path / "b")
    assert tries == 1


def test_cancellation_and_identity_recheck_before_every_attempt(monkeypatch, tmp_path: Path) -> None:
    from docwen_core.errors import CancellationRequested

    tries, checks = 0, 0
    monkeypatch.setattr(file_io, "_is_windows", lambda: True)
    monkeypatch.setattr(file_io.time, "sleep", lambda delay: None)

    def check():
        nonlocal checks
        checks += 1
        if checks == 2:
            raise CancellationRequested("cancelled")

    def denied(*args):
        nonlocal tries
        tries += 1
        raise _winerror(32)

    monkeypatch.setattr(file_io.os, "replace", denied)
    with pytest.raises(CancellationRequested):
        file_io.replace_authorized_file(tmp_path / "a", tmp_path / "b", before_attempt=check)
    assert (tries, checks) == (1, 2)
