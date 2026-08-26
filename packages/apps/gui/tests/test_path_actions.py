"""Permanent cross-platform open/reveal selector contracts."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from docwen_gui import path_actions
from docwen_gui.path_actions import PathActionResult

pytestmark = pytest.mark.unit


class _CompletedProcess:
    def __init__(self, returncode: int = 0, *, stdout: str = "", stderr: str = "") -> None:
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr


def test_windows_reveal_selects_exact_file(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    sample = tmp_path / "report.docx"
    sample.write_text("done", encoding="utf-8")
    commands: list[list[str]] = []
    monkeypatch.setattr(path_actions.sys, "platform", "win32")
    monkeypatch.setattr(
        path_actions.subprocess,
        "run",
        lambda command, **kwargs: commands.append(list(command)) or _CompletedProcess(),
    )

    result = path_actions.reveal_path(sample)

    assert result == PathActionResult(success=True, precise=True)
    assert commands == [["explorer", "/select,", str(sample)]]


@pytest.mark.skipif(os.name != "nt", reason="Win32 extended-path boundary")
def test_windows_reveal_probes_long_file_without_leaking_extended_prefix(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    remaining = 260 - len(os.path.abspath(tmp_path)) - 1
    if remaining < 1 or remaining > 255:
        pytest.skip("pytest temp root cannot express an exact 260-character file")
    sample = tmp_path / ("r" * remaining)
    path_actions.filesystem_path(sample).write_bytes(b"done")
    commands: list[list[str]] = []
    monkeypatch.setattr(path_actions.sys, "platform", "win32")
    monkeypatch.setattr(
        path_actions.subprocess,
        "run",
        lambda command, **kwargs: commands.append(list(command)) or _CompletedProcess(),
    )

    result = path_actions.reveal_path(sample)

    assert result == PathActionResult(success=True, precise=True)
    assert commands == [["explorer", "/select,", str(sample)]]
    assert not commands[0][-1].startswith("\\\\?\\")


def test_macos_reveal_uses_open_dash_r(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    sample = tmp_path / "report.docx"
    sample.write_text("done", encoding="utf-8")
    commands: list[list[str]] = []
    monkeypatch.setattr(path_actions.sys, "platform", "darwin")
    monkeypatch.setattr(
        path_actions.subprocess,
        "run",
        lambda command, **kwargs: commands.append(list(command)) or _CompletedProcess(),
    )

    result = path_actions.reveal_path(sample)

    assert result == PathActionResult(success=True, precise=True)
    assert commands == [["open", "-R", str(sample)]]


def test_linux_reveal_tries_standard_selectors_then_parent_fallback(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    sample = tmp_path / "report.docx"
    sample.write_text("done", encoding="utf-8")
    commands: list[list[str]] = []
    opened: list[Path] = []
    monkeypatch.setattr(path_actions.sys, "platform", "linux")
    monkeypatch.setattr(
        path_actions.subprocess,
        "run",
        lambda command, **kwargs: commands.append(list(command)) or _CompletedProcess(1, stderr="unavailable"),
    )
    monkeypatch.setattr(
        path_actions,
        "_open_with_desktop_services",
        lambda target: opened.append(target) or PathActionResult(success=True),
    )

    result = path_actions.reveal_path(sample)

    assert result == PathActionResult(success=True, fallback_used=True)
    assert commands == path_actions.linux_reveal_commands(sample)
    assert opened == [sample.parent]


def test_open_path_opens_file_then_falls_back_to_parent(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    sample = tmp_path / "report.docx"
    sample.write_text("done", encoding="utf-8")
    opened: list[Path] = []

    def _open(target: Path) -> PathActionResult:
        opened.append(target)
        return PathActionResult(success=len(opened) == 2, error_code=None if len(opened) == 2 else "open_failed")

    monkeypatch.setattr(path_actions, "_open_with_desktop_services", _open)

    result = path_actions.open_path(sample)

    assert result == PathActionResult(success=True, fallback_used=True)
    assert opened == [sample, sample.parent]


def test_failed_precise_command_falls_back_instead_of_claiming_success(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    sample = tmp_path / "report.docx"
    sample.write_text("done", encoding="utf-8")
    monkeypatch.setattr(path_actions.sys, "platform", "win32")
    monkeypatch.setattr(
        path_actions,
        "_run_command",
        lambda command: PathActionResult(success=False, error="blocked", error_code="command_failed"),
    )
    monkeypatch.setattr(
        path_actions,
        "_open_with_desktop_services",
        lambda target: PathActionResult(success=True),
    )

    result = path_actions.reveal_path(sample)

    assert result == PathActionResult(success=True, fallback_used=True)


def test_reveal_missing_path_fails_closed_without_opening_parent(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    missing = tmp_path / "missing.docx"
    opened: list[Path] = []
    monkeypatch.setattr(
        path_actions,
        "_open_with_desktop_services",
        lambda target: opened.append(target) or PathActionResult(success=True),
    )

    result = path_actions.reveal_path(missing)

    assert result == PathActionResult(success=False, error=str(missing), error_code="missing_path")
    assert opened == []
