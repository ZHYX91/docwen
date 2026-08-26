from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path

import pytest
from tools import windows_short_path

pytestmark = pytest.mark.unit


def test_tmp_path_resolves_inside_the_qa_runtime_without_changing_path_semantics(tmp_path: Path) -> None:
    runtime_value = os.environ.get("DOCWEN_PYTEST_RUNTIME_ROOT", "").strip()
    if os.name != "nt" or not runtime_value:
        pytest.skip("QA short runtime drive is not active")
    runtime_view = Path(runtime_value)
    if not runtime_view.drive or runtime_view != Path(runtime_view.anchor):
        pytest.skip("QA short runtime drive is not active")

    physical_runtime = runtime_view.resolve(strict=True)
    resolved_tmp = tmp_path.resolve(strict=True)

    assert physical_runtime in resolved_tmp.parents
    assert tmp_path.drive.casefold() == physical_runtime.drive.casefold()
    assert re.fullmatch(r"t[0-9a-f]{12}[0-9]+", tmp_path.name)


def _completed(*arguments: str, returncode: int = 0) -> subprocess.CompletedProcess[str]:
    return subprocess.CompletedProcess(arguments, returncode, stdout="", stderr="")


def test_mount_short_drive_uses_an_unused_letter_and_verifies_identity(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    target = tmp_path / "physical"
    target.mkdir()
    commands: list[tuple[str, ...]] = []
    monkeypatch.setattr(windows_short_path, "_is_windows", lambda: True)
    monkeypatch.setattr(windows_short_path, "_DRIVE_LETTERS", ("Z:",))
    monkeypatch.setattr(windows_short_path, "_drive_is_available", lambda _drive: True)
    monkeypatch.setattr(windows_short_path, "_same_directory", lambda _first, second: second == target.resolve())
    monkeypatch.setattr(
        windows_short_path,
        "_run_subst",
        lambda *arguments: commands.append(arguments) or _completed(*arguments),
    )

    assert windows_short_path.mount_short_drive(target) == "Z:"
    assert commands == [("Z:", str(target.resolve()))]


def test_mount_short_drive_rolls_back_an_unverifiable_mapping(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    target = tmp_path / "physical"
    target.mkdir()
    commands: list[tuple[str, ...]] = []
    monkeypatch.setattr(windows_short_path, "_is_windows", lambda: True)
    monkeypatch.setattr(windows_short_path, "_DRIVE_LETTERS", ("Z:",))
    monkeypatch.setattr(windows_short_path, "_drive_is_available", lambda _drive: True)
    monkeypatch.setattr(windows_short_path, "_same_directory", lambda _first, _second: False)
    monkeypatch.setattr(
        windows_short_path,
        "_run_subst",
        lambda *arguments: commands.append(arguments) or _completed(*arguments),
    )

    with pytest.raises(windows_short_path.ShortPathDriveError, match="no_safe_short_drive_available"):
        windows_short_path.mount_short_drive(target)

    assert commands == [("Z:", str(target.resolve())), ("Z:", "/D")]


def test_unmount_refuses_a_drive_that_points_somewhere_else(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    target = tmp_path / "physical"
    target.mkdir()
    monkeypatch.setattr(windows_short_path, "_is_windows", lambda: True)
    monkeypatch.setattr(windows_short_path.Path, "exists", lambda _path: True)
    monkeypatch.setattr(windows_short_path, "_same_directory", lambda _first, _second: False)

    with pytest.raises(windows_short_path.ShortPathDriveError, match="short_drive_target_mismatch"):
        windows_short_path.unmount_short_drive("Z:", expected_target=target)
