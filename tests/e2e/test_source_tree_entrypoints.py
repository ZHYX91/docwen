from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest
from tests.support.subprocess_runner import run_subprocess

pytestmark = pytest.mark.e2e

REPO_ROOT = Path(__file__).resolve().parents[2]


def _script_command(name: str) -> list[str]:
    script_name = f"{name}.exe" if os.name == "nt" else name
    script_path = Path(sys.executable).parent / script_name
    if script_path.exists():
        return [str(script_path)]
    pytest.fail(
        f"{name!r} console script was not installed next to the test interpreter: {script_path}. "
        "The release-gate source-tree entrypoint smoke must execute the installed console script, "
        "not a direct Python import fallback."
    )


def _entrypoint_env(tmp_path: Path) -> dict[str, str]:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    env["DOCWEN_CONFIG_DIR"] = str(tmp_path / "config_home")
    env["DOCWEN_LOG_DIR"] = str(tmp_path / "log_home")
    return env


@pytest.mark.release_gate
def test_source_tree_docwen_console_script_help_runs(tmp_path: Path) -> None:
    proc = run_subprocess(
        [*_script_command("docwen"), "--help"],
        cwd=REPO_ROOT,
        env=_entrypoint_env(tmp_path),
        timeout=30,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )

    assert proc.returncode == 0, proc.stderr
    assert "usage: docwen" in proc.stdout
    assert "ModuleNotFoundError" not in proc.stderr
    assert "Failed to load plugin" not in proc.stderr


@pytest.mark.release_gate
def test_source_tree_docwen_gui_console_script_autoclose_runs(tmp_path: Path) -> None:
    env = _entrypoint_env(tmp_path)
    env["DOCWEN_GUI_TEST_AUTOCLOSE_MS"] = "1000"
    env["DOCWEN_GUI_DISABLE_CONTROL"] = "1"
    if sys.platform == "linux" and not env.get("DISPLAY"):
        env["QT_QPA_PLATFORM"] = "offscreen"

    proc = run_subprocess(
        _script_command("docwen-gui"),
        cwd=REPO_ROOT,
        env=env,
        timeout=45,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )

    assert proc.returncode == 0, proc.stderr
    assert "ModuleNotFoundError" not in proc.stderr
    assert "Failed to load plugin" not in proc.stderr
    assert list((tmp_path / "log_home" / "logs").glob("*.log"))
