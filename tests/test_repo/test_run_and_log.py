from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from scripts.release import run_and_log

pytestmark = pytest.mark.unit


def _directory_link(link: Path, target: Path) -> None:
    if os.name == "nt":
        completed = subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(link), str(target.resolve())],
            check=False,
            capture_output=True,
            text=True,
        )
        assert completed.returncode == 0, completed.stderr or completed.stdout
    else:
        link.symlink_to(target, target_is_directory=True)


def _remove_directory_link(link: Path) -> None:
    """Remove a test link without following its target on either platform."""
    if os.name == "nt":
        link.rmdir()
    else:
        link.unlink()


def test_run_logged_command_records_exact_argv_cwd_environment_unicode_bytes_and_exit(tmp_path: Path) -> None:
    output = tmp_path / "command.json"
    code = (
        "import os,sys; "
        "sys.stdout.write('资料|' + os.environ['PYTHONUTF8'] + '|' + os.environ['DOCWEN_LOG_TO_TEMP'] + '\\n'); "
        "sys.stderr.write('错误\\n'); "
        "raise SystemExit(7)"
    )
    argv = [sys.executable, "-c", code]

    exit_code, returned = run_and_log.run_logged_command(
        argv,
        cwd=tmp_path,
        output=output,
        environment_overrides={"PYTHONUTF8": "1", "DOCWEN_LOG_TO_TEMP": ""},
    )

    assert exit_code == 7
    payload = json.loads(output.read_text(encoding="utf-8"))
    assert returned == {"output": str(output.resolve()), **payload}
    assert payload["schemaVersion"] == 1
    assert payload["argv"] == argv
    assert payload["cwd"] == str(tmp_path.resolve())
    assert payload["environmentOverrides"] == {"DOCWEN_LOG_TO_TEMP": "", "PYTHONUTF8": "1"}
    assert payload["exitCode"] == 7
    assert payload["spawnError"] is None
    expected_stdout = f"资料|1|{os.linesep}".encode()
    expected_stderr = f"错误{os.linesep}".encode()
    assert payload["stdout"] == {
        "bytes": len(expected_stdout),
        "sha256": hashlib.sha256(expected_stdout).hexdigest(),
        "encoding": "utf-8",
        "decodingErrors": False,
        "text": expected_stdout.decode(),
    }
    assert payload["stderr"]["text"] == expected_stderr.decode()
    assert payload["stderr"]["sha256"] == hashlib.sha256(expected_stderr).hexdigest()
    assert payload["startedAt"].endswith("Z")
    assert payload["endedAt"].endswith("Z")
    assert payload["durationSeconds"] >= 0


def test_output_is_claimed_before_execution_and_never_overwritten(tmp_path: Path) -> None:
    output = tmp_path / "existing.json"
    output.write_text("preserve me", encoding="utf-8")
    side_effect = tmp_path / "must-not-exist.txt"

    with pytest.raises(run_and_log.CommandLogError, match="output_exists"):
        run_and_log.run_logged_command(
            [sys.executable, "-c", f"from pathlib import Path; Path({str(side_effect)!r}).write_text('ran')"],
            cwd=tmp_path,
            output=output,
        )

    assert output.read_text(encoding="utf-8") == "preserve me"
    assert not side_effect.exists()


@pytest.mark.parametrize(
    ("values", "error"),
    [
        (["PATH=unsafe"], "environment_override_not_allowed"),
        (["PYTHONUTF8"], "invalid_environment_override"),
        (["PYTHONUTF8=1", "PYTHONUTF8=0"], "duplicate_environment_override"),
        (["DOCWEN_CONFIG_DIR=bad\0path"], "invalid_environment_override"),
    ],
)
def test_environment_overrides_are_strictly_allowlisted(values: list[str], error: str) -> None:
    with pytest.raises(run_and_log.CommandLogError, match=error):
        run_and_log.parse_environment_overrides(values)


def test_inherited_environment_is_not_serialized(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DOCWEN_TEST_SECRET_NOT_FOR_LOGS", "secret")
    output = tmp_path / "environment.json"

    exit_code, _ = run_and_log.run_logged_command(
        [sys.executable, "-c", "print('ok')"],
        cwd=tmp_path,
        output=output,
        environment_overrides={"PYTHONIOENCODING": "utf-8"},
    )

    assert exit_code == 0
    serialized = output.read_text(encoding="utf-8")
    assert "DOCWEN_TEST_SECRET_NOT_FOR_LOGS" not in serialized
    assert "secret" not in serialized
    assert json.loads(serialized)["environmentOverrides"] == {"PYTHONIOENCODING": "utf-8"}


def test_non_utf8_output_keeps_byte_identity_and_marks_replacement_text(tmp_path: Path) -> None:
    output = tmp_path / "non-utf8.json"

    exit_code, _ = run_and_log.run_logged_command(
        [sys.executable, "-c", "import sys; sys.stdout.buffer.write(b'\\xff')"],
        cwd=tmp_path,
        output=output,
    )

    assert exit_code == 0
    stdout = json.loads(output.read_text(encoding="utf-8"))["stdout"]
    assert stdout == {
        "bytes": 1,
        "sha256": hashlib.sha256(b"\xff").hexdigest(),
        "encoding": "utf-8",
        "decodingErrors": True,
        "text": "\ufffd",
    }


def test_spawn_failure_is_logged_without_misrepresenting_a_child_exit_code(tmp_path: Path) -> None:
    output = tmp_path / "spawn-failure.json"

    exit_code, _ = run_and_log.run_logged_command(
        [str(tmp_path / "definitely-missing-executable")],
        cwd=tmp_path,
        output=output,
    )

    payload = json.loads(output.read_text(encoding="utf-8"))
    assert exit_code == run_and_log.SPAWN_FAILURE_EXIT_CODE
    assert payload["exitCode"] is None
    assert payload["spawnError"]["type"] in {"FileNotFoundError", "PermissionError", "OSError"}
    assert payload["stderr"]["bytes"] > 0


def test_linked_working_directory_or_output_parent_is_rejected(tmp_path: Path) -> None:
    real_directory = tmp_path / "real"
    real_directory.mkdir()
    linked_directory = tmp_path / "linked"
    _directory_link(linked_directory, real_directory)

    with pytest.raises(run_and_log.CommandLogError, match="linked_path_rejected"):
        run_and_log.run_logged_command(
            [sys.executable, "-c", "print('no')"],
            cwd=linked_directory,
            output=tmp_path / "cwd-link.json",
        )
    with pytest.raises(run_and_log.CommandLogError, match="linked_path_rejected"):
        run_and_log.run_logged_command(
            [sys.executable, "-c", "print('no')"],
            cwd=tmp_path,
            output=linked_directory / "output.json",
        )
    assert not (real_directory / "output.json").exists()
    _remove_directory_link(linked_directory)


def test_cli_writes_the_log_and_returns_the_child_exit_code(tmp_path: Path) -> None:
    output = tmp_path / "cli.json"
    completed = subprocess.run(
        [
            sys.executable,
            "scripts/release/run_and_log.py",
            "--output",
            str(output),
            "--cwd",
            str(tmp_path),
            "--env",
            "PYTHONUTF8=1",
            "--",
            sys.executable,
            "-c",
            "print('CLI 日志'); raise SystemExit(3)",
        ],
        cwd=Path.cwd(),
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )

    assert completed.returncode == 3
    assert completed.stdout == ""
    assert completed.stderr == ""
    payload = json.loads(output.read_text(encoding="utf-8"))
    assert payload["exitCode"] == 3
    assert payload["stdout"]["text"] == f"CLI 日志{os.linesep}"
    assert payload["environmentOverrides"] == {"PYTHONUTF8": "1"}
