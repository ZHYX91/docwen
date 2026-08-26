"""Focused contracts for the packaged GUI settings-control gate."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _completed(payload: dict[str, object], *, returncode: int = 0) -> subprocess.CompletedProcess[str]:
    return subprocess.CompletedProcess(
        ["DocWenCLI.exe"],
        returncode,
        stdout=json.dumps(payload),
        stderr="",
    )


def _info_payload() -> dict[str, object]:
    return {
        "protocol_version": 3,
        "success": True,
        "command": "info",
        "data": {
            "capabilities": [
                {
                    "id": "gui.settings",
                    "contract_version": 1,
                    "state": "runtime_check_required",
                    "available": True,
                    "platforms": ["windows"],
                    "current_platform_supported": True,
                    "runtime_check_required": True,
                    "details": {"cold_start": True, "sections": ["proofread"]},
                }
            ]
        },
    }


def _status_payload(*, running: bool) -> dict[str, object]:
    data: dict[str, object] = {
        "state": "running" if running else "stopped",
        "running": running,
        "control_ready": running,
    }
    if running:
        data["pid"] = 42
    return {
        "protocol_version": 3,
        "success": True,
        "command": "gui status",
        "data": data,
    }


def _settings_payload(*, reused: bool) -> dict[str, object]:
    return {
        "protocol_version": 3,
        "success": True,
        "command": "gui open-settings",
        "data": {
            "accepted": True,
            "running": True,
            "action": "open_settings",
            "section": "proofread",
            "reused": reused,
        },
    }


def test_packaged_settings_gate_requires_exact_info_contract() -> None:
    from scripts.release import verify_packaged_gui

    result = verify_packaged_gui._verify_gui_settings_info_response(_completed(_info_payload()))

    assert result["id"] == "gui.settings"
    assert result["contract_version"] == 1
    assert result["runtime_check_required"] is True
    assert result["details"] == {"cold_start": True, "sections": ["proofread"]}


@pytest.mark.parametrize(
    ("path", "replacement"),
    [
        (("protocol_version",), 2),
        (("data", "capabilities", 0, "contract_version"), 2),
        (("data", "capabilities", 0, "runtime_check_required"), False),
        (("data", "capabilities", 0, "details", "sections"), ["general"]),
    ],
)
def test_packaged_settings_gate_rejects_info_contract_drift(
    path: tuple[str | int, ...],
    replacement: object,
) -> None:
    from scripts.release import verify_packaged_gui

    payload = _info_payload()
    target: object = payload
    for part in path[:-1]:
        target = target[part]  # type: ignore[index]
    target[path[-1]] = replacement  # type: ignore[index]

    with pytest.raises(RuntimeError, match="packaged_gui_settings_info_contract_mismatch"):
        verify_packaged_gui._verify_gui_settings_info_response(_completed(payload))


def test_packaged_settings_gate_requires_no_preexisting_gui() -> None:
    from scripts.release import verify_packaged_gui

    verify_packaged_gui._verify_gui_stopped_response(_completed(_status_payload(running=False)))

    with pytest.raises(RuntimeError, match="packaged_gui_control_not_stopped_before_cold_start"):
        verify_packaged_gui._verify_gui_stopped_response(_completed(_status_payload(running=True)))


@pytest.mark.parametrize(
    ("path", "replacement"),
    [
        (("protocol_version",), 2),
        (("command",), "gui open"),
        (("data", "state"), "stopped"),
        (("data", "pid"), True),
        (("data", "pid"), "42"),
        (("data", "pid"), 0),
    ],
)
def test_control_ready_poll_rejects_contract_drift_before_valid_response(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    path: tuple[str, ...],
    replacement: object,
) -> None:
    from scripts.release import verify_packaged_gui

    cli_path = tmp_path / "DocWenCLI.exe"
    cli_path.write_bytes(b"cli")
    invalid = _status_payload(running=True)
    target: object = invalid
    for part in path[:-1]:
        target = target[part]  # type: ignore[index]
    target[path[-1]] = replacement  # type: ignore[index]
    responses = [_completed(invalid), _completed(_status_payload(running=True))]
    calls: list[tuple[str, ...]] = []

    def fake_run(
        binary_path: Path,
        *args: str,
        cwd: Path,
        env: dict[str, str],
        timeout: int,
    ) -> subprocess.CompletedProcess[str]:
        del cwd, env, timeout
        assert binary_path == cli_path
        assert responses
        calls.append(args)
        return responses.pop(0)

    monkeypatch.setattr(verify_packaged_gui, "_run_with_env", fake_run)
    monkeypatch.setattr(verify_packaged_gui.time, "sleep", lambda _seconds: None)

    result = verify_packaged_gui._wait_for_control_ready(
        cli_path=cli_path,
        cwd=tmp_path,
        env={},
        timeout=1,
    )

    assert result["pid"] == 42
    assert calls == [
        ("gui", "status", "--json"),
        ("gui", "status", "--json"),
    ]


def test_ipc_smoke_starts_with_cli_open_settings_and_then_proves_reuse(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "candidate"
    binary_dir.mkdir()
    gui_path = binary_dir / verify_packaged_gui._default_binary_name()
    cli_path = binary_dir / verify_packaged_gui._default_cli_binary_name()
    gui_path.write_bytes(b"gui")
    cli_path.write_bytes(b"cli")
    calls: list[tuple[str, ...]] = []

    def fake_run(
        binary_path: Path,
        *args: str,
        cwd: Path,
        env: dict[str, str],
        timeout: int,
    ) -> subprocess.CompletedProcess[str]:
        del cwd, env, timeout
        assert binary_path == cli_path
        calls.append(args)
        if args == ("info", "--json"):
            return _completed(_info_payload())
        if args == ("gui", "status", "--json"):
            return _completed(_status_payload(running=False))
        if args[:2] == ("gui", "open-settings"):
            return _completed(_settings_payload(reused=len([call for call in calls if call[:2] == args[:2]]) > 1))
        if args[:2] == ("gui", "open"):
            return _completed({"protocol_version": 3, "success": True, "command": "gui open", "data": {}})
        raise AssertionError(args)

    monkeypatch.setattr(verify_packaged_gui, "_run_with_env", fake_run)
    monkeypatch.setattr(
        verify_packaged_gui,
        "_wait_for_control_ready",
        lambda **_kwargs: {
            "running": True,
            "control_ready": True,
            "pid": 42,
            "supported_actions": ["status", "activate", "open", "open_settings"],
            "settings_sections": ["proofread"],
        },
    )
    monkeypatch.setattr(verify_packaged_gui, "_wait_for_ipc_report_ready", lambda **_kwargs: None)
    monkeypatch.setattr(verify_packaged_gui, "_wait_for_control_stopped", lambda **_kwargs: None)
    monkeypatch.setattr(verify_packaged_gui, "_verify_ipc_smoke_report", lambda _path: None)

    completed = verify_packaged_gui._run_ipc_smoke(gui_path, cwd=tmp_path, binary_dir=binary_dir)

    assert completed.returncode == 0
    assert calls[0] == ("info", "--json")
    assert calls[1] == ("gui", "status", "--json")
    assert calls[2] == (
        "gui",
        "open-settings",
        "--section",
        "proofread",
        "--timeout",
        "30",
        "--json",
    )
    assert calls[3] == ("gui", "open-settings", "--section", "proofread", "--json")
    assert calls[4][:2] == ("gui", "open")


def test_ipc_smoke_never_terminates_a_preexisting_gui(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    binary_dir = tmp_path / "candidate"
    binary_dir.mkdir()
    gui_path = binary_dir / verify_packaged_gui._default_binary_name()
    cli_path = binary_dir / verify_packaged_gui._default_cli_binary_name()
    gui_path.write_bytes(b"gui")
    cli_path.write_bytes(b"cli")
    terminated: list[int] = []

    def fake_run(
        binary_path: Path,
        *args: str,
        cwd: Path,
        env: dict[str, str],
        timeout: int,
    ) -> subprocess.CompletedProcess[str]:
        del cwd, env, timeout
        assert binary_path == cli_path
        if args == ("info", "--json"):
            return _completed(_info_payload())
        if args == ("gui", "status", "--json"):
            return _completed(_status_payload(running=True))
        raise AssertionError("A pre-existing GUI must stop the gate before cold start")

    monkeypatch.setattr(verify_packaged_gui, "_run_with_env", fake_run)
    monkeypatch.setattr(
        verify_packaged_gui,
        "_terminate_test_gui",
        lambda **kwargs: terminated.append(kwargs["expected_pid"]),
    )

    with pytest.raises(RuntimeError, match="packaged_gui_control_not_stopped_before_cold_start"):
        verify_packaged_gui._run_ipc_smoke(gui_path, cwd=tmp_path, binary_dir=binary_dir)

    assert terminated == []


def test_ipc_smoke_cleanup_terminates_only_the_captured_owned_pid(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from scripts.release import verify_packaged_gui

    cli_path = tmp_path / verify_packaged_gui._default_cli_binary_name()
    cli_path.write_bytes(b"cli")
    payload = _status_payload(running=True)
    data = payload["data"]
    assert isinstance(data, dict)
    data["pid"] = 99
    killed: list[tuple[int, object]] = []
    monkeypatch.setattr(
        verify_packaged_gui,
        "_run_with_env",
        lambda *_args, **_kwargs: _completed(payload),
    )
    monkeypatch.setattr(
        verify_packaged_gui.os,
        "kill",
        lambda pid, sig: killed.append((pid, sig)),
    )

    verify_packaged_gui._terminate_test_gui(
        cli_path=cli_path,
        cwd=tmp_path,
        env={},
        expected_pid=42,
    )
    assert killed == []

    data["pid"] = "99"
    verify_packaged_gui._terminate_test_gui(
        cli_path=cli_path,
        cwd=tmp_path,
        env={},
        expected_pid=99,
    )
    assert killed == []

    data["pid"] = 99
    verify_packaged_gui._terminate_test_gui(
        cli_path=cli_path,
        cwd=tmp_path,
        env={},
        expected_pid=99,
    )
    assert killed == [(99, verify_packaged_gui.signal.SIGTERM)]
