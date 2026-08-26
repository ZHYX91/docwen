"""Protocol 3 CLI adapter tests for runtime/control."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract


class _Control:
    def __init__(self) -> None:
        self.calls: list[tuple[str, object, float]] = []

    def status(self, *, timeout: float) -> dict[str, object]:
        self.calls.append(("status", None, timeout))
        return {"state": "stopped", "running": False, "control_ready": False, "available": True}

    def activate(self, *, timeout: float) -> dict[str, object]:
        from docwen_cli.gui_control_port import GuiControlError

        self.calls.append(("activate", None, timeout))
        raise GuiControlError("gui_not_running", "DocWen GUI is not running.")

    def open(self, file_path: str | None, *, timeout: float) -> dict[str, object]:
        self.calls.append(("open", file_path, timeout))
        return {"accepted": True, "running": True, "action": "open", "file": file_path}

    def open_settings(self, section: str, *, timeout: float) -> dict[str, object]:
        self.calls.append(("open_settings", section, timeout))
        return {
            "accepted": True,
            "running": True,
            "action": "open_settings",
            "section": section,
            "reused": False,
        }


def test_gui_status_is_read_only_success_when_stopped(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    control = _Control()
    exit_code = main(["gui", "status", "--json"], gui_control_port_factory=lambda: control)

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert exit_code == 0
    assert captured.err == ""
    assert payload["success"] is True
    assert payload["command"] == "gui status"
    assert payload["data"]["state"] == "stopped"
    assert control.calls == [("status", None, 5.0)]


def test_gui_activate_stopped_is_typed_unavailable(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    control = _Control()
    exit_code = main(["gui", "activate", "--json"], gui_control_port_factory=lambda: control)

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert exit_code == 6
    assert captured.err == ""
    assert payload["error"]["category"] == "unavailable"
    assert payload["error"]["code"] == "gui_not_running"
    assert control.calls == [("activate", None, 10.0)]


def test_gui_endpoint_failure_is_typed_unavailable(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.gui_control_port import GuiControlError
    from docwen_cli.main import main

    class _EndpointUnavailableControl(_Control):
        def activate(self, *, timeout: float) -> dict[str, object]:
            self.calls.append(("activate", None, timeout))
            raise GuiControlError(
                "gui_control_endpoint_unavailable",
                "The DocWen GUI control endpoint is unavailable.",
            )

    control = _EndpointUnavailableControl()
    exit_code = main(["gui", "activate", "--json"], gui_control_port_factory=lambda: control)

    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert exit_code == 6
    assert captured.err == ""
    assert payload["error"]["category"] == "unavailable"
    assert payload["error"]["code"] == "gui_control_endpoint_unavailable"
    assert control.calls == [("activate", None, 10.0)]


def test_gui_open_passes_canonical_absolute_path(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    source = tmp_path / "中文 sample.md"
    source.write_text("hello", encoding="utf-8")
    control = _Control()
    exit_code = main(
        ["gui", "open", str(source.resolve()), "--json"],
        gui_control_port_factory=lambda: control,
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert payload["data"]["accepted"] is True
    assert control.calls == [("open", str(source.resolve()), 30.0)]


def test_gui_open_without_file_launches_or_activates(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    control = _Control()
    exit_code = main(["gui", "open", "--json"], gui_control_port_factory=lambda: control)

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert payload["success"] is True
    assert payload["command"] == "gui open"
    assert control.calls == [("open", None, 30.0)]


def test_gui_open_settings_passes_semantic_section(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    control = _Control()
    exit_code = main(
        ["gui", "open-settings", "--section", "proofread", "--timeout", "45", "--json", "--quiet"],
        gui_control_port_factory=lambda: control,
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 0
    assert payload["command"] == "gui open-settings"
    assert payload["data"]["section"] == "proofread"
    assert control.calls == [("open_settings", "proofread", 45.0)]


def test_gui_open_settings_rejects_unknown_section_as_typed_input(
    capsys: pytest.CaptureFixture[str],
) -> None:
    from docwen_cli.main import main

    control = _Control()
    exit_code = main(
        ["gui", "open-settings", "--section", "document", "--json"],
        gui_control_port_factory=lambda: control,
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 2
    assert payload["error"]["code"] == "invalid_arguments"
    assert control.calls == []


def test_gui_open_settings_preserves_typed_unavailable_exit(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.gui_control_port import GuiControlError
    from docwen_cli.main import main

    class _UnavailableSettingsControl(_Control):
        def open_settings(self, section: str, *, timeout: float) -> dict[str, object]:
            self.calls.append(("open_settings", section, timeout))
            raise GuiControlError(
                "settings_section_unavailable",
                "The requested GUI settings section could not be loaded.",
                details={"section": section},
            )

    control = _UnavailableSettingsControl()
    exit_code = main(
        ["gui", "open-settings", "--section", "proofread", "--json"],
        gui_control_port_factory=lambda: control,
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 6
    assert payload["error"]["category"] == "unavailable"
    assert payload["error"]["code"] == "settings_section_unavailable"
    assert payload["error"]["details"] == {"section": "proofread"}
    assert control.calls == [("open_settings", "proofread", 30.0)]


def test_packaged_adapter_launches_stopped_gui_without_file(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_bundle import gui_control_adapter
    from docwen_bundle.gui_control_adapter import GuiControlAdapter
    from docwen_runtime.control import ControlNotRunningError

    calls: list[tuple[str, object, float]] = []

    class _Client:
        def request(self, action: str, payload: object = None, *, timeout: float) -> dict[str, object]:
            calls.append((action, payload, timeout))
            if len(calls) == 1:
                raise ControlNotRunningError()
            if action == "status":
                return {"running": True, "control_ready": True}
            return {"accepted": True, "running": True, "action": action}

    adapter = GuiControlAdapter()
    adapter._client = _Client()  # type: ignore[assignment]
    launches: list[bool] = []
    monkeypatch.setattr(adapter, "_launch_gui", lambda: launches.append(True))
    monkeypatch.setattr(gui_control_adapter.time, "monotonic", lambda: 100.0)

    result = adapter.open(None, timeout=1)

    assert result == {"accepted": True, "running": True, "action": "activate"}
    assert launches == [True]
    assert calls == [
        ("activate", {"_deadline_monotonic": 101.0}, 1),
        ("status", {"_deadline_monotonic": 100.5}, 0.5),
        ("activate", {"_deadline_monotonic": 101.0}, 1.0),
    ]


def test_packaged_adapter_readiness_attempt_never_exceeds_outer_deadline(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_bundle import gui_control_adapter
    from docwen_bundle.gui_control_adapter import GuiControlAdapter

    calls: list[tuple[str, object, float]] = []

    class _Client:
        def request(self, action: str, payload: object = None, *, timeout: float) -> dict[str, object]:
            calls.append((action, payload, timeout))
            return {"running": True, "control_ready": True}

    clock = iter((100.0, 100.8, 100.9))
    monkeypatch.setattr(gui_control_adapter.time, "monotonic", lambda: next(clock))
    adapter = GuiControlAdapter()
    adapter._client = _Client()  # type: ignore[assignment]

    assert adapter._wait_until_ready(101.0)["running"] is True
    assert len(calls) == 1
    action, payload, timeout = calls[0]
    assert action == "status"
    assert payload == {"_deadline_monotonic": 101.0}
    assert timeout == pytest.approx(0.1)


def test_gui_open_rejects_relative_path_before_control_call(capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_cli.main import main

    control = _Control()
    exit_code = main(["gui", "open", "relative.md", "--json"], gui_control_port_factory=lambda: control)

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 2
    assert payload["error"]["code"] == "invalid_path"
    assert control.calls == []


def test_gui_open_missing_absolute_path_is_typed(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    from docwen_bundle.gui_control_adapter import GuiControlAdapter
    from docwen_cli.main import main

    missing = tmp_path / "不存在.md"
    exit_code = main(
        ["gui", "open", str(missing.resolve()), "--json"],
        gui_control_port_factory=GuiControlAdapter,
    )

    payload = json.loads(capsys.readouterr().out)
    assert exit_code == 2
    assert payload["error"]["code"] == "file_not_found"


def test_packaged_adapter_refuses_old_running_gui_without_starting_second_instance(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_bundle.gui_control_adapter import GuiControlAdapter
    from docwen_cli.gui_control_port import GuiControlError

    calls: list[str] = []

    class _Client:
        def request(self, action: str, payload: object = None, *, timeout: float) -> dict[str, object]:
            del payload, timeout
            calls.append(action)
            return {"running": True, "control_ready": True, "supported_actions": ["status", "open"]}

    adapter = GuiControlAdapter()
    adapter._client = _Client()  # type: ignore[assignment]
    launches: list[bool] = []
    monkeypatch.setattr(adapter, "_launch_gui", lambda: launches.append(True))

    with pytest.raises(GuiControlError) as exc_info:
        adapter.open_settings("proofread", timeout=1)

    assert exc_info.value.code == "capability_unavailable"
    assert exc_info.value.details["restart_required"] is True
    assert calls == ["status"]
    assert launches == []


def test_packaged_adapter_cold_starts_then_uses_advertised_settings_action(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_bundle.gui_control_adapter as adapter_module
    from docwen_bundle.gui_control_adapter import GuiControlAdapter
    from docwen_runtime.control import ControlNotRunningError

    calls: list[tuple[str, object]] = []

    class _Client:
        def request(self, action: str, payload: object = None, *, timeout: float) -> dict[str, object]:
            del timeout
            calls.append((action, payload))
            if len(calls) == 1:
                raise ControlNotRunningError()
            if action == "status":
                return {
                    "running": True,
                    "control_ready": True,
                    "supported_actions": ["status", "open_settings"],
                    "settings_sections": ["proofread"],
                }
            return {
                "accepted": True,
                "running": True,
                "action": "open_settings",
                "section": "proofread",
                "reused": False,
            }

    adapter = GuiControlAdapter()
    adapter._client = _Client()  # type: ignore[assignment]
    launches: list[bool] = []
    monkeypatch.setattr(adapter, "_launch_gui", lambda: launches.append(True))
    monkeypatch.setattr(adapter_module.time, "monotonic", lambda: 100.0)

    result = adapter.open_settings("proofread", timeout=1)

    assert result["accepted"] is True
    assert launches == [True]
    assert calls == [
        ("status", {"_deadline_monotonic": 101.0}),
        ("status", {"_deadline_monotonic": 100.5}),
        (
            "open_settings",
            {
                "section": "proofread",
                "_deadline_monotonic": 101.0,
            },
        ),
    ]


@pytest.mark.parametrize(
    "response",
    [
        {"accepted": False, "running": True, "action": "open_settings", "section": "proofread", "reused": False},
        {"accepted": True, "running": True, "action": "open_settings", "section": "document", "reused": False},
        {"accepted": True, "running": True, "action": "open_settings", "section": "proofread", "reused": 1},
    ],
)
def test_packaged_adapter_rejects_malformed_settings_response(response: dict[str, object]) -> None:
    from docwen_bundle.gui_control_adapter import GuiControlAdapter
    from docwen_cli.gui_control_port import GuiControlError

    class _Client:
        def request(self, action: str, payload: object = None, *, timeout: float) -> dict[str, object]:
            del payload, timeout
            if action == "status":
                return {
                    "running": True,
                    "control_ready": True,
                    "supported_actions": ["status", "open_settings"],
                    "settings_sections": ["proofread"],
                }
            return response

    adapter = GuiControlAdapter()
    adapter._client = _Client()  # type: ignore[assignment]

    with pytest.raises(GuiControlError) as exc_info:
        adapter.open_settings("proofread", timeout=1)

    assert exc_info.value.code == "invalid_control_response"


def test_packaged_adapter_rejects_malformed_running_status() -> None:
    from docwen_bundle.gui_control_adapter import GuiControlAdapter
    from docwen_cli.gui_control_port import GuiControlError

    class _Client:
        def request(self, action: str, payload: object = None, *, timeout: float) -> dict[str, object]:
            del action, payload, timeout
            return {"running": True, "control_ready": False}

    adapter = GuiControlAdapter()
    adapter._client = _Client()  # type: ignore[assignment]

    with pytest.raises(GuiControlError) as exc_info:
        adapter.status(timeout=1)

    assert exc_info.value.code == "invalid_control_response"
