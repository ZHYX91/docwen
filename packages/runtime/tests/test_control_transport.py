from __future__ import annotations

from uuid import uuid4

import pytest

pytestmark = pytest.mark.contract


def _app_name() -> str:
    return f"docwen-test-{uuid4().hex}"


def test_control_server_roundtrip_uses_typed_json_contract() -> None:
    from docwen_runtime.control import ControlClient, ControlServer

    app_name = _app_name()
    seen: list[tuple[str, dict[str, object]]] = []

    def _handler(action: str, payload: dict[str, object]) -> dict[str, object]:
        seen.append((action, payload))
        return {"running": True, "echo": payload}

    server = ControlServer(_handler, app_name=app_name)
    server.start()
    try:
        data = ControlClient(app_name=app_name).request("status", {"text": "中文"}, timeout=2)
    finally:
        server.stop()

    assert data == {"running": True, "echo": {"text": "中文"}}
    assert seen == [("status", {"text": "中文"})]


def test_control_transport_accepts_open_settings_without_protocol_bump() -> None:
    from docwen_runtime.control import CONTROL_PROTOCOL_VERSION, ControlClient, ControlServer

    app_name = _app_name()
    seen: list[tuple[str, dict[str, object]]] = []
    server = ControlServer(
        lambda action, payload: seen.append((action, payload)) or {"accepted": True},
        app_name=app_name,
    )
    server.start()
    try:
        data = ControlClient(app_name=app_name).request(
            "open_settings",
            {"section": "proofread"},
            timeout=2,
        )
    finally:
        server.stop()

    assert CONTROL_PROTOCOL_VERSION == 1
    assert data == {"accepted": True}
    assert seen == [("open_settings", {"section": "proofread"})]


def test_control_client_distinguishes_gui_not_running() -> None:
    from docwen_runtime.control import ControlClient, ControlNotRunningError

    with pytest.raises(ControlNotRunningError) as exc_info:
        ControlClient(app_name=_app_name()).request("status", timeout=0.2)

    assert exc_info.value.code == "gui_not_running"
    assert exc_info.value.details == {"running": False}


def test_control_client_clamps_transport_deadline_to_payload(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import transport

    client = transport.ControlClient(app_name=_app_name())
    captured_deadlines: list[float] = []

    class _ConnectObserved(RuntimeError):
        pass

    def _capture_connect(deadline: float) -> object:
        captured_deadlines.append(deadline)
        raise _ConnectObserved

    monkeypatch.setattr(transport.time, "monotonic", lambda: 100.0)
    monkeypatch.setattr(transport, "_validate_unix_client_endpoint", lambda _plan: None)
    monkeypatch.setattr(client, "_connect", _capture_connect)

    with pytest.raises(_ConnectObserved):
        client.request(
            "status",
            {"_deadline_monotonic": 101.0},
            timeout=30.0,
        )

    assert captured_deadlines == [101.0]


def test_control_client_rejects_expired_payload_deadline_before_connect(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlTimeoutError, transport

    client = transport.ControlClient(app_name=_app_name())
    connect_calls: list[float] = []
    monkeypatch.setattr(transport.time, "monotonic", lambda: 110.0)
    monkeypatch.setattr(client, "_connect", lambda deadline: connect_calls.append(deadline))

    with pytest.raises(ControlTimeoutError) as exc_info:
        client.request(
            "status",
            {"_deadline_monotonic": 101.0},
            timeout=30.0,
        )

    assert exc_info.value.code == "control_timeout"
    assert connect_calls == []


@pytest.mark.parametrize("invalid_deadline", [True, "101", float("nan"), float("inf"), 10**1000])
def test_control_client_rejects_invalid_payload_deadline(
    invalid_deadline: object,
) -> None:
    from docwen_runtime.control import ControlProtocolError, transport

    with pytest.raises(ControlProtocolError) as exc_info:
        transport.ControlClient(app_name=_app_name()).request(
            "status",
            {"_deadline_monotonic": invalid_deadline},
            timeout=30.0,
        )

    assert exc_info.value.code == "invalid_control_deadline"


def test_control_server_preserves_typed_handler_error() -> None:
    from docwen_runtime.control import ControlClient, ControlRemoteError, ControlRequestError, ControlServer

    app_name = _app_name()

    def _handler(_action: str, _payload: dict[str, object]) -> dict[str, object]:
        raise ControlRequestError("gui_command_failed", "Unable to activate.", details={"accepted": False})

    server = ControlServer(_handler, app_name=app_name)
    server.start()
    try:
        with pytest.raises(ControlRemoteError) as exc_info:
            ControlClient(app_name=app_name).request("activate", timeout=2)
    finally:
        server.stop()

    assert exc_info.value.code == "gui_command_failed"
    assert exc_info.value.details == {"accepted": False}


def test_control_rejects_incompatible_protocol_with_typed_details() -> None:
    from docwen_runtime.control.transport import ControlProtocolError, _validate_request

    with pytest.raises(ControlProtocolError) as exc_info:
        _validate_request({"control_version": 999, "request_id": "request", "action": "status", "payload": {}})

    assert exc_info.value.code == "unsupported_control_protocol"
    assert exc_info.value.details == {"expected": 1, "received": 999}


def test_control_server_hides_unhandled_handler_exception() -> None:
    from docwen_runtime.control import ControlClient, ControlRemoteError, ControlServer

    app_name = _app_name()

    def _handler(_action: str, _payload: dict[str, object]) -> dict[str, object]:
        raise RuntimeError("private path C:/secret must not cross the boundary")

    server = ControlServer(_handler, app_name=app_name)
    server.start()
    try:
        with pytest.raises(ControlRemoteError) as exc_info:
            ControlClient(app_name=app_name).request("activate", timeout=2)
    finally:
        server.stop()

    assert exc_info.value.code == "gui_command_failed"
    assert "secret" not in str(exc_info.value)
