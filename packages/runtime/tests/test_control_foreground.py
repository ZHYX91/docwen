"""Foreground permission is handed only to the connected GUI on user actions."""

import ctypes
import json
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from docwen_runtime.control import transport

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("pid_available", [False, True])
def test_foreground_permission_uses_connected_pipe_server_pid(monkeypatch, pid_available):
    events = []

    def get_pid(handle, out):
        assert handle == 321
        out._obj.value = 678
        return pid_available

    get_server_pid = Mock(side_effect=get_pid)
    allow = Mock(side_effect=lambda pid: events.append(pid))
    dlls = {
        "kernel32": SimpleNamespace(GetNamedPipeServerProcessId=get_server_pid),
        "user32": SimpleNamespace(AllowSetForegroundWindow=allow),
    }
    monkeypatch.setattr(transport.sys, "platform", "win32")
    monkeypatch.setattr(ctypes, "WinDLL", lambda name, **_kwargs: dlls[name], raising=False)
    channel = transport._WindowsPipeChannel(Mock(fileno=lambda: 321))
    channel.allow_foreground_activation()
    assert events == ([678] if pid_available else [])


@pytest.mark.parametrize(
    "action, grants", [("open", True), ("activate", True), ("open_settings", True), ("status", False)]
)
def test_foreground_grant_precedes_open_request_only(monkeypatch, action, grants):
    events = []
    request = {}
    channel = transport._WindowsPipeChannel(Mock())

    def send(_channel, data, **_kwargs):
        request.update(json.loads(data))
        events.append("send")

    def receive(_channel, *_args, **_kwargs):
        return json.dumps(
            {"control_version": 1, "request_id": request["request_id"], "success": True, "data": {}}
        ).encode()

    monkeypatch.setattr(transport.sys, "platform", "win32")
    monkeypatch.setattr(
        transport._WindowsPipeChannel, "allow_foreground_activation", lambda _channel: events.append("grant")
    )
    monkeypatch.setattr(transport._WindowsPipeChannel, "send_bytes", send)
    monkeypatch.setattr(transport._WindowsPipeChannel, "recv_bytes", receive)
    client = transport.ControlClient()
    monkeypatch.setattr(client, "_connect", lambda _deadline: channel)
    assert client.request(action, timeout=2) == {}
    assert events == (["grant", "send"] if grants else ["send"])
