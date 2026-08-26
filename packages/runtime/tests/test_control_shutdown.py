"""Control shutdown must not block a GUI main-thread quit callback."""

from __future__ import annotations

import sys
import threading
import time
from pathlib import Path
from typing import Any
from uuid import uuid4

import pytest

pytestmark = pytest.mark.contract


def test_start_timeout_requests_shutdown_before_delayed_listener_can_serve(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlServer, ControlTimeoutError, transport

    app_name = f"docwen-start-timeout-{uuid4().hex}"
    entered = threading.Event()
    release = threading.Event()
    original_listener = transport.Listener

    def _delayed_listener(*args: Any, **kwargs: Any) -> Any:
        entered.set()
        release.wait(2)
        return original_listener(*args, **kwargs)

    monkeypatch.setattr(transport, "Listener", _delayed_listener)
    server = ControlServer(lambda _action, _payload: {}, app_name=app_name)
    try:
        with pytest.raises(ControlTimeoutError):
            server.start(timeout=0.05)

        assert entered.is_set()
        assert server._stop.is_set()
        thread = server._thread
        assert thread is not None

        release.set()
        thread.join(2)
        assert not thread.is_alive()
        assert server._listener is None
        if sys.platform != "win32":
            assert server._address
            assert not Path(server._address).exists()
            assert server._unix_lease is None
        server.start(timeout=1)
        assert server._serving.is_set()
        server.stop(timeout=1)
        assert server._thread is None
    finally:
        release.set()
        if server._thread is not None and server._thread.is_alive():
            server.begin_stop()
            server._thread.join(2)


def test_begin_stop_is_non_blocking_while_handler_is_active() -> None:
    from docwen_runtime.control import ControlClient, ControlServer

    app_name = f"docwen-stop-{uuid4().hex}"
    entered = threading.Event()
    release = threading.Event()

    def _handler(_action: str, _payload: dict[str, object]) -> dict[str, object]:
        entered.set()
        release.wait(2)
        return {"accepted": True}

    server = ControlServer(_handler, app_name=app_name)
    server.start()
    client_thread = threading.Thread(
        target=lambda: ControlClient(app_name=app_name).request("activate", timeout=2),
        daemon=True,
    )
    client_thread.start()
    assert entered.wait(1)

    started = time.monotonic()
    server.begin_stop()
    assert time.monotonic() - started < 0.1

    release.set()
    client_thread.join(2)
    server.stop()


def test_stop_settles_server_when_peer_stalls_during_authentication() -> None:
    from multiprocessing.connection import Client

    from docwen_runtime.control import ControlServer

    app_name = f"docwen-stalled-auth-{uuid4().hex}"
    server = ControlServer(lambda _action, _payload: {}, app_name=app_name)
    server.start()
    peer = Client(server._address, family=server._family, authkey=None)
    try:
        assert peer.poll(1), "server did not begin the authentication challenge"
        started = time.monotonic()
        server.stop(timeout=2)
        assert time.monotonic() - started < 2
        assert server._thread is None
        assert server._listener is None
        assert server._unix_lease is None
        if sys.platform != "win32":
            assert not Path(server._address).exists()
    finally:
        peer.close()
        if server._thread is not None:
            server.stop(timeout=2)


@pytest.mark.skipif(sys.platform != "win32", reason="Windows named-pipe authentication contract")
def test_pipe_client_authentication_stall_obeys_request_deadline() -> None:
    from multiprocessing.connection import Listener

    from docwen_runtime.control import ControlClient, ControlTimeoutError

    client = ControlClient(app_name=f"docwen-pipe-auth-{uuid4().hex}")
    listener = Listener(client._address, family="AF_PIPE", authkey=None)
    accepted = threading.Event()
    release = threading.Event()

    def _accept_without_authentication() -> None:
        connection = listener.accept()
        try:
            accepted.set()
            release.wait(2)
        finally:
            connection.close()

    peer = threading.Thread(target=_accept_without_authentication, daemon=True)
    peer.start()
    started = time.monotonic()
    try:
        with pytest.raises(ControlTimeoutError):
            client.request("status", timeout=0.1)
        assert accepted.wait(0.5)
        assert time.monotonic() - started < 0.75
    finally:
        release.set()
        peer.join(1)
        listener.close()


@pytest.mark.skipif(sys.platform != "win32", reason="Windows named-pipe write deadline contract")
def test_pipe_client_large_write_stall_obeys_request_deadline() -> None:
    from multiprocessing.connection import Listener

    from docwen_runtime.control import ControlClient, ControlTimeoutError
    from docwen_runtime.control.transport import _PROTOCOL_KEY

    client = ControlClient(app_name=f"docwen-pipe-write-{uuid4().hex}")
    listener = Listener(client._address, family="AF_PIPE", authkey=_PROTOCOL_KEY)
    accepted = threading.Event()
    release = threading.Event()
    peer_errors: list[BaseException] = []

    def _authenticate_without_reading_request() -> None:
        try:
            connection = listener.accept()
            try:
                accepted.set()
                release.wait(2)
            finally:
                connection.close()
        except BaseException as exc:  # pragma: no cover - asserted below
            peer_errors.append(exc)

    peer = threading.Thread(target=_authenticate_without_reading_request, daemon=True)
    peer.start()
    started = time.monotonic()
    try:
        with pytest.raises(ControlTimeoutError) as exc_info:
            client.request("status", {"blob": "x" * 900_000}, timeout=0.1)
        assert accepted.wait(0.5)
        assert time.monotonic() - started < 0.75
        assert isinstance(exc_info.value.__cause__, TimeoutError)
    finally:
        release.set()
        peer.join(1)
        listener.close()
    assert not peer_errors


def test_control_client_maps_handshake_eof_to_not_running(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_runtime.control import ControlClient, ControlNotRunningError, transport

    client = ControlClient(app_name=f"docwen-handshake-{uuid4().hex}")
    client._family = "AF_UNIX"

    def _closed_handshake(*_args: object, **_kwargs: object) -> object:
        raise EOFError("listener closed during authentication")

    monkeypatch.setattr(
        transport.LocalUnixStreamChannel,
        "connect",
        classmethod(_closed_handshake),
    )
    if sys.platform != "win32":
        monkeypatch.setattr(transport, "_validate_unix_client_endpoint", lambda _plan: None)

    with pytest.raises(ControlNotRunningError) as exc_info:
        client.request("status", timeout=0.2)

    assert exc_info.value.code == "gui_not_running"
    assert isinstance(exc_info.value.__cause__, EOFError)


def test_stop_accepts_closed_handshake_only_after_server_thread_exits(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlClient, ControlNotRunningError, ControlServer

    app_name = f"docwen-stop-race-{uuid4().hex}"
    server = ControlServer(lambda _action, _payload: {}, app_name=app_name)
    server.start()
    thread = server._thread
    original_request = ControlClient.request

    def _wake_then_report_closed(
        client: ControlClient,
        action: str,
        payload: dict[str, object] | None = None,
        *,
        timeout: float,
    ) -> dict[str, object]:
        original_request(client, action, payload, timeout=timeout)
        raise ControlNotRunningError("Control listener closed after the stop wake-up.")

    monkeypatch.setattr(ControlClient, "request", _wake_then_report_closed)

    server.stop(timeout=1)
    server.stop(timeout=1)

    assert thread is not None
    assert not thread.is_alive()
    assert server._thread is None
    assert server._listener is None


def test_stop_keeps_live_thread_when_closed_handshake_did_not_wake_it(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import (
        ControlClient,
        ControlNotRunningError,
        ControlServer,
        ControlTimeoutError,
    )

    app_name = f"docwen-stop-live-{uuid4().hex}"
    server = ControlServer(lambda _action, _payload: {}, app_name=app_name)
    server.start()
    thread = server._thread

    def _closed_without_wake(*_args: object, **_kwargs: object) -> dict[str, object]:
        raise ControlNotRunningError()

    with monkeypatch.context() as patch:
        patch.setattr(ControlClient, "request", _closed_without_wake)
        with pytest.raises(ControlTimeoutError):
            server.stop(timeout=0.05)

    assert thread is not None
    assert thread.is_alive()
    assert server._thread is thread
    assert server._listener is not None

    server.stop(timeout=1)
    assert not thread.is_alive()
    assert server._thread is None


def test_stop_propagates_unknown_wake_error_after_settling_server(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlClient, ControlServer

    app_name = f"docwen-stop-error-{uuid4().hex}"
    server = ControlServer(lambda _action, _payload: {}, app_name=app_name)
    server.start()
    thread = server._thread
    original_request = ControlClient.request

    def _wake_then_fail(
        client: ControlClient,
        action: str,
        payload: dict[str, object] | None = None,
        *,
        timeout: float,
    ) -> dict[str, object]:
        original_request(client, action, payload, timeout=timeout)
        raise RuntimeError("unexpected wake failure")

    monkeypatch.setattr(ControlClient, "request", _wake_then_fail)

    with pytest.raises(RuntimeError, match="unexpected wake failure"):
        server.stop(timeout=1)

    assert thread is not None
    assert not thread.is_alive()
    assert server._thread is None
    assert server._listener is None
    server.stop(timeout=1)
