"""Dependency egress guard contracts."""

from __future__ import annotations

import json
import os
import socket
import stat
import subprocess
import sys
import threading
from collections.abc import Callable
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import cast
from uuid import uuid4

import pytest

pytestmark = pytest.mark.unit


def test_guard_installs_once_without_replacing_socket_constructor() -> None:
    from docwen_runtime.security import dependency_egress_guard, install_dependency_egress_guard

    original = socket.socket
    first = install_dependency_egress_guard()
    second = install_dependency_egress_guard()
    with dependency_egress_guard():
        assert socket.socket is original

    assert first.installed is True
    assert second.installed is True
    assert socket.socket is original


@pytest.mark.parametrize("operation", ["connect", "bind", "sendto"])
def test_guard_blocks_ipv4_io(operation: str) -> None:
    from docwen_runtime.security import NetworkAccessBlockedError, dependency_egress_guard

    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        with dependency_egress_guard(), pytest.raises(NetworkAccessBlockedError) as exc_info:
            if operation == "connect":
                sock.connect(("127.0.0.1", 9))
            elif operation == "bind":
                sock.bind(("127.0.0.1", 0))
            else:
                sock.sendto(b"probe", ("127.0.0.1", 9))
        assert exc_info.value.event == f"socket.{operation}"
        assert exc_info.value.code == "network_access_blocked"
    finally:
        sock.close()


def test_guard_blocks_ipv6_io_when_ipv6_is_available() -> None:
    from docwen_runtime.security import NetworkAccessBlockedError, dependency_egress_guard

    if not socket.has_ipv6:
        pytest.skip("IPv6 is unavailable")
    sock = socket.socket(socket.AF_INET6, socket.SOCK_STREAM)
    try:
        with dependency_egress_guard(), pytest.raises(NetworkAccessBlockedError):
            sock.bind(("::1", 0))
    finally:
        sock.close()


def test_guard_blocks_connect_ex() -> None:
    from docwen_runtime.security import NetworkAccessBlockedError, dependency_egress_guard

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        with dependency_egress_guard(), pytest.raises(NetworkAccessBlockedError) as exc_info:
            sock.connect_ex(("127.0.0.1", 9))
        assert exc_info.value.event == "socket.connect"
    finally:
        sock.close()


def test_guard_reads_native_family_when_socket_subclass_spoofs_attribute() -> None:
    from docwen_runtime.security import NetworkAccessBlockedError, dependency_egress_guard

    class SpoofedFamilySocket(socket.socket):
        @property
        def family(self) -> socket.AddressFamily:
            return socket.AF_UNSPEC

    sock = SpoofedFamilySocket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        assert sock.family == socket.AF_UNSPEC
        assert socket.socket.family.__get__(sock, socket.socket) == socket.AF_INET
        with dependency_egress_guard(), pytest.raises(NetworkAccessBlockedError) as exc_info:
            sock.bind(("127.0.0.1", 0))
        assert exc_info.value.event == "socket.bind"
    finally:
        sock.close()


def test_guard_blocks_sendmsg_when_supported() -> None:
    from docwen_runtime.security import NetworkAccessBlockedError, dependency_egress_guard

    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sendmsg = getattr(sock, "sendmsg", None)
    if not callable(sendmsg):
        sock.close()
        pytest.skip("socket.sendmsg is unavailable on this platform")
    sendmsg_call = cast(Callable[..., int], sendmsg)
    try:
        with dependency_egress_guard(), pytest.raises(NetworkAccessBlockedError) as exc_info:
            sendmsg_call([b"probe"], [], 0, ("127.0.0.1", 9))
        assert exc_info.value.event == "socket.sendmsg"
    finally:
        sock.close()


@pytest.mark.parametrize("host", ["example.invalid", "localhost", "127.0.0.1"])
def test_guard_blocks_all_resolver_use(host: str) -> None:
    from docwen_runtime.security import NetworkAccessBlockedError, dependency_egress_guard

    with dependency_egress_guard(), pytest.raises(NetworkAccessBlockedError) as exc_info:
        socket.getaddrinfo(host, 443)

    assert exc_info.value.event == "socket.getaddrinfo"


def test_guard_preserves_unix_domain_socket() -> None:
    from docwen_runtime.security import dependency_egress_guard

    unix_family = getattr(socket, "AF_UNIX", None)
    if not isinstance(unix_family, int):
        pytest.skip("AF_UNIX is unavailable")
    unix_socket_family = cast(socket.AddressFamily | int, unix_family)
    with TemporaryDirectory(prefix="dw-guard-", dir="/tmp") as temporary_directory:
        endpoint = Path(temporary_directory) / "guard.sock"
        sock = socket.socket(unix_socket_family, socket.SOCK_STREAM)
        try:
            with dependency_egress_guard():
                sock.bind(str(endpoint))
                sock.listen(1)
        finally:
            sock.close()
            endpoint.unlink(missing_ok=True)


def test_guard_preserves_runtime_control_transport(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_runtime.control import ControlClient, ControlServer
    from docwen_runtime.control.transport import _UNIX_SOCKET_PATH_MAX_BYTES, control_endpoint
    from docwen_runtime.security import dependency_egress_guard

    requested_runtime_root: Path | None = None
    if sys.platform != "win32":
        requested_runtime_root = tmp_path / ("overlong-runtime-root-" + "x" * 96)
        requested_runtime_root.mkdir(mode=0o700)
        monkeypatch.setenv("XDG_RUNTIME_DIR", str(requested_runtime_root))
    app_name = f"docwen-egress-test-{uuid4().hex}"
    address, family = control_endpoint(app_name)
    fallback_directory: Path | None = None
    if family == "AF_UNIX":
        assert requested_runtime_root is not None
        fallback_directory = Path(address).parent
        assert fallback_directory != requested_runtime_root
        assert len(os.fsencode(address)) <= _UNIX_SOCKET_PATH_MAX_BYTES
        assert not fallback_directory.exists()
    server = ControlServer(lambda action, payload: {"action": action, "payload": payload}, app_name=app_name)
    try:
        with dependency_egress_guard():
            server.start()
            if fallback_directory is not None:
                metadata = fallback_directory.stat()
                assert stat.S_IMODE(metadata.st_mode) == 0o700
                getuid = getattr(os, "getuid", None)
                if callable(getuid):
                    assert metadata.st_uid == getuid()
            response = ControlClient(app_name=app_name).request("status", {"probe": True}, timeout=2.0)
            assert response == {"action": "status", "payload": {"probe": True}}
            server.stop()
    finally:
        server.stop()
        if fallback_directory is not None:
            assert not fallback_directory.exists()


def test_nested_guard_keeps_process_active_until_outer_exit() -> None:
    from docwen_runtime.security import dependency_egress_guard, dependency_egress_guard_status

    with dependency_egress_guard() as outer:
        assert outer.state == "enforced"
        assert outer.bootstrap == "composition_root"
        with dependency_egress_guard() as inner:
            assert inner.active is True
        assert dependency_egress_guard_status().active is True
    status = dependency_egress_guard_status()

    assert status.installed is True
    assert status.active is False
    assert status.state == "installed"
    assert status.bootstrap == "none"
    assert status.to_dict()["external_processes"] == "not_managed"


def test_guard_is_process_wide_across_threads() -> None:
    from docwen_runtime.security import NetworkAccessBlockedError, dependency_egress_guard

    failures: list[BaseException] = []

    def resolve_in_worker() -> None:
        try:
            socket.getaddrinfo("localhost", 443)
        except BaseException as exc:  # captured for assertion in the owning thread
            failures.append(exc)

    with dependency_egress_guard():
        worker = threading.Thread(target=resolve_in_worker)
        worker.start()
        worker.join(timeout=5.0)

    assert worker.is_alive() is False
    assert len(failures) == 1
    assert isinstance(failures[0], NetworkAccessBlockedError)


def test_process_lifetime_activation_is_idempotent_in_an_isolated_process() -> None:
    script = (
        "import json; "
        "from docwen_runtime.security.network import "
        "activate_process_lifetime_dependency_egress_guard as activate; "
        "first=activate(); second=activate(); "
        "print(json.dumps({'first':first.to_dict(),'second':second.to_dict()}))"
    )
    completed = subprocess.run(
        [sys.executable, "-c", script],
        check=False,
        capture_output=True,
        text=True,
        timeout=15,
    )

    assert completed.returncode == 0, completed.stderr
    payload = json.loads(completed.stdout)
    assert payload["first"]["state"] == "enforced"
    assert payload["first"]["bootstrap"] == "pyinstaller_runtime_hook"
    assert payload["second"] == payload["first"]


def test_guard_deactivates_after_lifecycle_exit() -> None:
    from docwen_runtime.security import dependency_egress_guard

    with dependency_egress_guard():
        pass
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind(("127.0.0.1", 0))
    finally:
        sock.close()


def test_guard_does_not_control_separately_launched_processes() -> None:
    from docwen_runtime.security import dependency_egress_guard

    script = "import json,socket; s=socket.socket(socket.AF_INET,socket.SOCK_STREAM); print(json.dumps({'ok': True})); s.close()"
    with dependency_egress_guard():
        completed = subprocess.run(
            [sys.executable, "-c", script],
            check=False,
            capture_output=True,
            text=True,
            timeout=10,
        )

    assert completed.returncode == 0, completed.stderr
    assert json.loads(completed.stdout) == {"ok": True}
