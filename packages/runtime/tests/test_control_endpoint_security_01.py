"""Focused tests split from test_control_endpoint_security.py."""

from __future__ import annotations

from ._control_endpoint_security_support import (
    Any,
    Path,
    TemporaryDirectory,
    _private_directory,
    errno,
    os,
    pytest,
    socket,
    stat,
    threading,
    time,
    uuid4,
)

pytestmark = pytest.mark.contract


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_deadline_channel_rejects_partial_frame_header() -> None:
    from docwen_runtime.security.network import LocalUnixStreamChannel

    unix_family = getattr(socket, "AF_UNIX", None)
    assert isinstance(unix_family, int)
    local_socket, peer_socket = socket.socketpair(unix_family, socket.SOCK_STREAM)
    channel = LocalUnixStreamChannel.from_file_descriptor(local_socket.fileno())
    try:
        peer_socket.sendall(b"\x00")
        started = time.monotonic()
        with pytest.raises(TimeoutError):
            channel.recv_bytes(256, deadline=started + 0.05)
        assert time.monotonic() - started < 0.5
    finally:
        channel.close()
        local_socket.close()
        peer_socket.close()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_deadline_channel_uses_one_absolute_deadline_across_fragments() -> None:
    from docwen_runtime.security.network import LocalUnixStreamChannel

    unix_family = getattr(socket, "AF_UNIX", None)
    assert isinstance(unix_family, int)
    local_socket, peer_socket = socket.socketpair(unix_family, socket.SOCK_STREAM)
    channel = LocalUnixStreamChannel.from_file_descriptor(local_socket.fileno())

    def _drip_frame_header() -> None:
        for fragment in (b"\x00", b"\x00", b"\x00", b"\x01", b"x"):
            try:
                peer_socket.sendall(fragment)
            except OSError:
                return
            time.sleep(0.03)

    peer = threading.Thread(target=_drip_frame_header, daemon=True)
    peer.start()
    started = time.monotonic()
    try:
        with pytest.raises(TimeoutError):
            channel.recv_bytes(256, deadline=started + 0.07)
        elapsed = time.monotonic() - started
        assert 0.05 <= elapsed < 0.5
    finally:
        channel.close()
        local_socket.close()
        peer_socket.close()
        peer.join(0.5)


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_deadline_channel_rejects_oversized_frame_before_payload() -> None:
    from docwen_runtime.security.network import LocalUnixStreamChannel

    unix_family = getattr(socket, "AF_UNIX", None)
    assert isinstance(unix_family, int)
    local_socket, peer_socket = socket.socketpair(unix_family, socket.SOCK_STREAM)
    channel = LocalUnixStreamChannel.from_file_descriptor(local_socket.fileno())
    try:
        peer_socket.sendall((257).to_bytes(4, "big", signed=True))
        with pytest.raises(OSError, match="exceeds the configured limit"):
            channel.recv_bytes(256, deadline=time.monotonic() + 0.5)
    finally:
        channel.close()
        local_socket.close()
        peer_socket.close()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_deadline_channel_frames_non_byte_buffers_by_nbytes() -> None:
    from array import array

    from docwen_runtime.security.network import LocalUnixStreamChannel

    unix_family = getattr(socket, "AF_UNIX", None)
    assert isinstance(unix_family, int)
    local_socket, peer_socket = socket.socketpair(unix_family, socket.SOCK_STREAM)
    sender = LocalUnixStreamChannel.from_file_descriptor(local_socket.fileno())
    receiver = LocalUnixStreamChannel.from_file_descriptor(peer_socket.fileno())
    payload = array("I", [1, 2, 3])
    try:
        sender.send_bytes(payload, deadline=time.monotonic() + 0.5)
        assert receiver.recv_bytes(64, deadline=time.monotonic() + 0.5) == payload.tobytes()
    finally:
        sender.close()
        receiver.close()
        local_socket.close()
        peer_socket.close()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_client_authentication_stall_obeys_request_deadline(monkeypatch: pytest.MonkeyPatch) -> None:
    from multiprocessing.connection import Listener

    from docwen_runtime.control import ControlClient, ControlTimeoutError
    from docwen_runtime.control.transport import control_endpoint

    with TemporaryDirectory(prefix="dw-auth-", dir="/tmp") as temporary_directory:
        runtime_root = _private_directory(Path(temporary_directory) / "runtime")
        monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
        app_name = f"dw-auth-{uuid4().hex[:8]}"
        address, family = control_endpoint(app_name)
        listener = Listener(address, family=family, authkey=None)
        Path(address).chmod(0o600)
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
            with pytest.raises(ControlTimeoutError) as exc_info:
                ControlClient(app_name=app_name).request("status", timeout=0.1)
            assert accepted.wait(0.5)
            assert time.monotonic() - started < 0.75
            assert str(runtime_root) not in str(exc_info.value)
        finally:
            release.set()
            peer.join(1)
            listener.close()
            Path(address).unlink(missing_ok=True)


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_deadline_client_interoperates_with_stdlib_authenticated_listener(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from multiprocessing.connection import Listener

    from docwen_runtime.control import ControlClient
    from docwen_runtime.control.transport import (
        _PROTOCOL_KEY,
        CONTROL_PROTOCOL_VERSION,
        _decode,
        _encode,
        control_endpoint,
    )

    with TemporaryDirectory(prefix="dw-stdlib-", dir="/tmp") as temporary_directory:
        runtime_root = _private_directory(Path(temporary_directory) / "runtime")
        monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
        app_name = f"dw-stdlib-{uuid4().hex[:8]}"
        address, family = control_endpoint(app_name)
        listener = Listener(address, family=family, authkey=_PROTOCOL_KEY)
        Path(address).chmod(0o600)
        failures: list[BaseException] = []

        def _serve_with_stdlib_connection() -> None:
            try:
                connection = listener.accept()
                try:
                    request = _decode(connection.recv_bytes())
                    connection.send_bytes(
                        _encode(
                            {
                                "control_version": CONTROL_PROTOCOL_VERSION,
                                "request_id": request["request_id"],
                                "success": True,
                                "data": {"interop": "stdlib-listener"},
                            }
                        )
                    )
                finally:
                    connection.close()
            except BaseException as exc:
                failures.append(exc)

        peer = threading.Thread(target=_serve_with_stdlib_connection, daemon=True)
        peer.start()
        try:
            assert ControlClient(app_name=app_name).request("status", timeout=1) == {"interop": "stdlib-listener"}
            peer.join(1)
            assert not peer.is_alive()
            assert not failures
        finally:
            listener.close()
            Path(address).unlink(missing_ok=True)


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_deadline_server_interoperates_with_stdlib_authenticated_client() -> None:
    from multiprocessing.connection import Client

    from docwen_runtime.control import ControlServer
    from docwen_runtime.control.transport import (
        _PROTOCOL_KEY,
        CONTROL_PROTOCOL_VERSION,
        _decode,
        _encode,
    )

    app_name = f"dw-stdlib-server-{uuid4().hex[:8]}"
    server = ControlServer(lambda _action, _payload: {"interop": "stdlib-client"}, app_name=app_name)
    server.start()
    connection = Client(server._address, family="AF_UNIX", authkey=_PROTOCOL_KEY)
    request_id = str(uuid4())
    try:
        connection.send_bytes(
            _encode(
                {
                    "control_version": CONTROL_PROTOCOL_VERSION,
                    "request_id": request_id,
                    "action": "status",
                    "payload": {},
                }
            )
        )
        response = _decode(connection.recv_bytes())
        assert response == {
            "control_version": CONTROL_PROTOCOL_VERSION,
            "request_id": request_id,
            "success": True,
            "data": {"interop": "stdlib-client"},
            "error": None,
        }
    finally:
        connection.close()
        server.stop(timeout=2)


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_endpoint_planning_is_read_only_and_server_owns_lifecycle(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlClient, ControlServer
    from docwen_runtime.control.transport import control_endpoint

    runtime_root = _private_directory(tmp_path / "runtime")
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
    app_name = f"docwen-endpoint-{uuid4().hex}"
    address, family = control_endpoint(app_name)

    assert family == "AF_UNIX"
    assert not Path(address).exists()
    client = ControlClient(app_name=app_name)
    assert not Path(address).exists()

    server = ControlServer(lambda action, _payload: {"action": action}, app_name=app_name)
    server.start()
    try:
        endpoint = Path(address)
        metadata = endpoint.stat()
        assert stat.S_ISSOCK(metadata.st_mode)
        assert stat.S_IMODE(metadata.st_mode) == 0o600
        assert client.request("status", timeout=2.0) == {"action": "status"}
    finally:
        server.stop()

    assert not Path(address).exists()
    assert runtime_root.exists()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_distinct_apps_share_xdg_root_without_endpoint_collision(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlClient, ControlServer
    from docwen_runtime.control.transport import control_endpoint

    runtime_root = _private_directory(Path("/tmp") / f"dw-xdg-{uuid4().hex[:10]}")
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
    app_a = f"docwen-app-a-{uuid4().hex}"
    app_b = f"docwen-app-b-{uuid4().hex}"
    address_a, _family_a = control_endpoint(app_a)
    address_b, _family_b = control_endpoint(app_b)
    server_a = ControlServer(lambda _action, _payload: {"app": "a"}, app_name=app_a)
    server_b = ControlServer(lambda _action, _payload: {"app": "b"}, app_name=app_b)
    try:
        assert address_a != address_b
        server_a.start()
        server_b.start()
        assert ControlClient(app_name=app_a).request("status", timeout=2.0) == {"app": "a"}
        assert ControlClient(app_name=app_b).request("status", timeout=2.0) == {"app": "b"}
    finally:
        server_b.stop()
        server_a.stop()
        runtime_root.rmdir()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_second_server_cannot_unlink_live_same_app_endpoint(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlClient, ControlError, ControlServer

    runtime_root = _private_directory(Path("/tmp") / f"dw-live-{uuid4().hex[:10]}")
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
    app_name = f"docwen-live-{uuid4().hex}"
    first = ControlServer(lambda _action, _payload: {"server": "first"}, app_name=app_name)
    second = ControlServer(lambda _action, _payload: {"server": "second"}, app_name=app_name)
    try:
        first.start()
        with pytest.raises(ControlError) as exc_info:
            second.start()
        assert exc_info.value.code == "gui_control_start_failed"
        assert ControlClient(app_name=app_name).request("status", timeout=2.0) == {"server": "first"}
    finally:
        second.stop()
        first.stop()
        runtime_root.rmdir()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_stale_same_user_socket_is_recovered_by_exact_inode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlClient, ControlServer
    from docwen_runtime.control.transport import control_endpoint

    runtime_root = _private_directory(Path("/tmp") / f"dw-stale-{uuid4().hex[:10]}")
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
    app_name = f"docwen-stale-{uuid4().hex}"
    address, _family = control_endpoint(app_name)
    unix_family = getattr(socket, "AF_UNIX", None)
    assert isinstance(unix_family, int)
    stale = socket.socket(unix_family, socket.SOCK_STREAM)
    stale.bind(address)
    stale.close()
    stale_inode = Path(address).stat().st_ino
    removed_inodes: list[int] = []
    original_unlink = os.unlink

    def track_unlink(path: Any, *, dir_fd: int | None = None) -> None:
        removed_inodes.append(os.stat(path, dir_fd=dir_fd, follow_symlinks=False).st_ino)
        original_unlink(path, dir_fd=dir_fd)

    monkeypatch.setattr(os, "unlink", track_unlink)
    server = ControlServer(lambda _action, _payload: {"recovered": True}, app_name=app_name)
    try:
        server.start()
        assert removed_inodes[0] == stale_inode
        assert ControlClient(app_name=app_name).request("status", timeout=2.0) == {"recovered": True}
    finally:
        server.stop()
        runtime_root.rmdir()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_probe_enoent_never_unlinks_dirfd_visible_socket(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlError, ControlServer, transport

    runtime_root = _private_directory(Path("/tmp") / f"dw-enoent-{uuid4().hex[:10]}")
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
    app_name = f"docwen-enoent-{uuid4().hex}"
    address, _family = transport.control_endpoint(app_name)
    unix_family = getattr(socket, "AF_UNIX", None)
    assert isinstance(unix_family, int)
    stale = socket.socket(unix_family, socket.SOCK_STREAM)
    stale.bind(address)
    stale.close()

    monkeypatch.setattr(
        transport,
        "probe_local_unix_stream_endpoint",
        lambda _address: FileNotFoundError(errno.ENOENT, "injected pathname mismatch"),
    )
    server = ControlServer(lambda _action, _payload: {}, app_name=app_name)
    try:
        with pytest.raises(ControlError) as exc_info:
            server.start(timeout=1.0)
        assert exc_info.value.code == "gui_control_start_failed"
        assert Path(address).is_socket()
    finally:
        server.stop()
        Path(address).unlink()
        runtime_root.rmdir()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
@pytest.mark.parametrize("mode", [0o755, 0o770, 0o777])
def test_explicit_xdg_root_with_unsafe_mode_fails_closed_without_chmod(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    mode: int,
) -> None:
    from docwen_runtime.control import transport

    runtime_root = _private_directory(tmp_path / "unsafe-runtime")
    runtime_root.chmod(mode)
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))

    with pytest.raises(transport.ControlEndpointError) as exc_info:
        transport.control_endpoint(f"docwen-unsafe-{uuid4().hex}")

    assert exc_info.value.code == "gui_control_endpoint_unavailable"
    assert str(runtime_root) not in str(exc_info.value)
    assert stat.S_IMODE(runtime_root.stat().st_mode) == mode


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_explicit_xdg_symlink_fails_closed_without_touching_target(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import transport

    target = _private_directory(tmp_path / "target")
    link = tmp_path / "runtime-link"
    try:
        link.symlink_to(target, target_is_directory=True)
    except OSError as exc:
        pytest.skip(f"directory symlinks unavailable: {exc}")
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(link))

    with pytest.raises(transport.ControlEndpointError):
        transport.control_endpoint(f"docwen-symlink-{uuid4().hex}")

    assert stat.S_IMODE(target.stat().st_mode) == 0o700
    assert list(target.iterdir()) == []


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_selected_fallback_root_does_not_change_when_child_is_unsafe(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import transport

    first = _private_directory(tmp_path / "first")
    second = _private_directory(tmp_path / "second")
    monkeypatch.delenv("XDG_RUNTIME_DIR", raising=False)
    monkeypatch.setattr(transport, "_fallback_root_paths", lambda: (first, second))
    monkeypatch.setattr(transport, "_UNIX_SOCKET_PATH_MAX_BYTES", 4096)
    app_name = f"docwen-fixed-root-{uuid4().hex}"
    namespace = transport._user_namespace()
    app_namespace = transport.hashlib.sha256(app_name.encode()).hexdigest()[:12]
    child = first / f"docwen-{namespace}-{app_namespace}"
    child.write_text("occupied", encoding="utf-8")

    with pytest.raises(transport.ControlEndpointError):
        transport.control_endpoint(app_name)

    assert not any(second.iterdir())


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_managed_fallback_directory_is_private_and_removed_after_stop(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import ControlClient, ControlServer, transport

    long_xdg = _private_directory(tmp_path / ("runtime-" + "x" * 96))
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(long_xdg))
    monkeypatch.setattr(transport, "_fallback_root_paths", lambda: (Path("/tmp"),))
    app_name = f"docwen-managed-{uuid4().hex}"
    address, family = transport.control_endpoint(app_name)
    directory = Path(address).parent
    assert family == "AF_UNIX"
    assert not directory.exists()

    server = ControlServer(lambda _action, _payload: {"ok": True}, app_name=app_name)
    server.start()
    try:
        assert stat.S_IMODE(directory.stat().st_mode) == 0o700
        assert ControlClient(app_name=app_name).request("status", timeout=2.0) == {"ok": True}
    finally:
        server.stop()

    assert not directory.exists()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_client_revalidates_managed_directory_before_connect(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import transport

    first = _private_directory(tmp_path / "first")
    monkeypatch.delenv("XDG_RUNTIME_DIR", raising=False)
    monkeypatch.setattr(transport, "_fallback_root_paths", lambda: (first,))
    monkeypatch.setattr(transport, "_UNIX_SOCKET_PATH_MAX_BYTES", 4096)
    app_name = f"docwen-client-revalidate-{uuid4().hex}"
    client = transport.ControlClient(app_name=app_name)
    directory = Path(client._address).parent
    directory.mkdir(mode=0o755)
    directory.chmod(0o755)

    def _unsafe_connect(
        _channel_type: type[object],
        _address: str,
        *,
        deadline: float,
    ) -> object:
        del deadline
        pytest.fail("unsafe endpoint reached LocalUnixStreamChannel.connect")

    monkeypatch.setattr(
        transport.LocalUnixStreamChannel,
        "connect",
        classmethod(_unsafe_connect),
    )

    with pytest.raises(transport.ControlEndpointError):
        client.request("status", timeout=0.2)
