"""Focused tests split from test_control_endpoint_security.py."""

from __future__ import annotations

from ._control_endpoint_security_support import (
    Any,
    Path,
    _private_directory,
    errno,
    os,
    pytest,
    socket,
    sys,
    uuid4,
)

pytestmark = pytest.mark.contract


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_startup_socket_hardening_failure_sets_ready_and_cleans_fallback(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import transport

    long_xdg = _private_directory(tmp_path / ("runtime-" + "x" * 96))
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(long_xdg))
    monkeypatch.setattr(transport, "_fallback_root_paths", lambda: (Path("/tmp"),))
    app_name = f"docwen-harden-failure-{uuid4().hex}"
    address, _family = transport.control_endpoint(app_name)
    original_chmod = transport.os.chmod

    def fail_socket_chmod(path: Any, mode: int, **kwargs: Any) -> None:
        if mode == 0o600:
            raise PermissionError("injected socket hardening failure")
        original_chmod(path, mode, **kwargs)

    monkeypatch.setattr(transport.os, "chmod", fail_socket_chmod)
    server = transport.ControlServer(lambda _action, _payload: {}, app_name=app_name)

    with pytest.raises(transport.ControlError) as exc_info:
        server.start(timeout=1.0)

    assert exc_info.value.code == "gui_control_start_failed"
    assert not Path(address).exists()
    assert not Path(address).parent.exists()
    server.stop()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_listener_partial_bind_failure_is_claimed_and_cleaned(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import transport

    runtime_root = _private_directory(Path("/tmp") / f"dw-partial-{uuid4().hex[:10]}")
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
    app_name = f"docwen-partial-{uuid4().hex}"
    address, _family = transport.control_endpoint(app_name)

    def partial_listener(path: str, **_kwargs: object) -> None:
        unix_family = getattr(socket, "AF_UNIX", None)
        assert isinstance(unix_family, int)
        partial = socket.socket(unix_family, socket.SOCK_STREAM)
        try:
            partial.bind(path)
        finally:
            partial.close()
        raise OSError("injected failure after bind")

    monkeypatch.setattr(transport, "Listener", partial_listener)
    server = transport.ControlServer(lambda _action, _payload: {}, app_name=app_name)
    try:
        with pytest.raises(transport.ControlError) as exc_info:
            server.start(timeout=1.0)
        assert exc_info.value.code == "gui_control_start_failed"
        assert not Path(address).exists()
    finally:
        server.stop()
        runtime_root.rmdir()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_listener_eaddrinuse_race_does_not_claim_competing_socket(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import transport

    runtime_root = _private_directory(Path("/tmp") / f"dw-race-{uuid4().hex[:10]}")
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(runtime_root))
    app_name = f"docwen-race-{uuid4().hex}"
    address, _family = transport.control_endpoint(app_name)
    competitor: list[socket.socket] = []

    def racing_listener(path: str, **_kwargs: object) -> None:
        unix_family = getattr(socket, "AF_UNIX", None)
        assert isinstance(unix_family, int)
        listener = socket.socket(unix_family, socket.SOCK_STREAM)
        listener.bind(path)
        listener.listen(1)
        competitor.append(listener)
        raise OSError(errno.EADDRINUSE, "injected competing listener")

    monkeypatch.setattr(transport, "Listener", racing_listener)
    server = transport.ControlServer(lambda _action, _payload: {}, app_name=app_name)
    try:
        with pytest.raises(transport.ControlError) as exc_info:
            server.start(timeout=1.0)
        assert exc_info.value.code == "gui_control_start_failed"
        assert Path(address).is_socket()
    finally:
        server.stop()
        for listener in competitor:
            listener.close()
        Path(address).unlink()
        runtime_root.rmdir()


@pytest.mark.skipif(os.name == "nt", reason="Unix endpoint contract")
def test_shutdown_cleanup_failure_is_reported_by_stop(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_runtime.control import transport

    long_xdg = _private_directory(tmp_path / ("runtime-" + "x" * 96))
    monkeypatch.setenv("XDG_RUNTIME_DIR", str(long_xdg))
    monkeypatch.setattr(transport, "_fallback_root_paths", lambda: (Path("/tmp"),))
    app_name = f"docwen-cleanup-{uuid4().hex}"
    address, _family = transport.control_endpoint(app_name)
    managed_directory = Path(address).parent
    original_rmdir = transport.os.rmdir

    def fail_managed_rmdir(path: Any, *, dir_fd: int | None = None) -> None:
        if path == managed_directory.name and dir_fd is not None:
            raise OSError(errno.EIO, "injected managed directory cleanup failure")
        original_rmdir(path, dir_fd=dir_fd)

    monkeypatch.setattr(transport.os, "rmdir", fail_managed_rmdir)
    server = transport.ControlServer(lambda _action, _payload: {}, app_name=app_name)
    server.start()

    with pytest.raises(transport.ControlError) as exc_info:
        server.stop(timeout=1.0)

    assert exc_info.value.code == "gui_control_stop_failed"
    assert managed_directory.is_dir()
    original_rmdir(managed_directory)


def test_windows_pipe_planning_never_calls_unix_helper(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_runtime.control import transport

    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr(
        transport,
        "_plan_unix_endpoint",
        lambda *_args, **_kwargs: pytest.fail("Windows called Unix endpoint planning"),
    )
    monkeypatch.setenv("XDG_RUNTIME_DIR", "x" * 500)

    address, family = transport.control_endpoint("docwen")

    assert family == "AF_PIPE"
    assert address.startswith(r"\\.\pipe\docwen-runtime-control-v1-")


def test_client_endpoint_failure_is_typed_and_path_free(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_runtime.control import transport

    private_path = "C:/private/control/path" if os.name == "nt" else "/private/control/path"

    def unavailable(_app_name: str = "docwen") -> tuple[str, str]:
        try:
            raise OSError(private_path)
        except OSError as exc:
            raise transport.ControlEndpointError() from exc

    monkeypatch.setattr(transport, "control_endpoint", unavailable)
    client = transport.ControlClient(app_name="docwen")

    with pytest.raises(transport.ControlEndpointError) as exc_info:
        client.request("status", timeout=0.1)

    assert exc_info.value.code == "gui_control_endpoint_unavailable"
    assert private_path not in str(exc_info.value)


def test_server_endpoint_failure_maps_to_start_error(monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_runtime.control import transport

    if os.name == "nt":
        monkeypatch.setattr(
            transport,
            "control_endpoint",
            lambda _app_name="docwen": (_ for _ in ()).throw(transport.ControlEndpointError()),
        )
    else:
        monkeypatch.setattr(
            transport,
            "_plan_unix_endpoint",
            lambda *_args: (_ for _ in ()).throw(transport.ControlEndpointError()),
        )
    server = transport.ControlServer(lambda _action, _payload: {}, app_name="docwen")

    with pytest.raises(transport.ControlError) as exc_info:
        server.start(timeout=1.0)

    assert exc_info.value.code == "gui_control_start_failed"
    server.stop()
