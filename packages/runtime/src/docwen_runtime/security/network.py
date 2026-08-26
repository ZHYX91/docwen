"""Process-lifetime egress protection for in-process dependencies.

DocWen has no product feature that requires IP networking.  The guard in this
module therefore rejects DNS resolution and IPv4/IPv6 I/O while a supported
DocWen entry point is running.  Local transports that do not use IP -- Windows
named pipes and Unix-domain sockets -- remain available.

The implementation uses CPython audit events instead of replacing
``socket.socket``.  It is a defence against accidental dependency egress, not
an operating-system sandbox: native code and separately launched processes are
outside this boundary.  CPython does not emit an audit event for an ordinary
payload write on an already-connected or inherited socket, so that case is
also outside the guarantee.  Supported entry points activate before product
imports and do not intentionally establish such sockets.
"""

from __future__ import annotations

import os
import socket
import struct
import sys
import threading
import time
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from typing import Any, ClassVar, cast

_PROBE_EVENT = "docwen.dependency_egress_guard.probe"
_IP_EVENTS = frozenset(
    {
        "socket.bind",
        "socket.connect",
        "socket.sendmsg",
        "socket.sendto",
    }
)
_DNS_EVENTS = frozenset(
    {
        "socket.getaddrinfo",
        "socket.gethostbyaddr",
        "socket.gethostbyname",
        "socket.gethostbyname_ex",
        "socket.getnameinfo",
    }
)
_IP_FAMILIES = frozenset({socket.AF_INET, socket.AF_INET6})
_SOCKET_FAMILY_DESCRIPTOR = socket.socket.family
_FRAME_HEADER = struct.Struct("!i")
_LONG_FRAME_HEADER = struct.Struct("!Q")


def _remaining_deadline(deadline: float) -> float:
    remaining = deadline - time.monotonic()
    if remaining <= 0:
        raise TimeoutError("Local control transport deadline expired.")
    return remaining


class LocalUnixStreamChannel:
    """Deadline-bounded framed I/O for one local Unix stream socket.

    ``multiprocessing.connection`` uses a four-byte signed network-order frame
    length (and an eight-byte extension for very large frames).  Keeping that
    byte contract lets the runtime share its existing authentication protocol
    while ensuring a peer that sends no frame, or only part of one, cannot
    block the process beyond the caller's monotonic deadline.
    """

    __slots__ = ("_socket",)

    def __init__(self, channel_socket: socket.socket) -> None:
        self._socket = channel_socket

    @classmethod
    def connect(cls, address: str, *, deadline: float) -> LocalUnixStreamChannel:
        unix_family = getattr(socket, "AF_UNIX", None)
        if not isinstance(unix_family, int):
            raise OSError("Unix socket family is unavailable")
        channel = cls(socket.socket(unix_family, socket.SOCK_STREAM))
        try:
            channel._socket.settimeout(_remaining_deadline(deadline))
            channel._socket.connect(address)
            return channel
        except BaseException:
            channel.close()
            raise

    @classmethod
    def from_file_descriptor(cls, file_descriptor: int) -> LocalUnixStreamChannel:
        duplicate = os.dup(file_descriptor)
        try:
            channel_socket = socket.socket(fileno=duplicate)
        except BaseException:
            os.close(duplicate)
            raise
        channel = cls(channel_socket)
        unix_family = getattr(socket, "AF_UNIX", None)
        if (
            not isinstance(unix_family, int)
            or _SOCKET_FAMILY_DESCRIPTOR.__get__(
                channel_socket,
                socket.socket,
            )
            != unix_family
        ):
            channel.close()
            raise OSError("Control connection is not a Unix-domain socket.")
        return channel

    def close(self) -> None:
        self._socket.close()

    def authenticate_client(self, authkey: bytes, *, deadline: float) -> None:
        from multiprocessing.connection import answer_challenge, deliver_challenge

        adapter = _DeadlineFrameAdapter(self, deadline)
        answer_challenge(cast(Any, adapter), authkey)
        deliver_challenge(cast(Any, adapter), authkey)

    def authenticate_server(self, authkey: bytes, *, deadline: float) -> None:
        from multiprocessing.connection import answer_challenge, deliver_challenge

        adapter = _DeadlineFrameAdapter(self, deadline)
        deliver_challenge(cast(Any, adapter), authkey)
        answer_challenge(cast(Any, adapter), authkey)

    def send_bytes(
        self,
        payload: Any,
        *,
        deadline: float,
        offset: int = 0,
        size: int | None = None,
    ) -> None:
        view = memoryview(payload).cast("B")
        if offset < 0 or offset > len(view):
            raise ValueError("offset is outside the payload")
        selected_size = len(view) - offset if size is None else size
        if selected_size < 0 or offset + selected_size > len(view):
            raise ValueError("size is outside the payload")
        selected = view[offset : offset + selected_size]
        if selected_size <= 0x7FFFFFFF:
            header = _FRAME_HEADER.pack(selected_size)
        else:
            header = _FRAME_HEADER.pack(-1) + _LONG_FRAME_HEADER.pack(selected_size)
        self._send_all(header, deadline)
        if selected_size:
            self._send_all(selected, deadline)

    def recv_bytes(self, maxlength: int | None = None, *, deadline: float) -> bytes:
        size = _FRAME_HEADER.unpack(self._recv_exact(_FRAME_HEADER.size, deadline))[0]
        if size == -1:
            size = _LONG_FRAME_HEADER.unpack(self._recv_exact(_LONG_FRAME_HEADER.size, deadline))[0]
        elif size < 0:
            raise OSError("Invalid local control frame length.")
        if maxlength is not None and size > maxlength:
            raise OSError("Local control frame exceeds the configured limit.")
        return self._recv_exact(size, deadline)

    def _send_all(self, payload: bytes | memoryview, deadline: float) -> None:
        try:
            self._socket.settimeout(_remaining_deadline(deadline))
            self._socket.sendall(payload)
        except TimeoutError as exc:
            raise TimeoutError("Local control transport deadline expired.") from exc

    def _recv_exact(self, size: int, deadline: float) -> bytes:
        chunks: list[bytes] = []
        remaining_size = size
        while remaining_size:
            try:
                self._socket.settimeout(_remaining_deadline(deadline))
                chunk = self._socket.recv(remaining_size)
            except TimeoutError as exc:
                raise TimeoutError("Local control transport deadline expired.") from exc
            if not chunk:
                raise EOFError("Local control connection closed during a frame.")
            chunks.append(chunk)
            remaining_size -= len(chunk)
        return b"".join(chunks)


class _DeadlineFrameAdapter:
    __slots__ = ("_channel", "_deadline")

    def __init__(self, channel: LocalUnixStreamChannel, deadline: float) -> None:
        self._channel = channel
        self._deadline = deadline

    def send_bytes(self, payload: Any, offset: int = 0, size: int | None = None) -> None:
        self._channel.send_bytes(
            payload,
            deadline=self._deadline,
            offset=offset,
            size=size,
        )

    def recv_bytes(self, maxlength: int | None = None) -> bytes:
        return self._channel.recv_bytes(maxlength, deadline=self._deadline)


def probe_local_unix_stream_endpoint(
    address: str,
    *,
    timeout: float = 0.2,
) -> OSError | None:
    """Probe one local Unix stream endpoint without permitting IP egress.

    The security owner is the only production module allowed to construct raw
    sockets.  A successful connection returns ``None``; a connection failure
    returns the original :class:`OSError` so the control owner can apply its
    exact stale-endpoint and inode policy.
    """

    unix_family = getattr(socket, "AF_UNIX", None)
    if not isinstance(unix_family, int):
        raise OSError("Unix socket family is unavailable")
    probe = socket.socket(unix_family, socket.SOCK_STREAM)
    try:
        probe.settimeout(timeout)
        try:
            probe.connect(address)
        except OSError as exc:
            return exc
        return None
    finally:
        probe.close()


class NetworkAccessBlockedError(RuntimeError):
    """Raised when an in-process dependency attempts IP/DNS egress."""

    code: ClassVar[str] = "network_access_blocked"
    error_type: ClassVar[str] = "network_access_blocked"

    def __init__(self, event: str) -> None:
        super().__init__(f"DocWen dependency egress guard blocked {event}.")
        self.event = event


class NetworkGuardInstallationError(RuntimeError):
    """Typed fail-closed startup error for an unavailable audit guard."""

    code: ClassVar[str] = "security_check_failed"
    exit_code: ClassVar[int] = 5

    def __init__(self, *, cause: BaseException | None = None) -> None:
        super().__init__("DocWen dependency egress guard could not be enforced.")
        if cause is not None:
            self.__cause__ = cause


@dataclass(frozen=True, slots=True)
class NetworkGuardStatus:
    """Path-free, serializable status for diagnostics and tests."""

    state: str
    installed: bool
    active: bool
    scope: str = "docwen_python_process"
    policy: str = "deny_dns_and_ip"
    mechanism: str = "cpython_audit_hook"
    bootstrap: str = "none"
    local_transports: tuple[str, ...] = ("windows_named_pipe", "unix_domain_socket")
    external_processes: str = "not_managed"

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


_state_lock = threading.RLock()
_hook_installed = False
_active_depth = 0
_process_lifetime_active = False
_bootstrap = "none"
_expected_probe: object | None = None
_observed_probe: object | None = None


def _socket_family(args: tuple[Any, ...]) -> object | None:
    if not args:
        return None
    candidate = args[0]
    try:
        # Audit events provide a real ``socket.socket`` instance, but the
        # instance may be a subclass whose Python-level ``family`` attribute
        # lies about the underlying C socket.  Freeze and invoke the native
        # descriptor directly so an AF_INET/AF_INET6 socket cannot masquerade
        # as a permitted local transport.
        return _SOCKET_FAMILY_DESCRIPTOR.__get__(candidate, socket.socket)
    except (AttributeError, TypeError):
        return None


def _audit_hook(event: str, args: tuple[Any, ...]) -> None:
    global _observed_probe

    if event == _PROBE_EVENT:
        if args and args[0] is _expected_probe:
            _observed_probe = args[0]
        return
    if _active_depth <= 0:
        return
    if event in _DNS_EVENTS:
        raise NetworkAccessBlockedError(event)
    if event in _IP_EVENTS:
        family = _socket_family(args)
        if family is None or family in _IP_FAMILIES:
            raise NetworkAccessBlockedError(event)


def install_dependency_egress_guard() -> NetworkGuardStatus:
    """Install and verify the single process-wide CPython audit hook.

    ``sys.addaudithook`` hooks cannot be removed.  Installation is therefore
    idempotent; enforcement is activated separately for a bounded process
    lifecycle by :func:`dependency_egress_guard`.
    """

    global _expected_probe, _hook_installed, _observed_probe

    with _state_lock:
        if _hook_installed:
            return dependency_egress_guard_status()

        probe = object()
        _expected_probe = probe
        _observed_probe = None
        try:
            sys.addaudithook(_audit_hook)
            sys.audit(_PROBE_EVENT, probe)
        except BaseException as exc:
            _expected_probe = None
            raise NetworkGuardInstallationError(cause=exc) from exc
        if _observed_probe is not probe:
            _expected_probe = None
            raise NetworkGuardInstallationError()
        _expected_probe = None
        _hook_installed = True
        return dependency_egress_guard_status()


def dependency_egress_guard_status() -> NetworkGuardStatus:
    """Return the current process guard state without probing the network."""

    installed = _hook_installed
    active = installed and _active_depth > 0
    state = "enforced" if active else ("installed" if installed else "not_installed")
    return NetworkGuardStatus(
        state=state,
        installed=installed,
        active=active,
        bootstrap=_bootstrap if active else "none",
    )


def activate_process_lifetime_dependency_egress_guard(
    *,
    bootstrap: str = "pyinstaller_runtime_hook",
) -> NetworkGuardStatus:
    """Activate the guard until process exit.

    Frozen applications call this from DocWen's custom PyInstaller runtime
    hook.  PyInstaller executes custom hooks before its built-in hooks, so the
    guard is already active when the Qt runtime hook imports ``PySide6``.  The
    operation is idempotent because the frozen CLI/GUI composition root later
    enters the ordinary nested lifecycle context as well.
    """

    global _active_depth, _bootstrap, _process_lifetime_active

    if bootstrap != "pyinstaller_runtime_hook":
        raise ValueError("unsupported_dependency_egress_guard_bootstrap")
    install_dependency_egress_guard()
    with _state_lock:
        if not _process_lifetime_active:
            _active_depth += 1
            _process_lifetime_active = True
            _bootstrap = bootstrap
        return dependency_egress_guard_status()


@contextmanager
def dependency_egress_guard() -> Iterator[NetworkGuardStatus]:
    """Enforce dependency egress protection for the caller's full lifecycle."""

    global _active_depth, _bootstrap

    install_dependency_egress_guard()
    with _state_lock:
        if _active_depth == 0:
            _bootstrap = "composition_root"
        _active_depth += 1
    try:
        yield dependency_egress_guard_status()
    finally:
        with _state_lock:
            _active_depth = max(0, _active_depth - 1)
            if _active_depth == 0:
                _bootstrap = "none"


__all__ = [
    "LocalUnixStreamChannel",
    "NetworkAccessBlockedError",
    "NetworkGuardInstallationError",
    "NetworkGuardStatus",
    "activate_process_lifetime_dependency_egress_guard",
    "dependency_egress_guard",
    "dependency_egress_guard_status",
    "install_dependency_egress_guard",
    "probe_local_unix_stream_endpoint",
]
