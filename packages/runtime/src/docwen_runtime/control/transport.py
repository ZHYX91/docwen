"""Bounded local JSON transport for GUI runtime/control.

Windows uses a named pipe and Unix uses a user-private Unix-domain socket.
No status file, command directory, process-name probe, or shell transport is
part of this protocol.  The fixed multiprocessing key rejects unrelated
protocol clients but is not a user secret or an authentication boundary.
"""

from __future__ import annotations

import contextlib
import errno
import getpass
import hashlib
import json
import math
import os
import stat
import sys
import threading
import time
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from multiprocessing import AuthenticationError
from multiprocessing.connection import Listener
from pathlib import Path
from typing import Any, Protocol, cast
from uuid import uuid4

from docwen_runtime.security.network import LocalUnixStreamChannel, probe_local_unix_stream_endpoint

CONTROL_PROTOCOL_VERSION = 1
_MAX_MESSAGE_BYTES = 1024 * 1024
_PROTOCOL_KEY = hashlib.sha256(b"docwen/runtime-control/v1").digest()
_UNIX_SOCKET_PATH_MAX_BYTES = 103
_PEER_HANDSHAKE_TIMEOUT_SECONDS = 1.0
_PEER_FRAME_TIMEOUT_SECONDS = 1.0


class _PipeConnection(Protocol):
    def close(self) -> None: ...

    def fileno(self) -> int: ...

    def poll(self, timeout: float = 0.0) -> bool: ...

    def recv_bytes(self, maxlength: int | None = None) -> bytes: ...


class _WindowsPipeChannel:
    """Deadline-bounded message I/O over one Windows named pipe."""

    __slots__ = ("_connection",)

    def __init__(self, connection: _PipeConnection) -> None:
        self._connection = connection

    def close(self) -> None:
        self._connection.close()

    def authenticate_client(self, *, deadline: float) -> None:
        from multiprocessing.connection import answer_challenge, deliver_challenge

        adapter = _WindowsPipeAuthAdapter(self, deadline)
        answer_challenge(cast(Any, adapter), _PROTOCOL_KEY)
        deliver_challenge(cast(Any, adapter), _PROTOCOL_KEY)

    def authenticate_server(self, *, deadline: float) -> None:
        from multiprocessing.connection import answer_challenge, deliver_challenge

        adapter = _WindowsPipeAuthAdapter(self, deadline)
        deliver_challenge(cast(Any, adapter), _PROTOCOL_KEY)
        answer_challenge(cast(Any, adapter), _PROTOCOL_KEY)

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
        self._write_message(selected, deadline)

    def recv_bytes(self, maxlength: int | None = None, *, deadline: float) -> bytes:
        remaining = deadline - time.monotonic()
        if remaining <= 0 or not self._connection.poll(remaining):
            raise TimeoutError("Windows control pipe deadline expired.")
        return self._connection.recv_bytes(maxlength)

    def _write_message(self, payload: memoryview, deadline: float) -> None:
        if sys.platform != "win32":
            raise OSError("Windows control pipes are unavailable on this platform.")
        import _winapi

        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise TimeoutError("Windows control pipe deadline expired.")
        overlapped, error = _winapi.WriteFile(
            self._connection.fileno(),
            payload,
            overlapped=True,
        )
        if error == _winapi.ERROR_IO_PENDING:
            wait_ms = max(1, int(remaining * 1000 + 0.999))
            wait_result = _winapi.WaitForMultipleObjects([overlapped.event], False, wait_ms)
            if wait_result == _winapi.WAIT_TIMEOUT:
                overlapped.cancel()
                with contextlib.suppress(OSError):
                    overlapped.GetOverlappedResult(True)
                raise TimeoutError("Windows control pipe deadline expired.")
            if wait_result != _winapi.WAIT_OBJECT_0:
                overlapped.cancel()
                with contextlib.suppress(OSError):
                    overlapped.GetOverlappedResult(True)
                raise OSError("Windows control pipe wait failed.")
        written, result_error = overlapped.GetOverlappedResult(True)
        if result_error == _winapi.ERROR_OPERATION_ABORTED:
            raise TimeoutError("Windows control pipe deadline expired.")
        if result_error != 0 or written != len(payload):
            raise OSError("Windows control pipe write failed.")


class _WindowsPipeAuthAdapter:
    __slots__ = ("_channel", "_deadline")

    def __init__(self, channel: _WindowsPipeChannel, deadline: float) -> None:
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


class ControlError(RuntimeError):
    """Base class for local control failures with a stable machine code."""

    def __init__(self, code: str, message: str, *, details: Mapping[str, Any] | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.details = dict(details or {})


class ControlNotRunningError(ControlError):
    def __init__(self, message: str = "DocWen GUI is not running.") -> None:
        super().__init__("gui_not_running", message, details={"running": False})


class ControlTimeoutError(ControlError):
    def __init__(self, message: str = "DocWen GUI control timed out.") -> None:
        super().__init__("control_timeout", message)


class ControlProtocolError(ControlError):
    pass


class ControlRemoteError(ControlError):
    pass


class ControlRequestError(ControlError):
    """A typed error raised by the GUI-side request handler."""


class ControlEndpointError(ControlError):
    """The local control endpoint cannot be selected safely."""

    def __init__(self) -> None:
        super().__init__(
            "gui_control_endpoint_unavailable",
            "The DocWen GUI control endpoint is unavailable.",
        )


def _user_namespace() -> str:
    identity = f"{getpass.getuser()}|{Path.home()}".encode("utf-8", errors="replace")
    return hashlib.sha256(identity).hexdigest()[:16]


@dataclass(frozen=True, slots=True)
class _UnixEndpointPlan:
    root: Path
    root_device: int
    root_inode: int
    root_policy: str
    directory_name: str | None
    socket_name: str
    address: Path


@dataclass(slots=True)
class _UnixEndpointLease:
    plan: _UnixEndpointPlan
    root_fd: int
    directory_fd: int
    directory_device: int
    directory_inode: int
    socket_device: int | None = None
    socket_inode: int | None = None


def _unix_directory_open_flags() -> int:
    directory = getattr(os, "O_DIRECTORY", None)
    nofollow = getattr(os, "O_NOFOLLOW", None)
    if not isinstance(directory, int) or not isinstance(nofollow, int):
        raise OSError("secure directory descriptors are unavailable")
    return os.O_RDONLY | directory | nofollow | int(getattr(os, "O_CLOEXEC", 0))


def _current_uid() -> int:
    getuid = getattr(os, "getuid", None)
    if not callable(getuid):
        raise OSError("Unix user identity is unavailable")
    uid = getuid()
    if not isinstance(uid, int):
        raise OSError("Unix user identity is invalid")
    return uid


def _open_verified_unix_root(path: Path, *, required_policy: str | None = None) -> tuple[int, os.stat_result, str]:
    fd = os.open(path, _unix_directory_open_flags())
    try:
        metadata = os.fstat(fd)
        if not stat.S_ISDIR(metadata.st_mode):
            raise OSError("control endpoint root is not a directory")
        mode = stat.S_IMODE(metadata.st_mode)
        uid = _current_uid()
        if metadata.st_uid == uid and mode == 0o700:
            policy = "user-private"
        elif metadata.st_uid == 0 and bool(metadata.st_mode & stat.S_ISVTX) and mode & 0o003 == 0o003:
            policy = "root-owned-sticky"
        else:
            raise PermissionError("control endpoint root has unsafe ownership or permissions")
        if required_policy is not None and policy != required_policy:
            raise PermissionError("control endpoint root policy changed")
        return fd, metadata, policy
    except BaseException:
        os.close(fd)
        raise


def _fallback_root_paths() -> tuple[Path, ...]:
    return (Path("/tmp"), Path("/var/tmp"))


def _validate_existing_managed_directory(plan: _UnixEndpointPlan) -> None:
    if plan.directory_name is None:
        return
    root_fd, metadata, _policy = _open_verified_unix_root(
        plan.root,
        required_policy=plan.root_policy,
    )
    try:
        if (metadata.st_dev, metadata.st_ino) != (plan.root_device, plan.root_inode):
            raise OSError("control endpoint root identity changed")
        try:
            directory_fd = os.open(
                plan.directory_name,
                _unix_directory_open_flags(),
                dir_fd=root_fd,
            )
        except FileNotFoundError:
            return
        try:
            child = os.fstat(directory_fd)
            if (
                not stat.S_ISDIR(child.st_mode)
                or child.st_uid != _current_uid()
                or stat.S_IMODE(child.st_mode) != 0o700
            ):
                raise PermissionError("managed control directory is unsafe")
            path_child = os.stat(
                plan.directory_name,
                dir_fd=root_fd,
                follow_symlinks=False,
            )
            if (path_child.st_dev, path_child.st_ino) != (child.st_dev, child.st_ino):
                raise OSError("managed control directory identity changed")
        finally:
            os.close(directory_fd)
    finally:
        os.close(root_fd)


def _plan_unix_endpoint(app_name: str, namespace: str) -> _UnixEndpointPlan:
    app_namespace = hashlib.sha256(app_name.encode("utf-8", errors="replace")).hexdigest()[:12]
    socket_name = f"control-v1-{app_namespace}-{namespace}.sock"
    runtime_root = os.environ.get("XDG_RUNTIME_DIR", "").strip()
    try:
        if runtime_root:
            root = Path(runtime_root)
            if not root.is_absolute():
                raise PermissionError("XDG runtime root must be absolute")
            root_fd, metadata, policy = _open_verified_unix_root(root)
            os.close(root_fd)
            if policy != "user-private":
                raise PermissionError("XDG runtime root must be user-private")
            address = root / socket_name
            if len(os.fsencode(address)) <= _UNIX_SOCKET_PATH_MAX_BYTES:
                return _UnixEndpointPlan(
                    root=root,
                    root_device=metadata.st_dev,
                    root_inode=metadata.st_ino,
                    root_policy=policy,
                    directory_name=None,
                    socket_name=socket_name,
                    address=address,
                )

        directory_name = f"docwen-{namespace}-{app_namespace}"
        attempted: set[Path] = set()
        selected: _UnixEndpointPlan | None = None
        for candidate in _fallback_root_paths():
            try:
                root = candidate.resolve(strict=True)
                if root in attempted:
                    continue
                attempted.add(root)
                root_fd, metadata, policy = _open_verified_unix_root(root)
                os.close(root_fd)
            except OSError:
                continue
            address = root / directory_name / "control-v1.sock"
            if len(os.fsencode(address)) > _UNIX_SOCKET_PATH_MAX_BYTES:
                continue
            selected = _UnixEndpointPlan(
                root=root,
                root_device=metadata.st_dev,
                root_inode=metadata.st_ino,
                root_policy=policy,
                directory_name=directory_name,
                socket_name="control-v1.sock",
                address=address,
            )
            break
        if selected is None:
            raise OSError("no safe short Unix control root is available")
        _validate_existing_managed_directory(selected)
        return selected
    except ControlEndpointError:
        raise
    except OSError as exc:
        raise ControlEndpointError() from exc


def _prepare_unix_server_endpoint(plan: _UnixEndpointPlan) -> _UnixEndpointLease:
    root_fd, root_metadata, _policy = _open_verified_unix_root(
        plan.root,
        required_policy=plan.root_policy,
    )
    directory_fd: int | None = None
    try:
        if (root_metadata.st_dev, root_metadata.st_ino) != (plan.root_device, plan.root_inode):
            raise OSError("control endpoint root identity changed")
        if plan.directory_name is None:
            directory_fd = os.dup(root_fd)
        else:
            with contextlib.suppress(FileExistsError):
                os.mkdir(plan.directory_name, 0o700, dir_fd=root_fd)
            directory_fd = os.open(
                plan.directory_name,
                _unix_directory_open_flags(),
                dir_fd=root_fd,
            )
            child = os.fstat(directory_fd)
            if not stat.S_ISDIR(child.st_mode) or child.st_uid != _current_uid():
                raise PermissionError("managed control directory is unsafe")
            fchmod = getattr(os, "fchmod", None)
            if not callable(fchmod):
                raise OSError("secure descriptor chmod is unavailable")
            fchmod(directory_fd, 0o700)
        directory_metadata = os.fstat(directory_fd)
        if (
            not stat.S_ISDIR(directory_metadata.st_mode)
            or directory_metadata.st_uid != _current_uid()
            or stat.S_IMODE(directory_metadata.st_mode) != 0o700
        ):
            raise PermissionError("control endpoint directory is not user-private")
        if plan.directory_name is not None:
            path_metadata = os.stat(
                plan.directory_name,
                dir_fd=root_fd,
                follow_symlinks=False,
            )
            if (path_metadata.st_dev, path_metadata.st_ino) != (
                directory_metadata.st_dev,
                directory_metadata.st_ino,
            ):
                raise OSError("managed control directory identity changed")
        try:
            stale = os.stat(plan.socket_name, dir_fd=directory_fd, follow_symlinks=False)
        except FileNotFoundError:
            pass
        else:
            if not stat.S_ISSOCK(stale.st_mode) or stale.st_uid != _current_uid():
                raise PermissionError("refusing to replace an unsafe control endpoint")
            probe_error = probe_local_unix_stream_endpoint(str(plan.address))
            if probe_error is None:
                raise FileExistsError("refusing to replace a live control endpoint")
            if probe_error.errno not in {errno.ECONNREFUSED, errno.ENOENT}:
                raise probe_error
            probe_errno = probe_error.errno
            try:
                current = os.stat(
                    plan.socket_name,
                    dir_fd=directory_fd,
                    follow_symlinks=False,
                )
            except FileNotFoundError as exc:
                if probe_errno != errno.ENOENT:
                    raise PermissionError("stale control endpoint disappeared unexpectedly") from exc
            else:
                if probe_errno != errno.ECONNREFUSED:
                    raise PermissionError("control endpoint path resolution is inconsistent")
                if (
                    not stat.S_ISSOCK(current.st_mode)
                    or current.st_uid != _current_uid()
                    or (current.st_dev, current.st_ino) != (stale.st_dev, stale.st_ino)
                ):
                    raise PermissionError("stale control endpoint identity changed")
                os.unlink(plan.socket_name, dir_fd=directory_fd)
        return _UnixEndpointLease(
            plan=plan,
            root_fd=root_fd,
            directory_fd=directory_fd,
            directory_device=directory_metadata.st_dev,
            directory_inode=directory_metadata.st_ino,
        )
    except BaseException:
        if directory_fd is not None:
            os.close(directory_fd)
        os.close(root_fd)
        raise


def _validate_unix_client_endpoint(plan: _UnixEndpointPlan) -> None:
    root_fd, root_metadata, _policy = _open_verified_unix_root(
        plan.root,
        required_policy=plan.root_policy,
    )
    directory_fd: int | None = None
    try:
        if (root_metadata.st_dev, root_metadata.st_ino) != (plan.root_device, plan.root_inode):
            raise OSError("control endpoint root identity changed")
        if plan.directory_name is None:
            directory_fd = os.dup(root_fd)
        else:
            directory_fd = os.open(
                plan.directory_name,
                _unix_directory_open_flags(),
                dir_fd=root_fd,
            )
        directory_metadata = os.fstat(directory_fd)
        if (
            not stat.S_ISDIR(directory_metadata.st_mode)
            or directory_metadata.st_uid != _current_uid()
            or stat.S_IMODE(directory_metadata.st_mode) != 0o700
        ):
            raise PermissionError("control endpoint directory is unsafe")
        if plan.directory_name is not None:
            path_metadata = os.stat(
                plan.directory_name,
                dir_fd=root_fd,
                follow_symlinks=False,
            )
            if (path_metadata.st_dev, path_metadata.st_ino) != (
                directory_metadata.st_dev,
                directory_metadata.st_ino,
            ):
                raise OSError("managed control directory identity changed")
        socket_metadata = os.stat(
            plan.socket_name,
            dir_fd=directory_fd,
            follow_symlinks=False,
        )
        if (
            not stat.S_ISSOCK(socket_metadata.st_mode)
            or socket_metadata.st_uid != _current_uid()
            or stat.S_IMODE(socket_metadata.st_mode) != 0o600
        ):
            raise PermissionError("control endpoint socket is unsafe")
    finally:
        if directory_fd is not None:
            os.close(directory_fd)
        os.close(root_fd)


def _secure_bound_unix_socket(lease: _UnixEndpointLease) -> None:
    metadata = os.stat(
        lease.plan.socket_name,
        dir_fd=lease.directory_fd,
        follow_symlinks=False,
    )
    if not stat.S_ISSOCK(metadata.st_mode) or metadata.st_uid != _current_uid():
        raise PermissionError("bound control endpoint is unsafe")
    lease.socket_device = metadata.st_dev
    lease.socket_inode = metadata.st_ino
    os.chmod(
        lease.plan.socket_name,
        0o600,
        dir_fd=lease.directory_fd,
        follow_symlinks=False,
    )
    verified = os.stat(
        lease.plan.socket_name,
        dir_fd=lease.directory_fd,
        follow_symlinks=False,
    )
    if (
        (verified.st_dev, verified.st_ino) != (metadata.st_dev, metadata.st_ino)
        or not stat.S_ISSOCK(verified.st_mode)
        or verified.st_uid != _current_uid()
        or stat.S_IMODE(verified.st_mode) != 0o600
    ):
        raise PermissionError("bound control endpoint verification failed")
    lease.socket_device = verified.st_dev
    lease.socket_inode = verified.st_ino


def _claim_new_unix_socket_for_cleanup(
    lease: _UnixEndpointLease,
    startup_error: BaseException,
) -> None:
    """Record a socket created by a Listener constructor that later failed."""

    if lease.socket_device is not None or lease.socket_inode is not None:
        return
    if isinstance(startup_error, OSError) and startup_error.errno == errno.EADDRINUSE:
        return
    try:
        metadata = os.stat(
            lease.plan.socket_name,
            dir_fd=lease.directory_fd,
            follow_symlinks=False,
        )
    except FileNotFoundError:
        return
    if not stat.S_ISSOCK(metadata.st_mode) or metadata.st_uid != _current_uid():
        raise PermissionError("failed listener left an unsafe endpoint")
    lease.socket_device = metadata.st_dev
    lease.socket_inode = metadata.st_ino


def _release_unix_server_endpoint(lease: _UnixEndpointLease) -> None:
    failure: BaseException | None = None
    try:
        if lease.socket_device is not None and lease.socket_inode is not None:
            try:
                metadata = os.stat(
                    lease.plan.socket_name,
                    dir_fd=lease.directory_fd,
                    follow_symlinks=False,
                )
            except FileNotFoundError:
                pass
            else:
                if (
                    stat.S_ISSOCK(metadata.st_mode)
                    and metadata.st_uid == _current_uid()
                    and (metadata.st_dev, metadata.st_ino) == (lease.socket_device, lease.socket_inode)
                ):
                    os.unlink(lease.plan.socket_name, dir_fd=lease.directory_fd)
    except BaseException as exc:
        failure = exc
    try:
        os.close(lease.directory_fd)
    except BaseException as exc:
        if failure is None:
            failure = exc
    if lease.plan.directory_name is not None:
        try:
            metadata = os.stat(
                lease.plan.directory_name,
                dir_fd=lease.root_fd,
                follow_symlinks=False,
            )
        except FileNotFoundError:
            pass
        except BaseException as exc:
            if failure is None:
                failure = exc
        else:
            if (
                stat.S_ISDIR(metadata.st_mode)
                and metadata.st_uid == _current_uid()
                and (metadata.st_dev, metadata.st_ino) == (lease.directory_device, lease.directory_inode)
            ):
                try:
                    os.rmdir(lease.plan.directory_name, dir_fd=lease.root_fd)
                except FileNotFoundError:
                    pass
                except OSError as exc:
                    if exc.errno not in {errno.ENOTEMPTY, errno.EEXIST} and failure is None:
                        failure = exc
    try:
        os.close(lease.root_fd)
    except BaseException as exc:
        if failure is None:
            failure = exc
    if failure is not None:
        raise failure


def control_endpoint(app_name: str = "docwen") -> tuple[str, str]:
    """Plan the current user's stable endpoint without creating filesystem state."""

    namespace = _user_namespace()
    if sys.platform == "win32":
        return rf"\\.\pipe\{app_name}-runtime-control-v1-{namespace}", "AF_PIPE"

    return str(_plan_unix_endpoint(app_name, namespace).address), "AF_UNIX"


def _encode(payload: Mapping[str, Any]) -> bytes:
    raw = json.dumps(dict(payload), ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if len(raw) > _MAX_MESSAGE_BYTES:
        raise ControlProtocolError("control_message_too_large", "Control message exceeds the 1 MiB limit.")
    return raw


def _decode(raw: bytes) -> dict[str, Any]:
    if len(raw) > _MAX_MESSAGE_BYTES:
        raise ControlProtocolError("control_message_too_large", "Control message exceeds the 1 MiB limit.")
    try:
        value = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ControlProtocolError("invalid_control_message", "Control message is not valid UTF-8 JSON.") from exc
    if not isinstance(value, dict):
        raise ControlProtocolError("invalid_control_message", "Control message must be a JSON object.")
    return value


def _validate_request(message: Mapping[str, Any]) -> tuple[str, str, dict[str, Any]]:
    version = message.get("control_version")
    if version != CONTROL_PROTOCOL_VERSION:
        raise ControlProtocolError(
            "unsupported_control_protocol",
            f"Unsupported GUI control protocol: {version!r}.",
            details={"expected": CONTROL_PROTOCOL_VERSION, "received": version},
        )
    request_id = message.get("request_id")
    action = message.get("action")
    payload = message.get("payload", {})
    if not isinstance(request_id, str) or not request_id:
        raise ControlProtocolError("invalid_control_message", "request_id must be a non-empty string.")
    if action not in {"status", "activate", "open", "open_settings", "__stop__"}:
        raise ControlProtocolError("invalid_control_action", f"Unsupported GUI control action: {action!r}.")
    if not isinstance(payload, dict):
        raise ControlProtocolError("invalid_control_message", "payload must be a JSON object.")
    return request_id, str(action), payload


def _response(
    request_id: str,
    *,
    success: bool,
    data: Mapping[str, Any] | None = None,
    error: ControlError | None = None,
) -> dict[str, Any]:
    return {
        "control_version": CONTROL_PROTOCOL_VERSION,
        "request_id": request_id,
        "success": success,
        "data": dict(data or {}),
        "error": (
            None if error is None else {"code": error.code, "message": str(error), "details": dict(error.details)}
        ),
    }


class ControlClient:
    """Synchronous bounded client for the current user's GUI control endpoint."""

    def __init__(self, *, app_name: str = "docwen") -> None:
        self._app_name = app_name
        self._endpoint_error: ControlEndpointError | None = None
        try:
            self._address, self._family = control_endpoint(app_name)
        except ControlEndpointError as exc:
            self._address = ""
            self._family = "AF_PIPE" if sys.platform == "win32" else "AF_UNIX"
            self._endpoint_error = exc

    def request(self, action: str, payload: Mapping[str, Any] | None = None, *, timeout: float) -> dict[str, Any]:
        if timeout <= 0:
            raise ValueError("timeout must be greater than zero")
        request_payload = dict(payload or {})
        deadline = time.monotonic() + timeout
        if "_deadline_monotonic" in request_payload:
            raw_payload_deadline = request_payload["_deadline_monotonic"]
            if isinstance(raw_payload_deadline, bool) or not isinstance(raw_payload_deadline, (int, float)):
                raise ControlProtocolError(
                    "invalid_control_deadline",
                    "GUI control deadline must be a finite monotonic timestamp.",
                )
            try:
                payload_deadline = float(raw_payload_deadline)
            except (OverflowError, ValueError) as exc:
                raise ControlProtocolError(
                    "invalid_control_deadline",
                    "GUI control deadline must be a finite monotonic timestamp.",
                ) from exc
            if not math.isfinite(payload_deadline):
                raise ControlProtocolError(
                    "invalid_control_deadline",
                    "GUI control deadline must be a finite monotonic timestamp.",
                )
            deadline = min(deadline, payload_deadline)
        _remaining_timeout(deadline)
        if self._endpoint_error is not None:
            try:
                self._address, self._family = control_endpoint(self._app_name)
            except ControlEndpointError as exc:
                raise exc from self._endpoint_error
            self._endpoint_error = None
        if sys.platform != "win32":
            plan = _plan_unix_endpoint(self._app_name, _user_namespace())
            try:
                _validate_unix_client_endpoint(plan)
            except FileNotFoundError as exc:
                raise ControlNotRunningError() from exc
            except OSError as exc:
                raise ControlEndpointError() from exc
            self._address = str(plan.address)
            self._family = "AF_UNIX"
        request_id = str(uuid4())
        connection = self._connect(deadline)
        try:
            encoded_request = _encode(
                {
                    "control_version": CONTROL_PROTOCOL_VERSION,
                    "request_id": request_id,
                    "action": action,
                    "payload": request_payload,
                }
            )
            connection.send_bytes(encoded_request, deadline=deadline)
            response_bytes = connection.recv_bytes(_MAX_MESSAGE_BYTES, deadline=deadline)
            message = _decode(response_bytes)
        except TimeoutError as exc:
            raise ControlTimeoutError() from exc
        except (EOFError, BrokenPipeError, OSError) as exc:
            raise ControlNotRunningError("DocWen GUI control connection closed unexpectedly.") from exc
        finally:
            connection.close()

        if message.get("control_version") != CONTROL_PROTOCOL_VERSION or message.get("request_id") != request_id:
            raise ControlProtocolError("invalid_control_response", "GUI control response identity is invalid.")
        if message.get("success") is True:
            data = message.get("data")
            if not isinstance(data, dict):
                raise ControlProtocolError("invalid_control_response", "GUI control data must be an object.")
            return data
        error = message.get("error")
        if not isinstance(error, dict):
            raise ControlProtocolError("invalid_control_response", "GUI control error is missing.")
        code = error.get("code")
        text = error.get("message")
        details = error.get("details")
        if not isinstance(code, str) or not isinstance(text, str) or not isinstance(details, dict):
            raise ControlProtocolError("invalid_control_response", "GUI control error fields are invalid.")
        raise ControlRemoteError(code, text, details=details)

    def _connect(self, deadline: float) -> LocalUnixStreamChannel | _WindowsPipeChannel:
        _remaining_timeout(deadline)
        if self._family == "AF_UNIX":
            unix_channel: LocalUnixStreamChannel | None = None
            try:
                unix_channel = LocalUnixStreamChannel.connect(self._address, deadline=deadline)
                unix_channel.authenticate_client(_PROTOCOL_KEY, deadline=deadline)
                return unix_channel
            except TimeoutError as exc:
                if unix_channel is not None:
                    unix_channel.close()
                raise ControlTimeoutError() from exc
            except (AuthenticationError, EOFError, BrokenPipeError, OSError) as exc:
                if unix_channel is not None:
                    unix_channel.close()
                raise ControlNotRunningError("DocWen GUI control closed during the connection handshake.") from exc
        if self._family == "AF_PIPE":
            if sys.platform != "win32":
                raise ControlNotRunningError()
            pipe_channel: _WindowsPipeChannel | None = None
            try:
                pipe_channel = _connect_windows_pipe(self._address, deadline=deadline)
                pipe_channel.authenticate_client(deadline=deadline)
                return pipe_channel
            except TimeoutError as exc:
                if pipe_channel is not None:
                    pipe_channel.close()
                raise ControlTimeoutError() from exc
            except (AuthenticationError, EOFError, BrokenPipeError) as exc:
                if pipe_channel is not None:
                    pipe_channel.close()
                raise ControlNotRunningError("DocWen GUI control closed during the connection handshake.") from exc
            except (FileNotFoundError, ConnectionRefusedError, OSError) as exc:
                if pipe_channel is not None:
                    pipe_channel.close()
                raise ControlNotRunningError() from exc
        raise ControlNotRunningError()


def _remaining_timeout(deadline: float) -> float:
    remaining = deadline - time.monotonic()
    if remaining <= 0:
        raise ControlTimeoutError()
    return remaining


def _connect_windows_pipe(address: str, *, deadline: float) -> _WindowsPipeChannel:
    if sys.platform != "win32":
        raise OSError("Windows control pipes are unavailable on this platform.")
    import _winapi
    from multiprocessing.connection import PipeConnection

    retry_errors = {_winapi.ERROR_PIPE_BUSY, _winapi.ERROR_SEM_TIMEOUT}
    while True:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise TimeoutError("Windows control pipe connection deadline expired.")
        wait_ms = max(1, min(1000, int(remaining * 1000 + 0.999)))
        try:
            _winapi.WaitNamedPipe(address, wait_ms)
        except OSError as exc:
            if getattr(exc, "winerror", None) in retry_errors:
                continue
            raise
        try:
            handle = _winapi.CreateFile(
                address,
                _winapi.GENERIC_READ | _winapi.GENERIC_WRITE,
                0,
                _winapi.NULL,
                _winapi.OPEN_EXISTING,
                _winapi.FILE_FLAG_OVERLAPPED,
                _winapi.NULL,
            )
        except OSError as exc:
            if getattr(exc, "winerror", None) == _winapi.ERROR_PIPE_BUSY:
                continue
            raise
        try:
            _winapi.SetNamedPipeHandleState(
                handle,
                _winapi.PIPE_READMODE_MESSAGE,
                None,
                None,
            )
            return _WindowsPipeChannel(PipeConnection(handle))
        except BaseException:
            _winapi.CloseHandle(handle)
            raise


class ControlServer:
    """Single-threaded local control server with a caller-provided handler."""

    def __init__(
        self,
        handler: Callable[[str, dict[str, Any]], Mapping[str, Any]],
        *,
        app_name: str = "docwen",
    ) -> None:
        self._handler = handler
        self._app_name = app_name
        self._address = ""
        self._family = "AF_PIPE" if sys.platform == "win32" else "AF_UNIX"
        self._stop = threading.Event()
        self._ready = threading.Event()
        self._startup_decision = threading.Event()
        self._serving = threading.Event()
        self._serve_acknowledged = threading.Event()
        self._lifecycle_lock = threading.RLock()
        self._listener: Listener | None = None
        self._thread: threading.Thread | None = None
        self._startup_error: BaseException | None = None
        self._shutdown_error: BaseException | None = None
        self._unix_lease: _UnixEndpointLease | None = None

    def start(self, *, timeout: float = 3.0) -> None:
        if timeout <= 0:
            raise ValueError("timeout must be greater than zero")
        existing_thread = self._thread
        if existing_thread is not None:
            if existing_thread.is_alive():
                if self._serving.is_set() and not self._stop.is_set():
                    return
                raise ControlTimeoutError("A previous GUI control startup is still settling.")
            self._thread = None
            self._cleanup_endpoint()
            if self._shutdown_error is not None:
                raise ControlError(
                    "gui_control_stop_failed",
                    "A previous GUI control cleanup did not complete safely.",
                ) from self._shutdown_error
        deadline = time.monotonic() + timeout
        self._stop.clear()
        self._ready.clear()
        self._startup_decision.clear()
        self._serving.clear()
        self._serve_acknowledged.clear()
        self._startup_error = None
        self._shutdown_error = None
        self._thread = threading.Thread(target=self._serve, name="docwen-runtime-control", daemon=True)
        self._thread.start()
        if not self._ready.wait(max(0.0, deadline - time.monotonic())):
            # A startup worker may still be inside endpoint planning or the
            # Listener constructor.  Mark it for shutdown before reporting the
            # timeout so that, once the blocking operation settles, it cannot
            # enter the serving loop and leave a live endpoint behind.
            self.begin_stop()
            raise ControlTimeoutError("Timed out while starting DocWen GUI control.")
        if self._startup_error is not None:
            self._startup_decision.set()
            raise ControlError(
                "gui_control_start_failed", "Unable to start DocWen GUI control."
            ) from self._startup_error
        with self._lifecycle_lock:
            if time.monotonic() >= deadline or self._stop.is_set():
                self._stop.set()
                self._serving.clear()
                timed_out = True
            else:
                # Commit serving before releasing the worker.  This keeps the
                # caller and worker on one side of the same lifecycle boundary:
                # start either returns with a committed server or releases a
                # worker that can only clean up its listener.
                self._serving.set()
                timed_out = False
            self._startup_decision.set()
        if timed_out:
            raise ControlTimeoutError("Timed out while starting DocWen GUI control.")
        if not self._serve_acknowledged.wait(max(0.0, deadline - time.monotonic())):
            with self._lifecycle_lock:
                if self._serve_acknowledged.is_set():
                    return
                self._stop.set()
                self._serving.clear()
                self._startup_decision.set()
            raise ControlTimeoutError("Timed out while starting DocWen GUI control.")

    def stop(self, *, timeout: float = 3.0) -> None:
        if timeout <= 0:
            raise ValueError("timeout must be greater than zero")
        deadline = time.monotonic() + timeout
        self.begin_stop()
        thread = self._thread
        if thread is None:
            self._cleanup_endpoint()
            return
        if thread is threading.current_thread():
            # The serving thread cannot join itself.  Keep its identity until
            # another caller observes that it has actually exited.
            return

        listener = self._listener
        wake_error: BaseException | None = None
        if listener is not None and thread.is_alive():
            # Wake the blocking accept before closing its handle.  Closing a
            # Windows PipeListener first can leave an already-blocked accept
            # alive until process shutdown.
            try:
                ControlClient(app_name=self._app_name).request(
                    "__stop__",
                    timeout=min(_remaining_timeout(deadline), 0.5),
                )
            except BaseException as exc:
                # Defer the decision until after join: a closed handshake is
                # harmless only when the serving thread really has exited.
                wake_error = exc

        remaining = deadline - time.monotonic()
        if remaining > 0:
            thread.join(remaining)
        if thread.is_alive():
            raise ControlTimeoutError("Timed out while stopping DocWen GUI control.") from wake_error

        # _serve owns listener shutdown and clears _listener in its finally
        # block.  A completed join therefore proves both resources settled.
        if self._listener is not None:
            raise ControlError(
                "gui_control_stop_failed",
                "DocWen GUI control listener did not close after its thread exited.",
            )
        self._thread = None
        self._cleanup_endpoint()
        shutdown_error = self._shutdown_error
        self._shutdown_error = None
        if shutdown_error is not None:
            raise ControlError(
                "gui_control_stop_failed",
                "DocWen GUI control cleanup did not complete safely.",
            ) from shutdown_error
        if wake_error is not None and not isinstance(
            wake_error,
            (ControlNotRunningError, ControlTimeoutError),
        ):
            raise wake_error

    def begin_stop(self) -> None:
        """Request shutdown without joining the control thread.

        GUI ``aboutToQuit`` handlers must not join a control handler that may
        still be waiting for the current Qt event to return.
        """

        with self._lifecycle_lock:
            self._stop.set()
            self._serving.clear()
            self._startup_decision.set()

    def _serve(self) -> None:
        listener: Listener | None = None
        try:
            namespace = _user_namespace()
            if sys.platform == "win32":
                self._address, self._family = control_endpoint(self._app_name)
            else:
                plan = _plan_unix_endpoint(self._app_name, namespace)
                self._unix_lease = _prepare_unix_server_endpoint(plan)
                self._address = str(plan.address)
                self._family = "AF_UNIX"
            listener = Listener(
                self._address,
                family=self._family,
                authkey=None,
            )
            self._listener = listener
            if self._unix_lease is not None:
                _secure_bound_unix_socket(self._unix_lease)
        except BaseException as exc:
            if listener is not None:
                with contextlib.suppress(Exception):
                    listener.close()
            self._listener = None
            cleanup_error: BaseException | None = None
            try:
                if self._unix_lease is not None:
                    _claim_new_unix_socket_for_cleanup(self._unix_lease, exc)
            except BaseException as claim_exc:
                cleanup_error = claim_exc
            try:
                self._cleanup_endpoint()
            except BaseException as release_exc:
                if cleanup_error is None:
                    cleanup_error = release_exc
            self._startup_error = cleanup_error if cleanup_error is not None else exc
            self._ready.set()
            return
        self._ready.set()
        try:
            self._startup_decision.wait()
            with self._lifecycle_lock:
                committed = self._serving.is_set() and not self._stop.is_set()
                if committed:
                    self._serve_acknowledged.set()
            if not committed:
                return
            while not self._stop.is_set():
                try:
                    connection = listener.accept()
                except (EOFError, OSError):
                    if self._stop.is_set():
                        break
                    continue
                unix_channel: LocalUnixStreamChannel | None = None
                try:
                    if self._family == "AF_UNIX":
                        unix_channel = LocalUnixStreamChannel.from_file_descriptor(connection.fileno())
                        unix_channel.authenticate_server(
                            _PROTOCOL_KEY,
                            deadline=time.monotonic() + _PEER_HANDSHAKE_TIMEOUT_SECONDS,
                        )
                        self._serve_one_unix(unix_channel)
                    else:
                        pipe_channel = _WindowsPipeChannel(connection)
                        pipe_channel.authenticate_server(
                            deadline=time.monotonic() + _PEER_HANDSHAKE_TIMEOUT_SECONDS,
                        )
                        self._serve_one_pipe(pipe_channel)
                except (AuthenticationError, EOFError, BrokenPipeError, TimeoutError, OSError):
                    pass
                finally:
                    if unix_channel is not None:
                        unix_channel.close()
                    connection.close()
        finally:
            with self._lifecycle_lock:
                self._serving.clear()
                self._serve_acknowledged.clear()
            with contextlib.suppress(Exception):
                listener.close()
            self._listener = None
            try:
                self._cleanup_endpoint()
            except BaseException as exc:
                self._shutdown_error = exc

    def _serve_one_unix(self, connection: LocalUnixStreamChannel) -> None:
        request = connection.recv_bytes(
            _MAX_MESSAGE_BYTES,
            deadline=time.monotonic() + _PEER_FRAME_TIMEOUT_SECONDS,
        )
        response = self._response_for_request(request)
        connection.send_bytes(
            response,
            deadline=time.monotonic() + _PEER_FRAME_TIMEOUT_SECONDS,
        )

    def _serve_one_pipe(self, connection: _WindowsPipeChannel) -> None:
        request = connection.recv_bytes(
            _MAX_MESSAGE_BYTES,
            deadline=time.monotonic() + _PEER_FRAME_TIMEOUT_SECONDS,
        )
        response = self._response_for_request(request)
        connection.send_bytes(
            response,
            deadline=time.monotonic() + _PEER_FRAME_TIMEOUT_SECONDS,
        )

    def _response_for_request(self, request: bytes) -> bytes:
        request_id = "unknown"
        try:
            message = _decode(request)
            raw_id = message.get("request_id")
            if isinstance(raw_id, str) and raw_id:
                request_id = raw_id
            request_id, action, payload = _validate_request(message)
            data = {} if action == "__stop__" else self._handler(action, payload)
            response = _response(request_id, success=True, data=data)
        except ControlError as exc:
            response = _response(request_id, success=False, error=exc)
        except Exception:
            response = _response(
                request_id,
                success=False,
                error=ControlRequestError("gui_command_failed", "DocWen GUI could not process the request."),
            )
        return _encode(response)

    def _cleanup_endpoint(self) -> None:
        lease = self._unix_lease
        if lease is None:
            return
        self._unix_lease = None
        _release_unix_server_endpoint(lease)
