"""Platform-independent single-instance lock using file locks.

Uses msvcrt.locking on Windows and fcntl.flock on Unix for non-blocking
exclusive file locks.  Windows retains the historical temporary-directory
layout.  Unix uses a user-private runtime directory and a persistent lock
sentinel so that unlink/recreate races cannot split ownership between two
processes.

The lock is released on process exit (best-effort via explicit ``release()``
or normal file-descriptor cleanup).  A clean Unix release clears the
diagnostic PID but deliberately retains the secure sentinel.

This module is test-friendly: when ``docwen_runtime.ipc.is_ipc_disabled()``
returns True, ``acquire()`` always succeeds (no real lock is taken).
"""

from __future__ import annotations

import contextlib
import errno
import hashlib
import logging
import os
import stat
import sys
import tempfile
from pathlib import Path
from typing import IO

# Platform-specific locking primitives
if sys.platform == "win32":
    import msvcrt  # type: ignore[import-untyped]

    def _acquire_file_lock(fd: int) -> None:
        """Non-blocking exclusive lock on Windows."""
        msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)

    def _release_file_lock(fd: int) -> None:
        """Release exclusive lock on Windows."""
        msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)

else:
    import fcntl

    def _acquire_file_lock(fd: int) -> None:
        """Non-blocking exclusive lock on Unix."""
        fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)

    def _release_file_lock(fd: int) -> None:
        """Release exclusive lock on Unix."""
        fcntl.flock(fd, fcntl.LOCK_UN)


logger = logging.getLogger(__name__)

# ── Process-level guard ──────────────────────────────────────────────────
# On some platforms (notably Windows), file locks are per-process, not
# per-file-descriptor.  This means a second ``SingleInstance`` in the
# *same* process could acquire the same lock again, defeating the purpose
# of single-instance detection.
#
# We track which lock paths have already been acquired in this process
# and refuse to re-acquire them.  This also makes tests more robust
# because two ``SingleInstance`` objects in the same pytest process
# cannot both claim the same lock.
_acquired_locks: set[str] = set()

_UNIX_DIRECTORY_MODE = 0o700
_UNIX_LOCK_MODE = 0o600
_UNIX_LOCK_NAME = "instance.lock"


class SingleInstanceError(Exception):
    """Raised when single-instance operations fail in a diagnosable way."""


def _current_uid() -> int:
    getuid = getattr(os, "getuid", None)
    if not callable(getuid):
        raise SingleInstanceError("unix_user_identity_unavailable")
    uid = getuid()
    if not isinstance(uid, int):
        raise SingleInstanceError("unix_user_identity_invalid")
    return uid


def _fchmod(descriptor: int, mode: int) -> None:
    fchmod = getattr(os, "fchmod", None)
    if not callable(fchmod):
        raise SingleInstanceError("unix_descriptor_chmod_unavailable")
    fchmod(descriptor, mode)


def _unix_directory_open_flags() -> int:
    directory = getattr(os, "O_DIRECTORY", None)
    nofollow = getattr(os, "O_NOFOLLOW", None)
    if not isinstance(directory, int) or not isinstance(nofollow, int):
        raise SingleInstanceError("unix_secure_directory_descriptors_unavailable")
    return os.O_RDONLY | directory | nofollow | int(getattr(os, "O_CLOEXEC", 0))


def _unix_file_open_flags() -> int:
    nofollow = getattr(os, "O_NOFOLLOW", None)
    if not isinstance(nofollow, int):
        raise SingleInstanceError("unix_secure_file_descriptors_unavailable")
    return os.O_RDWR | nofollow | int(getattr(os, "O_CLOEXEC", 0))


def _verified_unix_root(path: Path, *, require_user_private: bool) -> tuple[int, os.stat_result, str]:
    if not path.is_absolute():
        raise SingleInstanceError("unix_ipc_root_not_absolute")
    try:
        path_metadata = path.lstat()
    except OSError as exc:
        raise SingleInstanceError(f"unix_ipc_root_unavailable:{exc.errno}") from exc
    if stat.S_ISLNK(path_metadata.st_mode):
        raise SingleInstanceError("unix_ipc_root_link_forbidden")

    try:
        descriptor = os.open(path, _unix_directory_open_flags())
    except OSError as exc:
        raise SingleInstanceError(f"unix_ipc_root_open_failed:{exc.errno}") from exc
    try:
        metadata = os.fstat(descriptor)
        if not stat.S_ISDIR(metadata.st_mode):
            raise SingleInstanceError("unix_ipc_root_not_directory")
        if (path_metadata.st_dev, path_metadata.st_ino) != (metadata.st_dev, metadata.st_ino):
            raise SingleInstanceError("unix_ipc_root_identity_changed")

        mode = stat.S_IMODE(metadata.st_mode)
        uid = _current_uid()
        if metadata.st_uid == uid and mode == _UNIX_DIRECTORY_MODE:
            policy = "user-private"
        elif metadata.st_uid == 0 and bool(metadata.st_mode & stat.S_ISVTX) and mode & 0o003 == 0o003:
            policy = "root-owned-sticky"
        else:
            raise SingleInstanceError("unix_ipc_root_permissions_unsafe")
        if require_user_private and policy != "user-private":
            raise SingleInstanceError("unix_xdg_runtime_dir_not_private")
        return descriptor, metadata, policy
    except BaseException:
        os.close(descriptor)
        raise


def _fallback_unix_roots() -> tuple[Path, ...]:
    candidates = (Path(tempfile.gettempdir()), Path("/tmp"), Path("/var/tmp"))
    roots: list[Path] = []
    for candidate in candidates:
        try:
            resolved = candidate.resolve(strict=True)
        except OSError:
            continue
        if resolved not in roots:
            roots.append(resolved)
    return tuple(roots)


def _select_unix_root() -> tuple[Path, int]:
    xdg_runtime_dir = os.environ.get("XDG_RUNTIME_DIR", "").strip()
    if xdg_runtime_dir:
        root = Path(xdg_runtime_dir)
        descriptor, _metadata, _policy = _verified_unix_root(root, require_user_private=True)
        return root, descriptor

    last_error: SingleInstanceError | None = None
    for root in _fallback_unix_roots():
        try:
            descriptor, _metadata, _policy = _verified_unix_root(root, require_user_private=False)
        except SingleInstanceError as exc:
            last_error = exc
            continue
        return root, descriptor
    if last_error is not None:
        raise SingleInstanceError("unix_ipc_safe_root_unavailable") from last_error
    raise SingleInstanceError("unix_ipc_safe_root_unavailable")


def _unix_namespace_name(app_name: str) -> str:
    digest = hashlib.sha256(app_name.encode("utf-8", errors="strict")).hexdigest()[:16]
    return f"docwen-instance-{_current_uid()}-{digest}"


def _verify_unix_ipc_directory_metadata(
    metadata: os.stat_result,
    *,
    expected_device: int | None = None,
    expected_inode: int | None = None,
) -> None:
    if not stat.S_ISDIR(metadata.st_mode):
        raise SingleInstanceError("unix_ipc_directory_not_directory")
    if metadata.st_uid != _current_uid() or stat.S_IMODE(metadata.st_mode) != _UNIX_DIRECTORY_MODE:
        raise SingleInstanceError("unix_ipc_directory_permissions_unsafe")
    if (
        expected_device is not None
        and expected_inode is not None
        and (metadata.st_dev, metadata.st_ino) != (expected_device, expected_inode)
    ):
        raise SingleInstanceError("unix_ipc_directory_identity_changed")


def _prepare_unix_ipc_directory(app_name: str) -> Path:
    root, root_descriptor = _select_unix_root()
    directory_name = _unix_namespace_name(app_name)
    created = False
    try:
        try:
            os.mkdir(directory_name, _UNIX_DIRECTORY_MODE, dir_fd=root_descriptor)
            created = True
        except FileExistsError:
            pass
        except OSError as exc:
            raise SingleInstanceError(f"unix_ipc_directory_create_failed:{exc.errno}") from exc
        try:
            directory_descriptor = os.open(
                directory_name,
                _unix_directory_open_flags(),
                dir_fd=root_descriptor,
            )
        except OSError as exc:
            raise SingleInstanceError(f"unix_ipc_directory_open_failed:{exc.errno}") from exc
        try:
            if created:
                _fchmod(directory_descriptor, _UNIX_DIRECTORY_MODE)
            metadata = os.fstat(directory_descriptor)
            _verify_unix_ipc_directory_metadata(metadata)
            try:
                path_metadata = os.stat(directory_name, dir_fd=root_descriptor, follow_symlinks=False)
            except OSError as exc:
                raise SingleInstanceError(f"unix_ipc_directory_revalidation_failed:{exc.errno}") from exc
            _verify_unix_ipc_directory_metadata(
                path_metadata,
                expected_device=metadata.st_dev,
                expected_inode=metadata.st_ino,
            )
        finally:
            os.close(directory_descriptor)
    finally:
        os.close(root_descriptor)
    return root / directory_name


def _open_verified_unix_ipc_directory(path: Path) -> int:
    try:
        path_metadata = path.lstat()
    except OSError as exc:
        raise SingleInstanceError(f"unix_ipc_directory_unavailable:{exc.errno}") from exc
    if stat.S_ISLNK(path_metadata.st_mode):
        raise SingleInstanceError("unix_ipc_directory_link_forbidden")
    try:
        descriptor = os.open(path, _unix_directory_open_flags())
    except OSError as exc:
        raise SingleInstanceError(f"unix_ipc_directory_open_failed:{exc.errno}") from exc
    try:
        metadata = os.fstat(descriptor)
        _verify_unix_ipc_directory_metadata(
            metadata,
            expected_device=path_metadata.st_dev,
            expected_inode=path_metadata.st_ino,
        )
        return descriptor
    except BaseException:
        os.close(descriptor)
        raise


def _verify_unix_lock_metadata(metadata: os.stat_result, path_metadata: os.stat_result) -> None:
    if not stat.S_ISREG(metadata.st_mode):
        raise SingleInstanceError("unix_lock_not_regular")
    if metadata.st_uid != _current_uid() or stat.S_IMODE(metadata.st_mode) != _UNIX_LOCK_MODE:
        raise SingleInstanceError("unix_lock_permissions_unsafe")
    if metadata.st_nlink != 1:
        raise SingleInstanceError("unix_lock_link_count_unsafe")
    if (metadata.st_dev, metadata.st_ino) != (path_metadata.st_dev, path_metadata.st_ino):
        raise SingleInstanceError("unix_lock_identity_changed")


def _open_unix_lock_file(ipc_dir: Path) -> IO[str]:
    directory_descriptor = _open_verified_unix_ipc_directory(ipc_dir)
    file_descriptor = -1
    created = False
    try:
        flags = _unix_file_open_flags()
        try:
            file_descriptor = os.open(
                _UNIX_LOCK_NAME,
                flags | os.O_CREAT | os.O_EXCL,
                _UNIX_LOCK_MODE,
                dir_fd=directory_descriptor,
            )
            created = True
        except FileExistsError:
            try:
                file_descriptor = os.open(_UNIX_LOCK_NAME, flags, dir_fd=directory_descriptor)
            except OSError as exc:
                raise SingleInstanceError(f"unix_lock_open_failed:{exc.errno}") from exc
        if created:
            _fchmod(file_descriptor, _UNIX_LOCK_MODE)
        metadata = os.fstat(file_descriptor)
        try:
            path_metadata = os.stat(_UNIX_LOCK_NAME, dir_fd=directory_descriptor, follow_symlinks=False)
        except OSError as exc:
            raise SingleInstanceError(f"unix_lock_revalidation_failed:{exc.errno}") from exc
        _verify_unix_lock_metadata(metadata, path_metadata)
        os.set_inheritable(file_descriptor, False)
        lock_file = os.fdopen(file_descriptor, "r+", encoding="utf-8", closefd=True)
        file_descriptor = -1
        return lock_file
    except SingleInstanceError:
        raise
    except OSError as exc:
        raise SingleInstanceError(f"unix_lock_open_failed:{exc.errno}") from exc
    finally:
        if file_descriptor >= 0:
            os.close(file_descriptor)
        os.close(directory_descriptor)


class SingleInstance:
    """File-lock based single-instance manager.

    Ensures only one instance of the application runs at a time by
    acquiring an exclusive, non-blocking file lock.

    Usage::

        si = SingleInstance(app_name="docwen")
        if si.acquire():
            # This is the first instance — run normally.
            ...
            si.release()
        else:
            # Another instance is already running.
            print("Application is already running.")
            si.close()

    The instance can also be used as a context manager::

        with SingleInstance("docwen") as acquired:
            if acquired:
                run_app()
    """

    # ── Construction ──────────────────────────────────────────────────────

    def __init__(self, app_name: str = "docwen") -> None:
        """Initialise the single-instance manager.

        Args:
            app_name: Application name used to derive the lock identity.  On
                Windows the lock file remains at
                ``{tempdir}/{app_name}/instance.lock``.  Unix hashes the name
                into a user-private runtime namespace.
        """
        if not app_name or not app_name.strip():
            raise ValueError("app_name must be a non-empty string")

        self._app_name: str = app_name.strip()
        self._lock_file: object | None = None
        self._ipc_dir: str = self._compute_ipc_dir()
        self._lock_path: str = str(Path(self._ipc_dir) / "instance.lock")
        self._acquired: bool = False

        if sys.platform == "win32":
            # Preserve the established Windows path and creation behavior.
            try:
                Path(self._ipc_dir).mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                raise SingleInstanceError(f"无法创建 IPC 目录 {self._ipc_dir}: {exc}") from exc

        logger.debug("IPC directory: %s", self._ipc_dir)

    def _compute_ipc_dir(self) -> str:
        """Return the IPC directory path.

        Windows uses the historical system-temp path.  Unix prefers a valid
        private ``XDG_RUNTIME_DIR`` and otherwise creates a per-user namespace
        below a verified temporary root.
        """
        if sys.platform != "win32":
            return str(_prepare_unix_ipc_directory(self._app_name))
        temp_dir = tempfile.gettempdir()
        return str(Path(temp_dir) / self._app_name)

    # ── Public API ────────────────────────────────────────────────────────

    def acquire(self) -> bool:
        """Attempt to acquire the single-instance lock.

        Returns:
            True if the lock was acquired (this is the first instance).
            False if another instance holds the lock.

        Raises:
            SingleInstanceError: On unexpected I/O errors (not on lock
                contention — those return False).
        """
        from . import is_ipc_disabled

        # Test-mode / debug bypass: always succeed.
        if is_ipc_disabled():
            logger.info("IPC disabled — single-instance lock bypassed")
            self._acquired = True
            return True

        # Process-level guard: prevent re-acquire in the same process.
        # On Windows file locks are per-process, so a second open()+lock
        # on the same lock file would succeed even though another instance
        # of SingleInstance in the same process already holds it.
        if self._lock_path in _acquired_locks:
            logger.info(
                "Another instance in this process already holds the lock: %s",
                self._lock_path,
            )
            self._acquired = False
            return False

        if sys.platform == "win32":
            return self._acquire_windows()
        return self._acquire_unix()

    def _acquire_windows(self) -> bool:
        """Acquire using the established Windows behavior unchanged."""
        try:
            lock_file = Path(self._lock_path).open("w", encoding="utf-8")  # noqa: SIM115 — lock must stay open
        except OSError as exc:
            raise SingleInstanceError(f"无法打开锁文件 {self._lock_path}: {exc}") from exc

        try:
            _acquire_file_lock(lock_file.fileno())
        except OSError:
            # Lock contention — another instance is running.
            with contextlib.suppress(Exception):
                lock_file.close()
            logger.info("Another instance is already running (lock held: %s)", self._lock_path)
            self._lock_file = None
            self._acquired = False
            return False

        # Lock acquired — write PID for diagnostic visibility.
        try:
            lock_file.write(str(os.getpid()))
            lock_file.flush()
        except OSError as exc:
            # Write failed but lock is held — release and report.
            with contextlib.suppress(Exception):
                _release_file_lock(lock_file.fileno())
            with contextlib.suppress(Exception):
                lock_file.close()
            raise SingleInstanceError(f"无法写入 PID 到锁文件 {self._lock_path}: {exc}") from exc

        self._lock_file = lock_file
        self._acquired = True
        _acquired_locks.add(self._lock_path)
        logger.info("Single-instance lock acquired: %s (PID %d)", self._lock_path, os.getpid())
        return True

    def _acquire_unix(self) -> bool:
        """Acquire a no-follow regular-file lock in the private Unix namespace."""
        lock_file = _open_unix_lock_file(Path(self._ipc_dir))
        try:
            _acquire_file_lock(lock_file.fileno())
        except OSError as exc:
            with contextlib.suppress(Exception):
                lock_file.close()
            if isinstance(exc, BlockingIOError) or exc.errno in {errno.EACCES, errno.EAGAIN}:
                logger.info("Another instance is already running (lock held: %s)", self._lock_path)
                self._lock_file = None
                self._acquired = False
                return False
            raise SingleInstanceError(f"unix_lock_acquire_failed:{exc.errno}") from exc

        try:
            lock_file.seek(0)
            lock_file.truncate(0)
            lock_file.write(str(os.getpid()))
            lock_file.flush()
            os.fsync(lock_file.fileno())
        except OSError as exc:
            with contextlib.suppress(Exception):
                _release_file_lock(lock_file.fileno())
            with contextlib.suppress(Exception):
                lock_file.close()
            raise SingleInstanceError(f"unix_lock_pid_write_failed:{exc.errno}") from exc

        self._lock_file = lock_file
        self._acquired = True
        _acquired_locks.add(self._lock_path)
        logger.info("Single-instance lock acquired: %s (PID %d)", self._lock_path, os.getpid())
        return True

    def release(self) -> None:
        """Release the single-instance lock; safe to call multiple times.

        Windows retains its historical best-effort lock-file deletion.
        Unix clears the diagnostic PID but keeps the sentinel inode so a
        waiter can never race an unlink and acquire a different lock file.
        Cleanup errors are logged rather than re-raised during shutdown.
        """
        from . import is_ipc_disabled

        if not self._acquired:
            return

        # Remove from process-level guard set.
        _acquired_locks.discard(self._lock_path)

        if is_ipc_disabled():
            self._acquired = False
            self._lock_file = None
            return

        lock_file = self._lock_file
        self._lock_file = None
        self._acquired = False

        if lock_file is None:
            return

        if sys.platform != "win32":
            self._release_unix(lock_file)
            return

        fd = getattr(lock_file, "fileno", lambda: -1)()

        # Release the OS-level lock.
        if fd >= 0:
            try:
                _release_file_lock(fd)
            except Exception as exc:
                err_no = getattr(exc, "errno", None)
                if not (isinstance(exc, PermissionError) or err_no == 13):
                    logger.debug("Unlock error (ignorable): %s", exc)

        # Close the file handle.
        with contextlib.suppress(Exception):
            lock_file.close()  # pyright: ignore[reportAttributeAccessIssue]

        # Best-effort delete the lock file.
        try:
            lock_path = Path(self._lock_path)
            if lock_path.exists():
                lock_path.unlink()
                logger.debug("Lock file deleted: %s", self._lock_path)
        except Exception as exc:
            logger.debug("Could not delete lock file (OS will reclaim): %s", exc)

        logger.info("Single-instance lock released: %s", self._lock_path)

    def _release_unix(self, lock_file: object) -> None:
        """Release Unix ownership without unlinking the shared sentinel inode."""
        fd = getattr(lock_file, "fileno", lambda: -1)()
        if fd >= 0:
            try:
                lock_file.seek(0)  # pyright: ignore[reportAttributeAccessIssue]
                lock_file.truncate(0)  # pyright: ignore[reportAttributeAccessIssue]
                lock_file.flush()  # pyright: ignore[reportAttributeAccessIssue]
            except Exception as exc:
                logger.debug("Could not clear Unix lock diagnostic PID: %s", exc)
            try:
                _release_file_lock(fd)
            except Exception as exc:
                logger.debug("Unix unlock error (ignorable during shutdown): %s", exc)
        with contextlib.suppress(Exception):
            lock_file.close()  # pyright: ignore[reportAttributeAccessIssue]
        logger.info("Single-instance lock released: %s", self._lock_path)

    def close(self) -> None:
        """Alias for release() — clean up without acquiring."""
        self.release()

    # ── Properties ────────────────────────────────────────────────────────

    @property
    def ipc_dir(self) -> str:
        """The IPC directory path used by this instance."""
        return self._ipc_dir

    @property
    def lock_path(self) -> str:
        """Full path to the lock file."""
        return self._lock_path

    @property
    def is_acquired(self) -> bool:
        """Whether the lock is currently held by this instance."""
        return self._acquired

    # ── Context manager ───────────────────────────────────────────────────

    def __enter__(self) -> bool:
        """Context manager: acquire the lock and return success status."""
        return self.acquire()

    def __exit__(self, *args: object) -> None:
        """Context manager: release the lock."""
        self.release()


def create_single_instance(app_name: str = "docwen") -> SingleInstance:
    """Convenience factory for creating a SingleInstance.

    Args:
        app_name: Application name for lock file derivation.

    Returns:
        A configured (but not yet acquired) SingleInstance.
    """
    return SingleInstance(app_name=app_name)


__all__ = [
    "SingleInstance",
    "SingleInstanceError",
    "create_single_instance",
]
