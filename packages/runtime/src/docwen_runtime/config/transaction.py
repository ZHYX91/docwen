"""Durable before-image journal and cooperative cross-process config lock."""

from __future__ import annotations

import base64
import contextlib
import hashlib
import json
import logging
import os
import stat
import threading
from collections.abc import Callable, Iterable
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Any

from docwen_runtime import toml_io

logger = logging.getLogger(__name__)

CONFIG_JOURNAL_NAME = ".docwen-config.transaction.json"
CONFIG_LOCK_NAME = ".docwen-config.lock"
JOURNAL_VERSION = 1
_PREPARED = "PREPARED"
_COMMITTED = "COMMITTED"
_PROCESS_LOCK_STATE = threading.local()


class ConfigJournalError(RuntimeError):
    """A journal cannot be trusted or recovered safely."""


@dataclass(frozen=True)
class UserFilePreimage:
    """Logical user path plus recoverable file identity and metadata."""

    path: Path
    content: bytes | None
    symlink_target: Path | None = None
    resolved_target: Path | None = None
    mode: int | None = None
    atime_ns: int | None = None
    mtime_ns: int | None = None


@dataclass(frozen=True)
class JournalRecord:
    """Validated journal payload."""

    operation: str
    state: str
    preimages: tuple[UserFilePreimage, ...]


def _lexical_absolute(path: Path) -> Path:
    return Path(os.path.abspath(path))


def _path_exists_or_is_symlink(path: Path) -> bool:
    return path.exists() or path.is_symlink()


def _capture_target(path: Path) -> tuple[bytes | None, int | None, int | None, int | None]:
    if not path.exists():
        return None, None, None, None
    file_stat = path.stat()
    return (
        path.read_bytes(),
        stat.S_IMODE(file_stat.st_mode),
        file_stat.st_atime_ns,
        file_stat.st_mtime_ns,
    )


def capture_user_file_preimage(path: Path) -> UserFilePreimage:
    """Capture bytes, link identity, and portable regular-file metadata."""
    if path.is_symlink():
        resolved_target = path.resolve(strict=False)
        content, mode, atime_ns, mtime_ns = _capture_target(resolved_target)
        return UserFilePreimage(
            path=path,
            content=content,
            symlink_target=path.readlink(),
            resolved_target=resolved_target,
            mode=mode,
            atime_ns=atime_ns,
            mtime_ns=mtime_ns,
        )
    content, mode, atime_ns, mtime_ns = _capture_target(path)
    return UserFilePreimage(
        path=path,
        content=content,
        mode=mode,
        atime_ns=atime_ns,
        mtime_ns=mtime_ns,
    )


def _restore_metadata(path: Path, preimage: UserFilePreimage) -> None:
    current_mode = stat.S_IMODE(path.stat().st_mode)
    if not current_mode & stat.S_IWUSR:
        os.chmod(path, current_mode | stat.S_IWUSR)
    descriptor = os.open(path, os.O_RDWR)
    try:
        if preimage.atime_ns is not None and preimage.mtime_ns is not None:
            os.utime(path, ns=(preimage.atime_ns, preimage.mtime_ns))
        if preimage.mode is not None:
            os.chmod(path, preimage.mode)
        os.fsync(descriptor)
    finally:
        os.close(descriptor)
    toml_io._sync_directory(path.parent)


def restore_user_file_preimage(
    preimage: UserFilePreimage,
    *,
    operation: str,
    atomic_write: Callable[[Path, bytes], None] = toml_io.atomic_write_bytes,
) -> None:
    """Durably restore one before-image; safe to repeat after process death."""

    def restore_regular(path: Path) -> None:
        if preimage.content is None:
            if _path_exists_or_is_symlink(path):
                toml_io.durable_unlink(path)
            return
        try:
            requires_restore = not path.exists() or path.read_bytes() != preimage.content
        except Exception as comparison_exc:
            logger.warning(
                "Configuration rollback comparison failed; forcing restore: operation=%s path=%s error=%s",
                operation,
                path,
                comparison_exc,
            )
            requires_restore = True
        if path.is_symlink():
            toml_io.durable_unlink(path)
            requires_restore = True
        if requires_restore:
            atomic_write(path, preimage.content)
        _restore_metadata(path, preimage)

    if preimage.symlink_target is None:
        restore_regular(preimage.path)
        return

    resolved_target = preimage.resolved_target
    if resolved_target is None:
        raise RuntimeError(f"missing resolved target for symlink preimage: {preimage.path}")
    restore_regular(resolved_target)
    current_target = preimage.path.readlink() if preimage.path.is_symlink() else None
    if current_target == preimage.symlink_target:
        return
    if _path_exists_or_is_symlink(preimage.path):
        toml_io.durable_unlink(preimage.path)
    preimage.path.parent.mkdir(parents=True, exist_ok=True)
    preimage.path.symlink_to(preimage.symlink_target, target_is_directory=False)
    toml_io._sync_directory(preimage.path.parent)


class _ProcessFileLock:
    def __init__(self, path: Path) -> None:
        self._path = path
        self._descriptor: int | None = None

    def __enter__(self) -> _ProcessFileLock:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        descriptor = os.open(self._path, os.O_RDWR | os.O_CREAT, 0o600)
        self._descriptor = descriptor
        if os.fstat(descriptor).st_size == 0:
            os.write(descriptor, b"\0")
            os.fsync(descriptor)
        os.lseek(descriptor, 0, os.SEEK_SET)
        if os.name == "nt":
            import msvcrt  # type: ignore[import-untyped]

            msvcrt.locking(descriptor, msvcrt.LK_LOCK, 1)
        else:
            import fcntl

            fcntl.flock(descriptor, fcntl.LOCK_EX)
        return self

    def __exit__(self, _exc_type: Any, _exc: Any, _traceback: Any) -> None:
        descriptor = self._descriptor
        self._descriptor = None
        if descriptor is None:
            return
        try:
            os.lseek(descriptor, 0, os.SEEK_SET)
            if os.name == "nt":
                import msvcrt  # type: ignore[import-untyped]

                msvcrt.locking(descriptor, msvcrt.LK_UNLCK, 1)
            else:
                import fcntl

                fcntl.flock(descriptor, fcntl.LOCK_UN)
        finally:
            os.close(descriptor)


@contextmanager
def process_config_lock(user_dir: Path):
    """Acquire the user config root's re-entrant process lock.

    The lock is a sibling of an optional ``configs`` directory, so a read-only
    startup does not materialize an otherwise absent override directory.
    """
    lock_path = _lexical_absolute(user_dir.parent / f".{user_dir.name}{CONFIG_LOCK_NAME}")
    depth = int(getattr(_PROCESS_LOCK_STATE, "depth", 0))
    active_path = getattr(_PROCESS_LOCK_STATE, "path", None)
    if depth:
        if active_path != lock_path:
            raise RuntimeError("nested configuration lock targets differ")
        _PROCESS_LOCK_STATE.depth = depth + 1
        try:
            yield
        finally:
            _PROCESS_LOCK_STATE.depth -= 1
        return

    with _ProcessFileLock(lock_path):
        _PROCESS_LOCK_STATE.path = lock_path
        _PROCESS_LOCK_STATE.depth = 1
        try:
            yield
        finally:
            del _PROCESS_LOCK_STATE.depth
            del _PROCESS_LOCK_STATE.path


def _canonical_json(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def _relative_logical_path(user_dir: Path, path: Path) -> str:
    root = _lexical_absolute(user_dir)
    logical = _lexical_absolute(path)
    try:
        return logical.relative_to(root).as_posix()
    except ValueError as exc:
        raise ConfigJournalError(f"configuration transaction journal path escapes user directory: {path}") from exc


def _serialize_preimage(user_dir: Path, preimage: UserFilePreimage) -> dict[str, Any]:
    return {
        "path": _relative_logical_path(user_dir, preimage.path),
        "content": (base64.b64encode(preimage.content).decode("ascii") if preimage.content is not None else None),
        "symlink_target": str(preimage.symlink_target) if preimage.symlink_target is not None else None,
        "mode": preimage.mode,
        "atime_ns": preimage.atime_ns,
        "mtime_ns": preimage.mtime_ns,
    }


def _journal_payload(
    user_dir: Path,
    operation: str,
    preimages: Iterable[UserFilePreimage],
    state: str,
) -> dict[str, Any]:
    return {
        "version": JOURNAL_VERSION,
        "operation": operation,
        "state": state,
        "preimages": [_serialize_preimage(user_dir, item) for item in preimages],
    }


def write_transaction_journal(
    user_dir: Path,
    operation: str,
    preimages: Iterable[UserFilePreimage],
    *,
    state: str,
) -> None:
    """Durably publish a PREPARED or COMMITTED checksummed journal."""
    if state not in {_PREPARED, _COMMITTED}:
        raise ValueError(f"unsupported configuration journal state: {state}")
    payload = _journal_payload(user_dir, operation, preimages, state)
    envelope = {
        "checksum": hashlib.sha256(_canonical_json(payload)).hexdigest(),
        "payload": payload,
    }
    toml_io.atomic_write_bytes(user_dir / CONFIG_JOURNAL_NAME, _canonical_json(envelope) + b"\n")


def mark_transaction_committed(
    user_dir: Path,
    operation: str,
    preimages: Iterable[UserFilePreimage],
) -> None:
    """Publish the durable commit linearization point."""
    write_transaction_journal(user_dir, operation, preimages, state=_COMMITTED)


def remove_transaction_journal(user_dir: Path) -> None:
    """Durably remove the authoritative journal, if present."""
    toml_io.durable_unlink(user_dir / CONFIG_JOURNAL_NAME, missing_ok=True)


def _decode_optional_int(value: Any, field: str) -> int | None:
    if value is None:
        return None
    if not isinstance(value, int):
        raise ConfigJournalError(f"configuration transaction journal has invalid {field}")
    return value


def _decode_preimage(
    user_dir: Path,
    raw: Any,
    allowed_paths: set[Path],
) -> UserFilePreimage:
    if not isinstance(raw, dict):
        raise ConfigJournalError("configuration transaction journal preimage is not an object")
    raw_path = raw.get("path")
    if not isinstance(raw_path, str):
        raise ConfigJournalError("configuration transaction journal path is invalid")
    relative = PurePosixPath(raw_path)
    if relative.is_absolute() or not relative.parts or ".." in relative.parts:
        raise ConfigJournalError("configuration transaction journal path escapes user directory")
    logical = _lexical_absolute(user_dir.joinpath(*relative.parts))
    if logical not in allowed_paths:
        raise ConfigJournalError(f"configuration transaction journal path is not registered: {raw_path}")
    encoded_content = raw.get("content")
    if encoded_content is None:
        content = None
    elif isinstance(encoded_content, str):
        try:
            content = base64.b64decode(encoded_content, validate=True)
        except Exception as exc:
            raise ConfigJournalError("configuration transaction journal content is invalid") from exc
    else:
        raise ConfigJournalError("configuration transaction journal content is invalid")
    raw_target = raw.get("symlink_target")
    if raw_target is not None and not isinstance(raw_target, str):
        raise ConfigJournalError("configuration transaction journal symlink target is invalid")
    symlink_target = Path(raw_target) if raw_target is not None else None
    resolved_target: Path | None = None
    if symlink_target is not None:
        target = symlink_target if symlink_target.is_absolute() else logical.parent / symlink_target
        resolved_target = target.resolve(strict=False)
    return UserFilePreimage(
        path=logical,
        content=content,
        symlink_target=symlink_target,
        resolved_target=resolved_target,
        mode=_decode_optional_int(raw.get("mode"), "mode"),
        atime_ns=_decode_optional_int(raw.get("atime_ns"), "atime_ns"),
        mtime_ns=_decode_optional_int(raw.get("mtime_ns"), "mtime_ns"),
    )


def read_transaction_journal(user_dir: Path, allowed_paths: Iterable[Path]) -> JournalRecord | None:
    """Read and strictly validate the authoritative journal."""
    journal_path = user_dir / CONFIG_JOURNAL_NAME
    if not journal_path.exists():
        return None
    try:
        envelope = json.loads(journal_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise ConfigJournalError("configuration transaction journal is unreadable") from exc
    if not isinstance(envelope, dict) or set(envelope) != {"checksum", "payload"}:
        raise ConfigJournalError("configuration transaction journal envelope is invalid")
    payload = envelope.get("payload")
    checksum = envelope.get("checksum")
    if not isinstance(payload, dict) or not isinstance(checksum, str):
        raise ConfigJournalError("configuration transaction journal envelope is invalid")
    if hashlib.sha256(_canonical_json(payload)).hexdigest() != checksum:
        raise ConfigJournalError("configuration transaction journal checksum mismatch")
    if payload.get("version") != JOURNAL_VERSION:
        raise ConfigJournalError("configuration transaction journal version is unsupported")
    operation = payload.get("operation")
    state = payload.get("state")
    raw_preimages = payload.get("preimages")
    if not isinstance(operation, str) or not operation:
        raise ConfigJournalError("configuration transaction journal operation is invalid")
    if state not in {_PREPARED, _COMMITTED}:
        raise ConfigJournalError("configuration transaction journal state is invalid")
    if not isinstance(raw_preimages, list):
        raise ConfigJournalError("configuration transaction journal preimages are invalid")
    allowed = {_lexical_absolute(path) for path in allowed_paths}
    preimages = tuple(_decode_preimage(user_dir, raw, allowed) for raw in raw_preimages)
    if len({item.path for item in preimages}) != len(preimages):
        raise ConfigJournalError("configuration transaction journal contains duplicate paths")
    return JournalRecord(operation=operation, state=state, preimages=preimages)


def _cleanup_journal_temps(user_dir: Path) -> None:
    pattern = f".{CONFIG_JOURNAL_NAME}.*.tmp"
    for temp_path in user_dir.glob(pattern):
        with contextlib.suppress(OSError):
            toml_io.durable_unlink(temp_path, missing_ok=True)


def recover_transaction_journal(
    user_dir: Path,
    allowed_paths: Iterable[Path],
    restore: Callable[[UserFilePreimage, str], None],
) -> str | None:
    """Recover PREPARED state or retain COMMITTED state under the process lock."""
    record = read_transaction_journal(user_dir, allowed_paths)
    if record is None:
        _cleanup_journal_temps(user_dir)
        return None
    if record.state == _COMMITTED:
        try:
            remove_transaction_journal(user_dir)
            _cleanup_journal_temps(user_dir)
        except OSError as exc:
            logger.warning(
                "Committed configuration journal cleanup deferred: operation=%s error=%s",
                record.operation,
                exc,
            )
        return _COMMITTED

    for preimage in reversed(record.preimages):
        restore(preimage, f"recovery:{record.operation}")
    remove_transaction_journal(user_dir)
    _cleanup_journal_temps(user_dir)
    return _PREPARED
