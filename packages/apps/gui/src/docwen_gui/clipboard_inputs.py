"""Managed, session-owned materialization for user-triggered clipboard Markdown."""

from __future__ import annotations

import hashlib
import os
import shutil
import unicodedata
import uuid
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path

_PREVIEW_MAX_CHARS = 240
_PREVIEW_MAX_LINES = 3
_SESSION_PREFIX = "session-"
_OWNER_LOCK_NAME = ".owner.lock"
_OWNER_MARKER = b"docwen-clipboard-session-v1\n"


@contextmanager
def _session_namespace_lock(root: Path):
    """Serialize session publication and retirement across all Store instances."""
    with (root / ".sessions.lock").open("a+b") as stream:
        if os.fstat(stream.fileno()).st_size == 0:
            stream.write(b"\0")
            stream.flush()
        _lock_owner_stream(stream, blocking=True)
        try:
            yield
        finally:
            _unlock_owner_stream(stream)


def _lock_owner_stream(stream, *, blocking: bool) -> None:
    stream.seek(0)
    if os.name == "nt":
        import msvcrt  # type: ignore[import-untyped]

        mode = msvcrt.LK_LOCK if blocking else msvcrt.LK_NBLCK
        msvcrt.locking(stream.fileno(), mode, 1)
    else:
        import fcntl

        flags = fcntl.LOCK_EX | (0 if blocking else fcntl.LOCK_NB)
        fcntl.flock(stream.fileno(), flags)


def _unlock_owner_stream(stream) -> None:
    stream.seek(0)
    if os.name == "nt":
        import msvcrt  # type: ignore[import-untyped]

        msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
    else:
        import fcntl

        fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


def _retire_stale_session(session: Path) -> None:
    """Keep recovery authority until all authored bytes have been removed.

    The caller holds the root namespace lock and has proved the owner idle.
    A partial content deletion must leave the marker for the next startup.
    """
    for child in session.iterdir():
        if child.name == _OWNER_LOCK_NAME:
            continue
        if child.is_dir() and not child.is_symlink():
            shutil.rmtree(child)
        else:
            child.unlink()
    (session / _OWNER_LOCK_NAME).unlink()
    session.rmdir()


@dataclass(frozen=True, slots=True)
class ClipboardInputDescriptor:
    path: str
    display_name: str
    preview: str
    size_bytes: int
    sha256: str


def bounded_plaintext_preview(text: str) -> str:
    """Return a short display-only preview without changing the stored bytes."""

    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    safe = "".join(
        char if char in {"\n", "\t"} or not unicodedata.category(char).startswith("C") else "\ufffd"
        for char in normalized
    )
    lines = safe.strip().splitlines()
    preview = "\n".join(lines[:_PREVIEW_MAX_LINES])
    truncated = len(lines) > _PREVIEW_MAX_LINES or len(preview) > _PREVIEW_MAX_CHARS
    preview = preview[:_PREVIEW_MAX_CHARS]
    return preview + ("…" if truncated else "")


class ClipboardInputStore:
    """Own clipboard snapshots until no UI, worker, or retry record references them."""

    def __init__(self, root_dir: str | os.PathLike[str]) -> None:
        root = Path(root_dir).expanduser()
        if not root.is_absolute():
            raise ValueError("clipboard input root must be absolute")
        root = Path(os.path.normpath(os.fspath(root)))
        if root == Path(root.anchor):
            raise ValueError("clipboard input root must not be a filesystem root")
        root.mkdir(parents=True, exist_ok=True)
        self._root = root
        self._session_root = root / f"{_SESSION_PREFIX}{uuid.uuid4().hex}"
        self._owner_lock = None
        self._cleanup_pending = False
        with _session_namespace_lock(root):
            self._cleanup_stale_sessions(root)
            self._session_root.mkdir()
            owner_lock = None
            try:
                lock_path = self._session_root / _OWNER_LOCK_NAME
                owner_lock = lock_path.open("x+b")
                owner_lock.write(b"\0")
                owner_lock.flush()
                _lock_owner_stream(owner_lock, blocking=True)
                # A valid recovery marker is only published with ownership held.
                owner_lock.seek(0)
                owner_lock.write(_OWNER_MARKER)
                owner_lock.flush()
                os.fsync(owner_lock.fileno())
                self._owner_lock = owner_lock
            except BaseException:
                if owner_lock is not None:
                    owner_lock.close()
                shutil.rmtree(self._session_root, ignore_errors=True)
                raise
        self._snapshots: dict[str, ClipboardInputDescriptor] = {}
        self._visible: set[str] = set()
        self._inspection: dict[str, set[str]] = {}
        self._active: dict[str, set[str]] = {}
        self._history: dict[str, set[str]] = {}
        self._sequence = 0

    @staticmethod
    def _key(path: str | os.PathLike[str]) -> str:
        return os.path.normcase(os.path.abspath(os.fspath(path)))

    @classmethod
    def _cleanup_stale_sessions(cls, root: Path) -> None:
        """Remove idle sessions while the caller holds the namespace lock.

        The namespace lock remains held through deletion, including on Windows
        where the per-session file must be closed before removing its directory.
        """

        try:
            candidates = tuple(root.iterdir())
        except OSError:
            return
        for session in candidates:
            if not session.name.startswith(_SESSION_PREFIX):
                continue
            try:
                if session.is_symlink() or not session.is_dir():
                    continue
                lock_path = session / _OWNER_LOCK_NAME
                if not lock_path.is_file():
                    continue
                stream = lock_path.open("r+b")
                try:
                    marker = stream.read(len(_OWNER_MARKER))
                    if marker != _OWNER_MARKER:
                        continue
                    try:
                        _lock_owner_stream(stream, blocking=False)
                    except OSError:
                        continue
                    else:
                        _unlock_owner_stream(stream)
                finally:
                    stream.close()
                _retire_stale_session(session)
            except OSError:
                continue

    def _release_owner_lock(self) -> None:
        stream = self._owner_lock
        self._owner_lock = None
        if stream is None:
            return
        try:
            _unlock_owner_stream(stream)
        finally:
            stream.close()

    @property
    def cleanup_pending(self) -> bool:
        return self._cleanup_pending

    @property
    def session_root(self) -> Path:
        return self._session_root

    def create(self, text: str, *, display_name_template: str) -> ClipboardInputDescriptor:
        """Freeze the current plain text as exact UTF-8 bytes in this owned session."""

        if not isinstance(text, str):
            raise TypeError("clipboard Markdown must be plain text")
        if not text.strip():
            raise ValueError("clipboard Markdown is empty")
        self._sequence += 1
        display_name = display_name_template.format(index=self._sequence).strip()
        if not display_name:
            display_name = f"Clipboard Markdown {self._sequence}.md"
        path = self._session_root / f"clipboard-{uuid.uuid4().hex}.md"
        payload = text.encode("utf-8")
        descriptor = ClipboardInputDescriptor(
            path=str(path),
            display_name=display_name,
            preview=bounded_plaintext_preview(text),
            size_bytes=len(payload),
            sha256=hashlib.sha256(payload).hexdigest(),
        )
        try:
            with path.open("xb") as stream:
                stream.write(payload)
                stream.flush()
                os.fsync(stream.fileno())
        except BaseException:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                self._snapshots[self._key(path)] = descriptor
            raise
        self._snapshots[self._key(path)] = descriptor
        return descriptor

    def descriptor(self, path: str | os.PathLike[str]) -> ClipboardInputDescriptor | None:
        return self._snapshots.get(self._key(path))

    def is_snapshot(self, path: str | os.PathLike[str]) -> bool:
        return self._key(path) in self._snapshots

    def snapshot_available(self, path: str | os.PathLike[str]) -> bool:
        descriptor = self.descriptor(path)
        if descriptor is None:
            return False
        snapshot = Path(descriptor.path)
        try:
            if not snapshot.is_file() or snapshot.stat().st_size != descriptor.size_bytes:
                return False
            digest = hashlib.sha256()
            with snapshot.open("rb") as stream:
                while chunk := stream.read(1024 * 1024):
                    digest.update(chunk)
            return digest.hexdigest() == descriptor.sha256
        except OSError:
            return False

    def retain_inspection(self, owner: str, paths: list[str] | tuple[str, ...]) -> None:
        retained = {key for path in paths if (key := self._key(path)) in self._snapshots}
        if retained:
            self._inspection[owner] = retained

    def release_inspection(self, owner: str) -> None:
        self._inspection.pop(owner, None)
        self._collect_unowned()

    def sync_visible(self, paths: list[str] | tuple[str, ...]) -> None:
        self._visible = {key for path in paths if (key := self._key(path)) in self._snapshots}
        self._collect_unowned()

    def retain_active(self, owner: str, paths: list[str] | tuple[str, ...]) -> None:
        retained = {key for path in paths if (key := self._key(path)) in self._snapshots}
        if retained:
            self._active[owner] = retained

    def release_active(self, owner: str) -> None:
        self._active.pop(owner, None)
        self._collect_unowned()

    def retain_history(self, owner: str, paths: list[str] | tuple[str, ...]) -> None:
        retained = {key for path in paths if (key := self._key(path)) in self._snapshots}
        if retained:
            self._history[owner] = retained
        else:
            self._history.pop(owner, None)

    def release_history(self, owner: str) -> None:
        self._history.pop(owner, None)
        self._collect_unowned()

    def discard_if_unowned(self, path: str | os.PathLike[str]) -> None:
        self._collect_unowned({self._key(path)})

    def _owned_keys(self) -> set[str]:
        owned = set(self._visible)
        for paths in self._inspection.values():
            owned.update(paths)
        for paths in self._active.values():
            owned.update(paths)
        for paths in self._history.values():
            owned.update(paths)
        return owned

    def _collect_unowned(self, candidates: set[str] | None = None) -> None:
        owned = self._owned_keys()
        keys = set(self._snapshots) if candidates is None else candidates & set(self._snapshots)
        for key in keys - owned:
            descriptor = self._snapshots.pop(key)
            try:
                Path(descriptor.path).unlink(missing_ok=True)
            except OSError:
                self._snapshots[key] = descriptor

    def close(self) -> None:
        """Release owned files without forgetting cleanup failures."""

        self._visible.clear()
        self._history.clear()
        self._collect_unowned()
        if self._inspection or self._active or self._snapshots:
            self._cleanup_pending = True
            return

        try:
            with _session_namespace_lock(self._root):
                self._release_owner_lock()
                shutil.rmtree(self._session_root)
        except FileNotFoundError:
            self._cleanup_pending = False
        except OSError:
            self._cleanup_pending = True
        else:
            self._cleanup_pending = False


__all__ = ["ClipboardInputDescriptor", "ClipboardInputStore", "bounded_plaintext_preview"]
