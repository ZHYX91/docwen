"""Managed, session-owned materialization for user-triggered clipboard Markdown."""

from __future__ import annotations

import os
import shutil
import unicodedata
import uuid
from dataclasses import dataclass
from pathlib import Path

_PREVIEW_MAX_CHARS = 240
_PREVIEW_MAX_LINES = 3


@dataclass(frozen=True, slots=True)
class ClipboardInputDescriptor:
    path: str
    display_name: str
    preview: str
    size_bytes: int


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
        self._session_root = root / f"session-{uuid.uuid4().hex}"
        self._session_root.mkdir()
        self._snapshots: dict[str, ClipboardInputDescriptor] = {}
        self._visible: set[str] = set()
        self._active: dict[str, set[str]] = {}
        self._history: dict[str, set[str]] = {}
        self._sequence = 0

    @staticmethod
    def _key(path: str | os.PathLike[str]) -> str:
        return os.path.normcase(os.path.abspath(os.fspath(path)))

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
        with path.open("xb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        descriptor = ClipboardInputDescriptor(
            path=str(path),
            display_name=display_name,
            preview=bounded_plaintext_preview(text),
            size_bytes=len(payload),
        )
        self._snapshots[self._key(path)] = descriptor
        return descriptor

    def descriptor(self, path: str | os.PathLike[str]) -> ClipboardInputDescriptor | None:
        return self._snapshots.get(self._key(path))

    def is_snapshot(self, path: str | os.PathLike[str]) -> bool:
        return self._key(path) in self._snapshots

    def snapshot_available(self, path: str | os.PathLike[str]) -> bool:
        descriptor = self.descriptor(path)
        return descriptor is not None and Path(descriptor.path).is_file()

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
            self._history.setdefault(owner, set()).update(retained)

    def release_history(self, owner: str) -> None:
        self._history.pop(owner, None)
        self._collect_unowned()

    def discard_if_unowned(self, path: str | os.PathLike[str]) -> None:
        self._collect_unowned({self._key(path)})

    def _owned_keys(self) -> set[str]:
        owned = set(self._visible)
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
        """Release session-owned snapshots after active workers have drained."""

        self._visible.clear()
        self._history.clear()
        self._collect_unowned()
        if self._active:
            return
        shutil.rmtree(self._session_root, ignore_errors=True)
        self._snapshots.clear()


__all__ = ["ClipboardInputDescriptor", "ClipboardInputStore", "bounded_plaintext_preview"]
