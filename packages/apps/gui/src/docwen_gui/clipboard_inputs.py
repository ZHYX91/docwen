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

from docwen_core.models.clipboard_document import (
    MAX_CLIPBOARD_IMAGE_PIXELS,
    MAX_CLIPBOARD_RESOURCE_BYTES,
    MAX_CLIPBOARD_RESOURCES,
    load_clipboard_document_bytes,
)
from docwen_gui.clipboard_image_bytes import ClipboardImageBytesError, inspect_png_bytes, preflight_png_resources

_PREVIEW_MAX_CHARS = 240
_PREVIEW_MAX_LINES = 3
_SESSION_PREFIX = "session-"
_OWNER_LOCK_NAME = ".owner.lock"
_OWNER_MARKER = b"docwen-clipboard-session-v1\n"
_BUNDLE_MARKER = b"docwen-clipboard-bundle-v1\n"
_BUNDLE_MARKER_NAME = ".bundle.marker"
_BUNDLE_PARTIAL_NAME = ".bundle.partial"


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


@dataclass(frozen=True, slots=True)
class ClipboardResourceDescriptor:
    resource_id: str
    path: str
    logical_path: str
    media_type: str
    size_bytes: int
    sha256: str


@dataclass(frozen=True, slots=True)
class ClipboardSnapshotBundle:
    root_path: str
    main: ClipboardInputDescriptor
    resources: tuple[ClipboardResourceDescriptor, ...]
    marker_path: str


def _frozen_file_available(path: Path, size_bytes: int, sha256: str) -> bool:
    stat = path.lstat()
    if path.is_symlink() or getattr(stat, "st_file_attributes", 0) & 0x400:
        return False
    if not path.is_file() or stat.st_size != size_bytes:
        return False
    digest = hashlib.sha256()
    remaining = size_bytes
    with path.open("rb") as stream:
        while remaining:
            chunk = stream.read(min(1024 * 1024, remaining))
            if not chunk:
                return False
            digest.update(chunk)
            remaining -= len(chunk)
        if stream.read(1):
            return False
    return digest.hexdigest() == sha256


def _bundle_member_available(root: Path, path: Path) -> bool:
    if not root.is_absolute() or not path.is_absolute():
        return False
    try:
        relative = path.relative_to(root)
    except ValueError:
        return False
    if ".." in relative.parts:
        return False
    for member in (path, *path.parents):
        stat = member.lstat()
        if member.is_symlink() or getattr(stat, "st_file_attributes", 0) & 0x400:
            return False
        if member == root:
            return True
    return False


def clipboard_bundle_available(bundle: ClipboardSnapshotBundle) -> bool:
    """Revalidate frozen Store facts without accessing a mutable Store or Qt."""
    try:
        root = Path(bundle.root_path)
        marker = Path(bundle.marker_path)
        if marker != root / _BUNDLE_MARKER_NAME or not _bundle_member_available(root, marker):
            return False
        if not root.is_dir() or not marker.is_file() or os.path.lexists(root / _BUNDLE_PARTIAL_NAME):
            return False
        with marker.open("rb") as stream:
            if stream.read(len(_BUNDLE_MARKER) + 1) != _BUNDLE_MARKER:
                return False
        for descriptor in (bundle.main, *bundle.resources):
            candidate = Path(descriptor.path)
            if not _bundle_member_available(root, candidate):
                return False
            if not _frozen_file_available(candidate, descriptor.size_bytes, descriptor.sha256):
                return False
        return True
    except OSError:
        return False


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
        self._bundles: dict[str, ClipboardSnapshotBundle] = {}
        self._bundle_members: dict[str, str] = {}
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

    def create_binary(
        self,
        payload: bytes,
        *,
        suffix: str,
        display_name_template: str,
        preview: str = "",
    ) -> ClipboardInputDescriptor:
        """Freeze one already-normalized binary input as a session-owned source."""

        if not isinstance(payload, bytes) or not payload or not suffix.startswith("."):
            raise ValueError("clipboard binary snapshot is invalid")
        self._sequence += 1
        display_name = display_name_template.format(index=self._sequence).strip()
        if not display_name:
            display_name = f"Clipboard Input {self._sequence}{suffix}"
        path = self._session_root / f"clipboard-{uuid.uuid4().hex}{suffix}"
        descriptor = ClipboardInputDescriptor(
            path=str(path),
            display_name=display_name,
            preview=preview[:_PREVIEW_MAX_CHARS],
            size_bytes=len(payload),
            sha256=hashlib.sha256(payload).hexdigest(),
        )
        try:
            self._write_fsynced(path, payload)
        except BaseException:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                self._snapshots[self._key(path)] = descriptor
            raise
        self._snapshots[self._key(path)] = descriptor
        return descriptor

    @staticmethod
    def _write_fsynced(path: Path, payload: bytes) -> None:
        with path.open("xb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())

    def create_bundle(
        self,
        payload: bytes,
        *,
        display_name_template: str,
        preview: str,
        suffix: str = ".dwclip",
        resources: tuple[tuple[str, str, str, bytes], ...] = (),
    ) -> ClipboardSnapshotBundle:
        """Freeze one structured main input and all of its resource bytes as one owner group."""

        if not isinstance(payload, bytes) or not payload:
            raise ValueError("clipboard bundle main payload must be non-empty bytes")
        document = load_clipboard_document_bytes(payload)
        supplied: dict[str, tuple[str, str, bytes]] = {}
        for resource_id, logical_path, media_type, resource_bytes in resources:
            if (
                not resource_id
                or resource_id in supplied
                or not logical_path
                or not media_type
                or not isinstance(resource_bytes, bytes)
            ):
                raise ValueError("invalid clipboard bundle resource")
            supplied[resource_id] = (logical_path, media_type, resource_bytes)
        declarations = {item.resource_id: item for item in document.resources}
        if set(supplied) != set(declarations):
            raise ValueError("clipboard bundle resources do not match document declarations")
        try:
            preflight_png_resources(
                (supplied[item.resource_id][2] for item in document.resources if item.media_type == "image/png"),
                max_resources=MAX_CLIPBOARD_RESOURCES,
                max_bytes=MAX_CLIPBOARD_RESOURCE_BYTES,
                max_pixels=MAX_CLIPBOARD_IMAGE_PIXELS,
            )
        except ClipboardImageBytesError as exc:
            raise ValueError("clipboard bundle PNG resource is invalid") from exc
        ordered_resources: list[tuple[str, str, str, bytes]] = []
        total_resource_bytes = 0
        for declaration in document.resources:
            logical_path, media_type, resource_bytes = supplied[declaration.resource_id]
            digest = hashlib.sha256(resource_bytes).hexdigest()
            if declaration.media_type == "image/png":
                try:
                    frozen_png = inspect_png_bytes(resource_bytes)
                except ClipboardImageBytesError as exc:
                    raise ValueError("clipboard bundle PNG resource is invalid") from exc
                if (
                    declaration.pixel_width != frozen_png.width
                    or declaration.pixel_height != frozen_png.height
                    or declaration.rgba_sha256 != frozen_png.rgba_sha256
                ):
                    raise ValueError("clipboard bundle PNG pixel identity does not match document declaration")
            if (
                logical_path != declaration.logical_path
                or media_type != declaration.media_type
                or len(resource_bytes) != declaration.size_bytes
                or digest != declaration.sha256
            ):
                raise ValueError("clipboard bundle resource identity does not match document declaration")
            total_resource_bytes += len(resource_bytes)
            ordered_resources.append((declaration.resource_id, logical_path, media_type, resource_bytes))
        if total_resource_bytes > MAX_CLIPBOARD_RESOURCE_BYTES:
            raise ValueError("clipboard bundle resource byte budget exceeded")
        self._sequence += 1
        display_name = display_name_template.format(index=self._sequence).strip()
        if not display_name:
            display_name = f"Clipboard Document {self._sequence}{suffix}"
        bundle_root = self._session_root / f"bundle-{uuid.uuid4().hex}"
        main_path = bundle_root / f"main{suffix}"
        marker_path = bundle_root / _BUNDLE_MARKER_NAME
        partial_path = bundle_root / _BUNDLE_PARTIAL_NAME
        main = ClipboardInputDescriptor(
            path=str(main_path),
            display_name=display_name,
            preview=preview[:_PREVIEW_MAX_CHARS],
            size_bytes=len(payload),
            sha256=hashlib.sha256(payload).hexdigest(),
        )
        resource_descriptors: list[ClipboardResourceDescriptor] = []
        seen_ids: set[str] = set()
        seen_logical: set[str] = set()
        try:
            bundle_root.mkdir()
            self._write_fsynced(partial_path, b"creating\n")
            resource_root = bundle_root / "resources"
            if resources:
                resource_root.mkdir()
            for index, (resource_id, logical_path, media_type, resource_bytes) in enumerate(ordered_resources):
                if (
                    not resource_id
                    or resource_id in seen_ids
                    or not logical_path
                    or logical_path in seen_logical
                    or not isinstance(resource_bytes, bytes)
                ):
                    raise ValueError("invalid clipboard bundle resource")
                seen_ids.add(resource_id)
                seen_logical.add(logical_path)
                resource_path = resource_root / f"resource-{index:04d}"
                self._write_fsynced(resource_path, resource_bytes)
                resource_descriptors.append(
                    ClipboardResourceDescriptor(
                        resource_id=resource_id,
                        path=str(resource_path),
                        logical_path=logical_path,
                        media_type=media_type,
                        size_bytes=len(resource_bytes),
                        sha256=hashlib.sha256(resource_bytes).hexdigest(),
                    )
                )
            self._write_fsynced(main_path, payload)
            partial_path.unlink()
            self._write_fsynced(marker_path, _BUNDLE_MARKER)
        except BaseException:
            try:
                shutil.rmtree(bundle_root)
            except OSError:
                self._cleanup_pending = True
            raise

        bundle = ClipboardSnapshotBundle(
            root_path=str(bundle_root),
            main=main,
            resources=tuple(resource_descriptors),
            marker_path=str(marker_path),
        )
        main_key = self._key(main.path)
        self._snapshots[main_key] = main
        self._bundles[main_key] = bundle
        self._bundle_members[main_key] = main_key
        for resource in resource_descriptors:
            self._bundle_members[self._key(resource.path)] = main_key
        return bundle

    def _snapshot_key(self, path: str | os.PathLike[str]) -> str | None:
        key = self._key(path)
        if key in self._snapshots:
            return key
        return self._bundle_members.get(key)

    def descriptor(self, path: str | os.PathLike[str]) -> ClipboardInputDescriptor | None:
        key = self._snapshot_key(path)
        return self._snapshots.get(key) if key is not None else None

    def bundle(self, path: str | os.PathLike[str]) -> ClipboardSnapshotBundle | None:
        key = self._snapshot_key(path)
        return self._bundles.get(key) if key is not None else None

    def is_snapshot(self, path: str | os.PathLike[str]) -> bool:
        return self._snapshot_key(path) is not None

    @staticmethod
    def _descriptor_available(path: str, size_bytes: int, sha256: str) -> bool:
        return _frozen_file_available(Path(path), size_bytes, sha256)

    def snapshot_available(self, path: str | os.PathLike[str]) -> bool:
        key = self._snapshot_key(path)
        if key is None:
            return False
        descriptor = self._snapshots[key]
        bundle = self._bundles.get(key)
        try:
            if bundle is not None:
                return clipboard_bundle_available(bundle)
            return self._descriptor_available(descriptor.path, descriptor.size_bytes, descriptor.sha256)
        except OSError:
            return False

    def retain_inspection(self, owner: str, paths: list[str] | tuple[str, ...]) -> None:
        retained = {key for path in paths if (key := self._snapshot_key(path)) is not None}
        if retained:
            self._inspection[owner] = retained

    def release_inspection(self, owner: str) -> None:
        self._inspection.pop(owner, None)
        self._collect_unowned()

    def sync_visible(self, paths: list[str] | tuple[str, ...]) -> None:
        self._visible = {key for path in paths if (key := self._snapshot_key(path)) is not None}
        self._collect_unowned()

    def retain_active(self, owner: str, paths: list[str] | tuple[str, ...]) -> None:
        retained = {key for path in paths if (key := self._snapshot_key(path)) is not None}
        if retained:
            self._active[owner] = retained

    def release_active(self, owner: str) -> None:
        self._active.pop(owner, None)
        self._collect_unowned()

    def retain_history(self, owner: str, paths: list[str] | tuple[str, ...]) -> None:
        retained = {key for path in paths if (key := self._snapshot_key(path)) is not None}
        if retained:
            self._history[owner] = retained
        else:
            self._history.pop(owner, None)

    def release_history(self, owner: str) -> None:
        self._history.pop(owner, None)
        self._collect_unowned()

    def discard_if_unowned(self, path: str | os.PathLike[str]) -> None:
        key = self._snapshot_key(path)
        if key is not None:
            self._collect_unowned({key})

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
            descriptor = self._snapshots[key]
            bundle = self._bundles.get(key)
            try:
                if bundle is not None:
                    shutil.rmtree(bundle.root_path)
                else:
                    Path(descriptor.path).unlink(missing_ok=True)
            except OSError:
                continue
            self._snapshots.pop(key, None)
            removed_bundle = self._bundles.pop(key, None)
            if removed_bundle is not None:
                self._bundle_members.pop(key, None)
                for resource in removed_bundle.resources:
                    self._bundle_members.pop(self._key(resource.path), None)

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


__all__ = [
    "ClipboardInputDescriptor",
    "ClipboardInputStore",
    "ClipboardResourceDescriptor",
    "ClipboardSnapshotBundle",
    "bounded_plaintext_preview",
]
