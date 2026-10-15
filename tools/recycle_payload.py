"""Read a relocation-stable payload identity without following reparse targets."""

from __future__ import annotations

import hashlib
import json
import os
import stat
from contextlib import contextmanager
from pathlib import Path
from typing import Any, BinaryIO

PAYLOAD_SCHEMA = "docwen.recycle-payload.v1"


def _extended(path: Path) -> str:
    value = str(path)
    if os.name != "nt" or value.startswith("\\\\?\\"):
        return value
    if value.startswith("\\\\"):
        return "\\\\?\\UNC\\" + value[2:]
    return "\\\\?\\" + value


def _metadata(path: Path) -> os.stat_result:
    return os.lstat(_extended(path))


def _reparse(metadata: os.stat_result) -> bool:
    return stat.S_ISLNK(metadata.st_mode) or bool(getattr(metadata, "st_file_attributes", 0) & 0x400)


def _same(before: os.stat_result, after: os.stat_result, *, cross_api: bool = False) -> bool:
    windows_cross_api = cross_api and os.name == "nt"
    return (
        before.st_dev,
        before.st_ino,
        stat.S_IFMT(before.st_mode) if windows_cross_api else before.st_mode,
        before.st_size if stat.S_ISREG(before.st_mode) and not _reparse(before) else None,
        before.st_mtime_ns,
        None if windows_cross_api else before.st_ctime_ns,
        getattr(before, "st_file_attributes", 0),
        getattr(before, "st_reparse_tag", 0),
    ) == (
        after.st_dev,
        after.st_ino,
        stat.S_IFMT(after.st_mode) if windows_cross_api else after.st_mode,
        after.st_size if stat.S_ISREG(after.st_mode) and not _reparse(after) else None,
        after.st_mtime_ns,
        None if windows_cross_api else after.st_ctime_ns,
        getattr(after, "st_file_attributes", 0),
        getattr(after, "st_reparse_tag", 0),
    )


def assert_plain_chain(path: Path) -> None:
    current = path
    while True:
        if _reparse(_metadata(current)):
            raise OSError(f"recycle_payload_linked_boundary:{current}")
        if current.parent == current:
            return
        current = current.parent


@contextmanager
def plain_reader(path: Path, expected: os.stat_result):
    """Open the entry itself; never open a link's target as regular file contents."""
    assert_plain_chain(path.parent)
    if os.name == "nt":
        import msvcrt

        import win32api
        import win32file

        handle = win32file.CreateFile(
            _extended(path), 0x80000000, 7, None, 3, 0x00200000 | 0x02000000, None
        )  # OPEN_REPARSE_POINT | BACKUP_SEMANTICS
        raw_handle = handle.Detach()
        try:
            descriptor = msvcrt.open_osfhandle(raw_handle, os.O_RDONLY | os.O_BINARY)
        except BaseException:
            win32api.CloseHandle(raw_handle)
            raise
    else:
        descriptor = os.open(path, os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0))
    stream: BinaryIO = os.fdopen(descriptor, "rb")
    try:
        observed = os.fstat(stream.fileno())
        # Windows lstat adds extension-derived permission bits and can report a
        # different ctime from CRT-backed fstat. Keep full mode/ctime within each
        # API; cross-API identity still binds every
        # File ID, type, size, mtime and reparse/attribute observation.
        if _reparse(observed) or not stat.S_ISREG(observed.st_mode) or not _same(expected, observed, cross_api=True):
            raise OSError(f"recycle_payload_entry_changed:{path}")
        yield stream
        if not _same(observed, os.fstat(stream.fileno())) or not _same(expected, _metadata(path)):
            raise OSError(f"recycle_payload_changed_during_read:{path}")
    finally:
        stream.close()


def read_plain_bytes(path: Path, *, max_bytes: int) -> bytes:
    metadata = _metadata(path)
    if _reparse(metadata) or not stat.S_ISREG(metadata.st_mode) or metadata.st_size > max_bytes:
        raise OSError(f"recycle_metadata_not_plain_or_bounded:{path}")
    with plain_reader(path, metadata) as stream:
        data = stream.read(max_bytes + 1)
    if len(data) > max_bytes:
        raise OSError(f"recycle_metadata_too_large:{path}")
    return data


def _reparse_data(path: Path) -> bytes:
    assert_plain_chain(path.parent)
    if os.name == "nt":
        import win32file

        handle = win32file.CreateFile(_extended(path), 0, 7, None, 3, 0x00200000 | 0x02000000, None)
        try:
            return win32file.DeviceIoControl(handle, 0x000900A8, None, 16384)
        finally:
            handle.Close()
    return os.fsencode(os.readlink(path))


def snapshot_payload(root: Path) -> dict[str, Any]:
    """Hash physical entries. Lease text and absolute link text remain original bytes."""
    root = Path(os.path.abspath(root))
    assert_plain_chain(root.parent)
    before_root = _metadata(root)
    if _reparse(before_root) or not (stat.S_ISDIR(before_root.st_mode) or stat.S_ISREG(before_root.st_mode)):
        raise OSError(f"recycle_payload_plain_file_or_directory_required:{root}")
    content = hashlib.sha256()
    records: list[dict[str, Any]] = []
    links: list[dict[str, Any]] = []
    visited: list[tuple[Path, os.stat_result]] = []
    total_bytes = files = directories = 0

    def record(path: Path, metadata: os.stat_result, kind: str) -> dict[str, Any]:
        row = {
            "path": "." if path == root else str(path.relative_to(root)).replace(os.sep, "/"),
            "kind": kind,
            "device": metadata.st_dev,
            "inode": metadata.st_ino,
            "mode": stat.S_IFMT(metadata.st_mode),
            "size": metadata.st_size if kind == "file" else 0,
        }
        records.append(row)
        visited.append((path, metadata))
        return row

    def regular_file(path: Path, metadata: os.stat_result) -> None:
        nonlocal total_bytes, files
        row = record(path, metadata, "file")
        content.update(row["path"].encode("utf-8"))
        content.update(b"\0")
        with plain_reader(path, metadata) as stream:
            while block := stream.read(1024 * 1024):
                content.update(block)
        content.update(b"\0")
        total_bytes += metadata.st_size
        files += 1

    if stat.S_ISREG(before_root.st_mode):
        regular_file(root, before_root)
    else:
        record(root, before_root, "directory")
        directories += 1
        # Match the existing housekeeping content stream: files within each
        # directory, followed by sorted plain children. Reparse entries never recurse.
        stack = [root]
        while stack:
            current = stack.pop()
            assert_plain_chain(current)
            plain_children: list[Path] = []
            with os.scandir(_extended(current)) as scan:
                names = sorted(entry.name for entry in scan)
            regular: list[tuple[Path, os.stat_result]] = []
            for name in names:
                child = current / name
                metadata = _metadata(child)
                if _reparse(metadata):
                    row = record(child, metadata, "reparse")
                    raw = _reparse_data(child)
                    row.update(tag=getattr(metadata, "st_reparse_tag", 0), rawSha256=hashlib.sha256(raw).hexdigest())
                    links.append(dict(row))
                    if not _same(metadata, _metadata(child)):
                        raise OSError(f"recycle_reparse_changed_during_read:{child}")
                elif stat.S_ISDIR(metadata.st_mode):
                    record(child, metadata, "directory")
                    directories += 1
                    plain_children.append(child)
                elif stat.S_ISREG(metadata.st_mode):
                    regular.append((child, metadata))
                else:
                    raise OSError(f"recycle_payload_special_file:{child}")
            for child, metadata in regular:
                regular_file(child, metadata)
            stack.extend(reversed(plain_children))
    for path, metadata in visited:
        if not _same(metadata, _metadata(path)):
            raise OSError(f"recycle_payload_changed_during_snapshot:{path}")
    records.sort(key=lambda row: row["path"])
    return {
        "schema": PAYLOAD_SCHEMA,
        "kind": "directory" if stat.S_ISDIR(before_root.st_mode) else "file",
        "rootDevice": before_root.st_dev,
        "rootInode": before_root.st_ino,
        "bytes": total_bytes,
        "files": files,
        "directories": directories,
        "contentSha256": content.hexdigest(),
        "entriesSha256": hashlib.sha256(
            json.dumps(records, sort_keys=True, separators=(",", ":")).encode()
        ).hexdigest(),
        "reparsePoints": sorted(links, key=lambda row: row["path"]),
    }


def require_payload_identity(identity: object) -> dict[str, Any]:
    if not isinstance(identity, dict) or identity.get("schema") != PAYLOAD_SCHEMA:
        raise ValueError("recycle_payload_proof_missing_regenerate_plan")
    for field in ("rootDevice", "rootInode", "bytes", "files", "directories"):
        if not isinstance(identity.get(field), int) or isinstance(identity[field], bool) or identity[field] < 0:
            raise ValueError(f"invalid_recycle_payload_identity:{field}")
    for field in ("contentSha256", "entriesSha256"):
        value = identity.get(field)
        if (
            not isinstance(value, str)
            or len(value) != 64
            or any(character not in "0123456789abcdef" for character in value)
        ):
            raise ValueError(f"invalid_recycle_payload_identity:{field}")
    if identity.get("kind") not in {"file", "directory"} or not isinstance(identity.get("reparsePoints"), list):
        raise ValueError("invalid_recycle_payload_shape")
    return identity
