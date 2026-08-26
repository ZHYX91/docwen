from __future__ import annotations

import hashlib
import json
import os
import stat
from pathlib import Path
from typing import Any

from scripts.release import candidate_evidence


class V4EvidenceContractError(RuntimeError):
    """Evidence bytes and paths are unsafe or ambiguous."""


def validate_relative_path(value: object, *, label: str) -> str:
    if not isinstance(value, str) or not value or "\\" in value or ":" in value:
        raise V4EvidenceContractError(f"{label}_path_invalid:{value}")
    path = Path(value)
    if path.is_absolute() or path.drive or any(part in {"", ".", ".."} for part in value.split("/")):
        raise V4EvidenceContractError(f"{label}_path_invalid:{value}")
    return value


def safe_regular_file(path: Path, *, label: str) -> Path:
    if any(":" in part for part in Path(os.path.abspath(path)).parts[1:]):
        raise V4EvidenceContractError(f"{label}_ads_or_colon_path_rejected:{path}")
    try:
        safe = candidate_evidence._safe_regular_file(path, label=label)
    except candidate_evidence.EvidenceError as exc:
        raise V4EvidenceContractError(str(exc)) from exc
    stats = safe.lstat()
    if not stat.S_ISREG(stats.st_mode) or stats.st_nlink != 1:
        raise V4EvidenceContractError(f"{label}_hardlink_or_nonregular_rejected:{safe}")
    return safe


def _signature(stats: os.stat_result) -> tuple[int, int, int, int, int]:
    return (
        stats.st_ino,
        stats.st_dev,
        stats.st_nlink,
        stats.st_size,
        stats.st_mtime_ns,
    )


def _read_stable(path: Path, *, label: str, collect: bool) -> tuple[Path, bytes | None, int, str]:
    safe = safe_regular_file(path, label=label)
    before = safe.lstat()
    digest = hashlib.sha256()
    chunks: list[bytes] | None = [] if collect else None
    with safe.open("rb") as handle:
        opened_before = os.fstat(handle.fileno())
        while chunk := handle.read(1024 * 1024):
            digest.update(chunk)
            if chunks is not None:
                chunks.append(chunk)
        opened_after = os.fstat(handle.fileno())
    after = safe.lstat()
    if (
        not (_signature(before) == _signature(opened_before) == _signature(opened_after) == _signature(after))
        or before.st_mode != after.st_mode
        or opened_before.st_mode != opened_after.st_mode
        or before.st_ctime_ns != after.st_ctime_ns
        or opened_before.st_ctime_ns != opened_after.st_ctime_ns
    ):
        raise V4EvidenceContractError(f"{label}_changed_during_read:{safe}")
    return safe, b"".join(chunks) if chunks is not None else None, after.st_size, digest.hexdigest()


def _root(path: Path) -> Path:
    try:
        return candidate_evidence._safe_existing_directory(path, label="identity_root")
    except candidate_evidence.EvidenceError as exc:
        raise V4EvidenceContractError(str(exc)) from exc


def file_identity(path: Path, *, relative_to: Path, label: str = "identity_file") -> dict[str, object]:
    root = _root(relative_to)
    safe, _, size, digest = _read_stable(path, label=label, collect=False)
    try:
        relative = safe.relative_to(root).as_posix()
    except ValueError as exc:
        raise V4EvidenceContractError(f"identity_outside_root:{safe}:{root}") from exc
    validate_relative_path(relative, label="identity")
    return {"relativePath": relative, "bytes": size, "sha256": digest}


def sha256_file(path: Path, *, label: str = "sha256_file") -> str:
    return _read_stable(path, label=label, collect=False)[3]


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            raise V4EvidenceContractError(f"duplicate_json_key:{key}")
        value[key] = item
    return value


def read_json_object(
    path: Path, *, relative_to: Path, label: str, expected_sha256: str | None = None
) -> tuple[dict[str, Any], dict[str, object]]:
    safe, raw, size, digest = _read_stable(path, label=label, collect=True)
    if expected_sha256 is not None and digest != expected_sha256:
        raise V4EvidenceContractError(f"{label}_sha256_mismatch")
    assert raw is not None
    try:
        value = json.loads(raw.decode("utf-8", errors="strict"), object_pairs_hook=_reject_duplicate_keys)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise V4EvidenceContractError(f"{label}_invalid_json:{safe}") from exc
    if not isinstance(value, dict):
        raise V4EvidenceContractError(f"{label}_not_object:{safe}")
    try:
        relative = safe.relative_to(_root(relative_to)).as_posix()
    except ValueError as exc:
        raise V4EvidenceContractError(f"identity_outside_root:{safe}:{relative_to}") from exc
    validate_relative_path(relative, label="identity")
    return value, {"relativePath": relative, "bytes": size, "sha256": digest}


def capture_tree_stable(root: Path) -> dict[str, object]:
    try:
        first = candidate_evidence._capture_tree_stable(root)
    except candidate_evidence.EvidenceError as exc:
        raise V4EvidenceContractError(str(exc)) from exc
    raw_files = first.get("files")
    if not isinstance(raw_files, list):
        raise V4EvidenceContractError("tree_files_invalid")
    inodes: set[tuple[int, int]] = set()
    hardlinks: list[str] = []
    for raw in raw_files:
        if not isinstance(raw, dict):
            raise V4EvidenceContractError("tree_file_invalid")
        relative = validate_relative_path(raw.get("path"), label="tree")
        try:
            safe = candidate_evidence._safe_regular_file(root / relative, label="tree_file")
        except candidate_evidence.EvidenceError as exc:
            raise V4EvidenceContractError(str(exc)) from exc
        stats = safe.lstat()
        inode = (stats.st_dev, stats.st_ino)
        if inode in inodes:
            raise V4EvidenceContractError(f"duplicate_tree_inode:{relative}")
        inodes.add(inode)
        if stats.st_nlink != 1:
            hardlinks.append(relative)
    if hardlinks:
        raise V4EvidenceContractError(f"tree_hardlink_rejected:{hardlinks[0]}")
    try:
        second = candidate_evidence._capture_tree_stable(root)
    except candidate_evidence.EvidenceError as exc:
        raise V4EvidenceContractError(str(exc)) from exc
    if first != second:
        raise V4EvidenceContractError("tree_changed_during_safe_snapshot")
    return first
