from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import stat
import sys
from collections.abc import Sequence
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any, NoReturn

if __package__ in {None, ""}:
    _BOOTSTRAP_ROOT = Path(__file__).resolve().parents[1]
    if str(_BOOTSTRAP_ROOT) not in sys.path:
        sys.path.insert(0, str(_BOOTSTRAP_ROOT))

from tools.windows_short_path import ShortPathDriveError, unmount_short_drive
from tools.workspace_root import WORKSPACE_ROOT_ENV as _WORKSPACE_ROOT_ENV
from tools.workspace_root import resolve_workspace_root

LEASE_NAME = ".docwen-temp-lease.json"
PLAN_SCHEMA = "docwen.housekeeping-plan.v1"
WORKSPACE_ROOT_ENV = _WORKSPACE_ROOT_ENV
DEFAULT_FAILURE_TTL = timedelta(hours=72)
DEFAULT_FAILURE_MAX_PER_KIND = 2
MANAGED_ROOT_NAMES = ("temp", "build", "tmp")
PROTECTED_WORKSPACE_ROOT_NAMES = (
    "acceptance",
    "artifacts",
    "backups",
    "tools",
    "quarantine",
)
DEPENDENCY_DIRECTORY_NAMES = frozenset({".venv", "node_modules"})
SUCCESS_STATES = frozenset({"clean", "completed", "completed-success", "success", "succeeded"})
FAILURE_STATES = frozenset(
    {
        "failed",
        "failure",
        "interrupted",
        "retained-cleanup-failure",
        "retained-failure",
        "retained-interrupted",
    }
)


class HousekeepingError(ValueError):
    """A housekeeping plan or target failed a safety boundary."""


def _absolute(path: Path) -> Path:
    return Path(os.path.abspath(os.fspath(path)))


def _is_reparse(metadata: os.stat_result) -> bool:
    flag = int(getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
    return bool(flag and int(getattr(metadata, "st_file_attributes", 0)) & flag)


def _is_link_or_reparse(metadata: os.stat_result) -> bool:
    return stat.S_ISLNK(metadata.st_mode) or _is_reparse(metadata)


def _is_within(path: Path, parent: Path) -> bool:
    absolute_path = _absolute(path)
    absolute_parent = _absolute(parent)
    return absolute_path == absolute_parent or absolute_parent in absolute_path.parents


def _chain_is_plain(path: Path, *, boundary: Path) -> bool:
    current = _absolute(path)
    stop = _absolute(boundary)
    while True:
        metadata = current.lstat()
        if _is_link_or_reparse(metadata):
            return False
        if current == stop:
            return True
        parent = current.parent
        if parent == current:
            return False
        current = parent


def _raise_walk_error(error: OSError) -> NoReturn:
    raise error


def _read_json_object(path: Path) -> dict[str, Any]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise HousekeepingError(f"json_object_required:{path}")
    return payload


def _canonical_json(payload: object) -> bytes:
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def _sha256_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def _read_lease(marker: Path, *, temp_root: Path) -> tuple[Path, dict[str, Any]]:
    runtime_root = marker.parent
    if not _chain_is_plain(runtime_root, boundary=temp_root):
        raise ValueError("linked_or_reparse_path")
    payload = _read_json_object(marker)
    if payload.get("schemaVersion") != 1:
        raise ValueError("invalid_schema")
    if not str(payload.get("owner", "")).startswith("docwen."):
        raise ValueError("foreign_owner")
    if payload.get("root") != str(runtime_root.resolve(strict=True)):
        raise ValueError("root_mismatch")
    return runtime_root, payload


def _created_at(payload: dict[str, Any]) -> datetime:
    value = payload.get("createdAt")
    if not isinstance(value, str):
        raise ValueError("created_at_missing")
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        raise ValueError("created_at_timezone_missing")
    return parsed.astimezone(UTC)


def _process_alive(pid: object) -> bool:
    if not isinstance(pid, int) or pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


def _lease_markers(temp_root: Path, *, errors: list[dict[str, str]] | None = None) -> list[Path]:
    markers: list[Path] = []

    def record_walk_error(error: OSError) -> None:
        if errors is not None:
            errors.append({"path": str(getattr(error, "filename", temp_root)), "reason": str(error)})

    for current, directories, files in os.walk(
        temp_root,
        topdown=True,
        followlinks=False,
        onerror=record_walk_error,
    ):
        current_path = Path(current)
        safe_directories: list[str] = []
        for name in sorted(directories):
            child = current_path / name
            if name.casefold() == ".workspace":
                if errors is not None:
                    errors.append({"path": str(child), "reason": "nested_workspace_boundary"})
                continue
            try:
                metadata = child.lstat()
            except OSError as error:
                record_walk_error(error)
                continue
            if not _is_link_or_reparse(metadata):
                safe_directories.append(name)
        directories[:] = safe_directories
        if LEASE_NAME in files:
            markers.append(current_path / LEASE_NAME)
            directories[:] = []
    return sorted(markers, key=lambda path: os.path.normcase(str(path)))


def _windows_extended_path(path: Path) -> str:
    absolute = os.path.abspath(os.fspath(path))
    if os.name != "nt" or absolute.startswith("\\\\?\\"):
        return absolute
    if absolute.startswith("\\\\"):
        return f"\\\\?\\UNC\\{absolute[2:]}"
    return f"\\\\?\\{absolute}"


def _logical_path(path: str | Path) -> Path:
    """Return a normal absolute path after extended-length Windows I/O."""

    value = os.fspath(path)
    if os.name != "nt":
        return Path(value)
    if value.startswith("\\\\?\\UNC\\"):
        return Path(f"\\\\{value[8:]}")
    if value.startswith("\\\\?\\"):
        return Path(value[4:])
    return Path(value)


def _remove_owned_readonly_path(function: object, path: str, error: BaseException) -> None:
    if not isinstance(error, PermissionError) or not callable(function):
        raise error
    os.chmod(path, stat.S_IWRITE)
    function(path)


def cleanup(*, workspace_root: Path, max_age: timedelta, apply: bool, now: datetime | None = None) -> dict[str, Any]:
    """Return the legacy lease preview without mutating the workspace.

    Direct mutation used to be available through ``apply=True``. Housekeeping now
    requires a saved, fingerprinted plan, so callers must use
    :func:`apply_saved_plan` instead. The preview shape remains compatible with
    older QA and operator checks.
    """

    if apply:
        raise HousekeepingError("direct_apply_disabled_use_saved_plan")
    workspace = _absolute(workspace_root)
    temp_root = workspace / "temp"
    result: dict[str, Any] = {
        "schemaVersion": 1,
        "workspaceRoot": str(workspace),
        "tempRoot": str(temp_root),
        "mode": "dry-run",
        "eligible": [],
        "removed": [],
        "skipped": [],
    }
    if not temp_root.exists():
        return result
    if not temp_root.is_dir() or not _chain_is_plain(temp_root, boundary=workspace):
        raise ValueError(f"unsafe_temp_root:{temp_root}")
    current_time = (now or datetime.now(UTC)).astimezone(UTC)
    scan_observations: list[dict[str, str]] = []
    markers = _lease_markers(temp_root, errors=scan_observations)
    result["skipped"].extend(scan_observations)
    for marker in markers:
        try:
            runtime_root, payload = _read_lease(marker, temp_root=temp_root)
            age = current_time - _created_at(payload)
            if age < max_age:
                result["skipped"].append({"path": str(runtime_root), "reason": "not_expired"})
                continue
            if _process_alive(payload.get("pid")):
                result["skipped"].append({"path": str(runtime_root), "reason": "owner_process_alive"})
                continue
            result["eligible"].append(str(runtime_root))
        except (OSError, ValueError, json.JSONDecodeError) as error:
            result["skipped"].append({"path": str(marker.parent), "reason": str(error)})
    return result


def retention_decision(
    *,
    state: object,
    age: timedelta,
    same_kind_rank: int,
    failure_ttl: timedelta = DEFAULT_FAILURE_TTL,
    failure_max_per_kind: int = DEFAULT_FAILURE_MAX_PER_KIND,
) -> dict[str, object]:
    """Return the deterministic retention decision for one dead leased root."""

    if failure_ttl < timedelta(0):
        raise HousekeepingError("failure_ttl_must_be_non_negative")
    if failure_max_per_kind < 0:
        raise HousekeepingError("failure_max_per_kind_must_be_non_negative")
    if same_kind_rank < 0:
        raise HousekeepingError("same_kind_rank_must_be_non_negative")
    normalized_state = str(state or "").strip().casefold()
    if normalized_state in SUCCESS_STATES:
        return {"eligible": True, "reason": "success_scratch"}
    is_failure = normalized_state in FAILURE_STATES
    if is_failure:
        if age >= failure_ttl:
            return {"eligible": True, "reason": "failure_ttl_expired"}
        if same_kind_rank >= failure_max_per_kind:
            return {"eligible": True, "reason": "failure_retention_cap"}
        return {"eligible": False, "reason": "failure_retained"}
    return {"eligible": False, "reason": "state_not_terminal"}


def _resolve_reparse_target(path: Path) -> Path:
    try:
        return _absolute(path.resolve(strict=True))
    except (OSError, RuntimeError) as error:
        raise HousekeepingError(f"unresolved_reparse_target:{path}:{error}") from error


def _snapshot_tree(root: Path) -> dict[str, Any]:
    absolute_root = _absolute(root)
    root_metadata = os.lstat(_windows_extended_path(absolute_root))
    if not stat.S_ISDIR(root_metadata.st_mode) or _is_link_or_reparse(root_metadata):
        raise HousekeepingError(f"plain_directory_required:{absolute_root}")
    digest = hashlib.sha256()
    content_digest = hashlib.sha256()
    total_bytes = 0
    file_count = 0
    directory_count = 1
    latest_mtime_ns = root_metadata.st_mtime_ns
    reparse_points: list[dict[str, str]] = []
    lease_marker_paths: list[Path] = []
    lease_roots: list[Path] = []

    def record(kind: str, relative: str, metadata: os.stat_result, extra: str = "") -> None:
        nonlocal latest_mtime_ns
        latest_mtime_ns = max(latest_mtime_ns, metadata.st_mtime_ns)
        digest.update(
            _canonical_json(
                {
                    "kind": kind,
                    "path": relative.replace(os.sep, "/"),
                    "device": metadata.st_dev,
                    "inode": metadata.st_ino,
                    "mode": stat.S_IFMT(metadata.st_mode),
                    "size": metadata.st_size,
                    "mtimeNs": metadata.st_mtime_ns,
                    "extra": extra,
                }
            )
        )

    record("directory", ".", root_metadata)
    for current, directories, files in os.walk(
        _windows_extended_path(absolute_root),
        topdown=True,
        followlinks=False,
        onerror=_raise_walk_error,
    ):
        current_path = _logical_path(current)
        safe_directories: list[str] = []
        for name in sorted(directories):
            child = current_path / name
            metadata = os.lstat(_windows_extended_path(child))
            relative = str(child.relative_to(absolute_root))
            if _is_link_or_reparse(metadata):
                target = _resolve_reparse_target(child)
                if not _is_within(target, absolute_root):
                    raise HousekeepingError(f"reparse_target_outside_target:{child}:{target}")
                reparse_points.append({"path": relative, "target": str(target)})
                record("reparse-directory", relative, metadata, str(target))
                continue
            if not stat.S_ISDIR(metadata.st_mode):
                raise HousekeepingError(f"unexpected_directory_entry:{child}")
            directory_count += 1
            safe_directories.append(name)
            record("directory", relative, metadata)
        directories[:] = safe_directories
        for name in sorted(files):
            child = current_path / name
            metadata = os.lstat(_windows_extended_path(child))
            relative = str(child.relative_to(absolute_root))
            if _is_link_or_reparse(metadata):
                target = _resolve_reparse_target(child)
                if not _is_within(target, absolute_root):
                    raise HousekeepingError(f"reparse_target_outside_target:{child}:{target}")
                reparse_points.append({"path": relative, "target": str(target)})
                record("reparse-file", relative, metadata, str(target))
                continue
            if not stat.S_ISREG(metadata.st_mode):
                raise HousekeepingError(f"unsupported_special_file:{child}")
            file_count += 1
            total_bytes += metadata.st_size
            record("file", relative, metadata)
            content_digest.update(relative.replace(os.sep, "/").encode("utf-8"))
            content_digest.update(b"\0")
            with open(_windows_extended_path(child), "rb") as stream:
                while chunk := stream.read(1024 * 1024):
                    content_digest.update(chunk)
            content_digest.update(b"\0")
            after_read = os.lstat(_windows_extended_path(child))
            if (
                after_read.st_dev,
                after_read.st_ino,
                after_read.st_mode,
                after_read.st_size,
                after_read.st_mtime_ns,
            ) != (
                metadata.st_dev,
                metadata.st_ino,
                metadata.st_mode,
                metadata.st_size,
                metadata.st_mtime_ns,
            ):
                raise HousekeepingError(f"target_changed_during_snapshot:{child}")
            if name == LEASE_NAME:
                inside_leased_root = any(
                    lease_root == current_path or lease_root in current_path.parents for lease_root in lease_roots
                )
                if not inside_leased_root:
                    lease_marker_paths.append(child)
                    lease_roots.append(current_path)

    after_root = os.lstat(_windows_extended_path(absolute_root))
    if (
        after_root.st_dev,
        after_root.st_ino,
        after_root.st_mode,
        after_root.st_mtime_ns,
    ) != (
        root_metadata.st_dev,
        root_metadata.st_ino,
        root_metadata.st_mode,
        root_metadata.st_mtime_ns,
    ):
        raise HousekeepingError(f"target_changed_during_snapshot:{absolute_root}")

    leases = [_snapshot_lease(marker, root=absolute_root) for marker in sorted(lease_marker_paths, key=str)]
    return {
        "bytes": total_bytes,
        "files": file_count,
        "directories": directory_count,
        "rootDevice": root_metadata.st_dev,
        "rootInode": root_metadata.st_ino,
        "rootMtimeNs": root_metadata.st_mtime_ns,
        "latestMtimeNs": latest_mtime_ns,
        "metadataSha256": digest.hexdigest(),
        "contentSha256": content_digest.hexdigest(),
        "reparsePoints": sorted(reparse_points, key=lambda item: item["path"]),
        "leases": leases,
    }


def _snapshot_lease(marker: Path, *, root: Path) -> dict[str, Any]:
    payload_bytes = marker.read_bytes()
    payload = json.loads(payload_bytes.decode("utf-8"))
    if not isinstance(payload, dict) or payload.get("schemaVersion") != 1:
        raise HousekeepingError(f"invalid_lease_schema:{marker}")
    owner = payload.get("owner")
    if not isinstance(owner, str) or not owner.startswith("docwen."):
        raise HousekeepingError(f"foreign_lease_owner:{marker}")
    lease_root = marker.parent.resolve(strict=True)
    if payload.get("root") != str(lease_root):
        raise HousekeepingError(f"lease_root_mismatch:{marker}")
    if not _is_within(lease_root, root):
        raise HousekeepingError(f"lease_outside_target:{marker}")
    return {
        "path": str(marker.relative_to(root)).replace(os.sep, "/"),
        "sha256": _sha256_bytes(payload_bytes),
        "owner": owner,
        "kind": payload.get("kind"),
        "pid": payload.get("pid"),
        "createdAt": payload.get("createdAt"),
        "state": payload.get("state"),
        "root": payload.get("root"),
        "shortDrive": payload.get("shortDrive"),
    }


def _managed_roots(workspace: Path) -> tuple[Path, ...]:
    return tuple(workspace / name for name in MANAGED_ROOT_NAMES)


def _protected_roots(workspace: Path) -> tuple[Path, ...]:
    return tuple(workspace / name for name in PROTECTED_WORKSPACE_ROOT_NAMES)


def _classify_target(
    target: Path,
    *,
    workspace: Path,
    clean_dependencies: bool,
) -> tuple[str, Path]:
    absolute_target = _absolute(target)
    engineering_root = workspace.parent
    if clean_dependencies and absolute_target.name not in DEPENDENCY_DIRECTORY_NAMES:
        raise HousekeepingError(f"clean_deps_name_required:{absolute_target}")
    if not _is_within(absolute_target, engineering_root):
        raise HousekeepingError(f"target_outside_engineering_root:{absolute_target}")
    for protected in _protected_roots(workspace):
        if _is_within(absolute_target, protected):
            raise HousekeepingError(f"protected_target:{absolute_target}")
    repository_root = engineering_root / "repos"
    if _is_within(absolute_target, repository_root):
        if not clean_dependencies or absolute_target.name not in DEPENDENCY_DIRECTORY_NAMES:
            raise HousekeepingError(f"repository_target_requires_clean_deps:{absolute_target}")
        if absolute_target == repository_root:
            raise HousekeepingError(f"repository_root_forbidden:{absolute_target}")
        return "clean-deps", repository_root
    for managed_root in _managed_roots(workspace):
        if _is_within(absolute_target, managed_root):
            if absolute_target == managed_root:
                raise HousekeepingError(f"managed_root_forbidden:{absolute_target}")
            if absolute_target.name in DEPENDENCY_DIRECTORY_NAMES:
                if not clean_dependencies:
                    raise HousekeepingError(f"dependency_target_requires_clean_deps:{absolute_target}")
                return "clean-deps", managed_root
            return "explicit", managed_root
    if absolute_target.parent == engineering_root and absolute_target.name.casefold().startswith("temp-"):
        return "explicit-bypass", engineering_root
    raise HousekeepingError(f"target_not_in_allowed_root:{absolute_target}")


def _validate_target_chain(target: Path, *, boundary: Path) -> None:
    if not target.exists():
        raise HousekeepingError(f"target_missing:{target}")
    if not _chain_is_plain(target, boundary=boundary):
        raise HousekeepingError(f"target_chain_linked_or_reparse:{target}")


def _entry_for_target(
    target: Path,
    *,
    workspace: Path,
    reason: str,
    source: str | None = None,
    clean_dependencies: bool = False,
) -> dict[str, Any]:
    normalized_reason = reason.strip()
    if not normalized_reason:
        raise HousekeepingError("target_reason_required")
    classified_source, boundary = _classify_target(
        target,
        workspace=workspace,
        clean_dependencies=clean_dependencies,
    )
    if source is None:
        source = classified_source
    absolute_target = _absolute(target)
    _validate_target_chain(absolute_target, boundary=boundary)
    identity = _snapshot_tree(absolute_target)
    for lease in identity["leases"]:
        if _process_alive(lease.get("pid")):
            raise HousekeepingError(f"lease_process_alive:{absolute_target}:{lease.get('pid')}")
    return {
        "path": str(absolute_target),
        "source": source,
        "reason": normalized_reason,
        "allowedBoundary": str(_absolute(boundary)),
        "identity": identity,
    }


def _assert_non_overlapping(entries: Sequence[dict[str, Any]]) -> None:
    paths = sorted((_absolute(Path(str(entry["path"]))) for entry in entries), key=lambda path: len(path.parts))
    for index, path in enumerate(paths):
        for parent in paths[:index]:
            if _is_within(path, parent):
                raise HousekeepingError(f"overlapping_targets:{parent}:{path}")


def _bypass_observations(workspace: Path) -> list[dict[str, Any]]:
    engineering_root = workspace.parent
    observations: list[dict[str, Any]] = []
    for path in sorted(engineering_root.glob("temp-*"), key=lambda item: os.path.normcase(str(item))):
        try:
            metadata = path.lstat()
            observations.append(
                {
                    "path": str(_absolute(path)),
                    "isDirectory": stat.S_ISDIR(metadata.st_mode),
                    "isReparse": _is_link_or_reparse(metadata),
                    "rootMtimeNs": metadata.st_mtime_ns,
                    "reportedOnly": True,
                }
            )
        except OSError as error:
            observations.append({"path": str(_absolute(path)), "error": str(error), "reportedOnly": True})
    return observations


def _automatic_lease_candidates(
    *,
    workspace: Path,
    current_time: datetime,
    failure_ttl: timedelta,
    failure_max_per_kind: int,
) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    discovered: list[tuple[Path, dict[str, Any], Path]] = []
    skipped: list[dict[str, str]] = []
    for managed_root in _managed_roots(workspace):
        if not managed_root.exists():
            continue
        if not managed_root.is_dir() or not _chain_is_plain(managed_root, boundary=workspace):
            raise HousekeepingError(f"unsafe_managed_root:{managed_root}")
        scan_errors: list[dict[str, str]] = []
        markers = _lease_markers(managed_root, errors=scan_errors)
        skipped.extend(scan_errors)
        for marker in markers:
            try:
                leased_root, payload = _read_lease(marker, temp_root=managed_root)
                discovered.append((leased_root, payload, managed_root))
            except (OSError, ValueError, json.JSONDecodeError) as error:
                skipped.append({"path": str(marker.parent), "reason": str(error)})

    failure_groups: dict[str, list[tuple[Path, dict[str, Any], Path]]] = {}
    for candidate in discovered:
        state = str(candidate[1].get("state") or "").strip().casefold()
        if state in FAILURE_STATES:
            kind = str(candidate[1].get("kind") or candidate[1].get("owner") or "unknown")
            failure_groups.setdefault(kind, []).append(candidate)
    failure_rank: dict[str, int] = {}
    for candidates in failure_groups.values():
        candidates.sort(key=lambda item: _created_at(item[1]), reverse=True)
        for rank, (path, _, _) in enumerate(candidates):
            failure_rank[os.path.normcase(str(path))] = rank

    entries: list[dict[str, Any]] = []
    for leased_root, payload, managed_root in discovered:
        if _process_alive(payload.get("pid")):
            skipped.append({"path": str(leased_root), "reason": "owner_process_alive"})
            continue
        if leased_root.name in DEPENDENCY_DIRECTORY_NAMES:
            skipped.append({"path": str(leased_root), "reason": "dependency_target_requires_clean_deps"})
            continue
        age = current_time - _created_at(payload)
        decision = retention_decision(
            state=payload.get("state"),
            age=age,
            same_kind_rank=failure_rank.get(os.path.normcase(str(leased_root)), 0),
            failure_ttl=failure_ttl,
            failure_max_per_kind=failure_max_per_kind,
        )
        if not decision["eligible"]:
            skipped.append({"path": str(leased_root), "reason": str(decision["reason"])})
            continue
        entry = _entry_for_target(
            leased_root,
            workspace=workspace,
            reason=str(decision["reason"]),
            source="lease-policy",
        )
        entry["allowedBoundary"] = str(_absolute(managed_root))
        entries.append(entry)
    return entries, skipped


def _plan_without_fingerprint(plan: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in plan.items() if key != "planFingerprint"}


def _plan_fingerprint(plan: dict[str, Any]) -> str:
    return _sha256_bytes(_canonical_json(_plan_without_fingerprint(plan)))


def create_plan(
    *,
    workspace_root: Path,
    explicit_targets: Sequence[Path] = (),
    reason: str | None = None,
    clean_dependency_targets: Sequence[Path] = (),
    now: datetime | None = None,
    failure_ttl: timedelta = DEFAULT_FAILURE_TTL,
    failure_max_per_kind: int = DEFAULT_FAILURE_MAX_PER_KIND,
) -> dict[str, Any]:
    workspace = _absolute(workspace_root)
    if not workspace.is_dir() or not _chain_is_plain(workspace, boundary=workspace):
        raise HousekeepingError(f"unsafe_workspace_root:{workspace}")
    current_time = (now or datetime.now(UTC)).astimezone(UTC)
    entries: list[dict[str, Any]] = []
    skipped: list[dict[str, str]] = []
    if explicit_targets or clean_dependency_targets:
        if explicit_targets and (reason is None or not reason.strip()):
            raise HousekeepingError("explicit_target_reason_required")
        entries.extend(
            _entry_for_target(target, workspace=workspace, reason=str(reason), clean_dependencies=False)
            for target in explicit_targets
        )
        entries.extend(
            _entry_for_target(
                target,
                workspace=workspace,
                reason="explicit_clean_deps",
                clean_dependencies=True,
            )
            for target in clean_dependency_targets
        )
    else:
        automatic_entries, skipped = _automatic_lease_candidates(
            workspace=workspace,
            current_time=current_time,
            failure_ttl=failure_ttl,
            failure_max_per_kind=failure_max_per_kind,
        )
        entries.extend(automatic_entries)
    entries.sort(key=lambda entry: os.path.normcase(str(entry["path"])))
    _assert_non_overlapping(entries)
    plan: dict[str, Any] = {
        "schema": PLAN_SCHEMA,
        "createdAt": current_time.isoformat().replace("+00:00", "Z"),
        "workspaceRoot": str(workspace),
        "engineeringRoot": str(workspace.parent),
        "policy": {
            "managedRoots": [str(path) for path in _managed_roots(workspace)],
            "protectedRoots": [str(path) for path in _protected_roots(workspace)] + [str(workspace.parent / "repos")],
            "failureTtlSeconds": int(failure_ttl.total_seconds()),
            "failureMaxPerKind": failure_max_per_kind,
            "dependencyDirectories": sorted(DEPENDENCY_DIRECTORY_NAMES),
        },
        "entries": entries,
        "observations": {
            "rootTempBypasses": _bypass_observations(workspace),
            "skipped": sorted(skipped, key=lambda item: os.path.normcase(item["path"])),
        },
    }
    plan["planFingerprint"] = _plan_fingerprint(plan)
    return plan


def _diagnostics_root(workspace: Path) -> Path:
    return workspace / "diagnostics"


def _validate_plan_path(plan_path: Path, *, workspace: Path, must_exist: bool) -> Path:
    absolute_path = _absolute(plan_path)
    diagnostics = _diagnostics_root(workspace)
    if not _is_within(absolute_path, diagnostics) or absolute_path == diagnostics:
        raise HousekeepingError(f"plan_path_must_be_under_diagnostics:{absolute_path}")
    parent = absolute_path.parent
    if not parent.exists() or not _chain_is_plain(parent, boundary=diagnostics):
        raise HousekeepingError(f"unsafe_plan_parent:{parent}")
    if must_exist:
        metadata = absolute_path.lstat()
        if not stat.S_ISREG(metadata.st_mode) or _is_link_or_reparse(metadata):
            raise HousekeepingError(f"plain_plan_file_required:{absolute_path}")
    return absolute_path


def save_plan(plan: dict[str, Any], plan_path: Path) -> Path:
    workspace = _absolute(Path(str(plan.get("workspaceRoot", ""))))
    if plan.get("schema") != PLAN_SCHEMA or plan.get("planFingerprint") != _plan_fingerprint(plan):
        raise HousekeepingError("invalid_plan_before_save")
    destination = _validate_plan_path(plan_path, workspace=workspace, must_exist=False)
    try:
        with destination.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(plan, stream, ensure_ascii=False, indent=2, sort_keys=True)
            stream.write("\n")
    except FileExistsError as error:
        raise HousekeepingError(f"plan_path_exists:{destination}") from error
    return destination


def load_plan(plan_path: Path, *, workspace_root: Path | None = None) -> dict[str, Any]:
    candidate = _absolute(plan_path)
    if workspace_root is not None:
        _validate_plan_path(candidate, workspace=_absolute(workspace_root), must_exist=True)
    else:
        metadata = candidate.lstat()
        if not stat.S_ISREG(metadata.st_mode) or _is_link_or_reparse(metadata):
            raise HousekeepingError(f"plain_plan_file_required:{candidate}")
    payload = _read_json_object(candidate)
    if payload.get("schema") != PLAN_SCHEMA:
        raise HousekeepingError("unsupported_plan_schema")
    workspace = _absolute(Path(str(payload.get("workspaceRoot", ""))))
    if workspace_root is not None and workspace != _absolute(workspace_root):
        raise HousekeepingError(f"plan_workspace_mismatch:{workspace}")
    _validate_plan_path(candidate, workspace=workspace, must_exist=True)
    if payload.get("planFingerprint") != _plan_fingerprint(payload):
        raise HousekeepingError("plan_fingerprint_mismatch")
    return payload


def _revalidate_entry(entry: dict[str, Any], *, workspace: Path) -> dict[str, Any]:
    path = _absolute(Path(str(entry.get("path", ""))))
    source = str(entry.get("source", ""))
    clean_dependencies = source == "clean-deps"
    classified_source, boundary = _classify_target(
        path,
        workspace=workspace,
        clean_dependencies=clean_dependencies,
    )
    source_matches = source == classified_source or (source == "lease-policy" and classified_source == "explicit")
    if not source_matches:
        raise HousekeepingError(f"plan_source_mismatch:{path}:{source}:{classified_source}")
    planned_boundary = _absolute(Path(str(entry.get("allowedBoundary", ""))))
    if planned_boundary != _absolute(boundary):
        raise HousekeepingError(f"allowed_boundary_mismatch:{path}")
    _validate_target_chain(path, boundary=boundary)
    current_identity = _snapshot_tree(path)
    if current_identity != entry.get("identity"):
        raise HousekeepingError(f"target_identity_changed:{path}")
    for lease in current_identity["leases"]:
        if _process_alive(lease.get("pid")):
            raise HousekeepingError(f"lease_process_alive:{path}:{lease.get('pid')}")
    return current_identity


def _unmount_entry_short_drives(entry: dict[str, Any]) -> None:
    path = _absolute(Path(str(entry["path"])))
    for lease in entry["identity"]["leases"]:
        short_drive = lease.get("shortDrive")
        if short_drive is None:
            continue
        if not isinstance(short_drive, str):
            raise HousekeepingError(f"invalid_short_drive:{path}")
        lease_root = path / Path(str(lease["path"])).parent
        try:
            unmount_short_drive(short_drive, expected_target=lease_root)
        except ShortPathDriveError as error:
            raise HousekeepingError(str(error)) from error


def apply_saved_plan(plan_path: Path, *, workspace_root: Path) -> dict[str, Any]:
    """Apply one saved plan after a full preflight and per-target revalidation."""

    plan = load_plan(plan_path, workspace_root=workspace_root)
    workspace = _absolute(Path(str(plan["workspaceRoot"])))
    entries = plan.get("entries")
    if not isinstance(entries, list):
        raise HousekeepingError("plan_entries_must_be_list")
    _assert_non_overlapping(entries)
    for entry in entries:
        if not isinstance(entry, dict):
            raise HousekeepingError("plan_entry_must_be_object")
        _revalidate_entry(entry, workspace=workspace)
    removed: list[str] = []
    removed_entries: list[dict[str, object]] = []
    for entry in entries:
        _revalidate_entry(entry, workspace=workspace)
        _unmount_entry_short_drives(entry)
        path = _absolute(Path(str(entry["path"])))
        shutil.rmtree(_windows_extended_path(path), onexc=_remove_owned_readonly_path)
        removed.append(str(path))
        removed_entries.append({"path": str(path), "bytes": int(entry["identity"]["bytes"])})
    return {
        "schema": "docwen.housekeeping-apply-result.v1",
        "planPath": str(_absolute(plan_path)),
        "planFingerprint": plan["planFingerprint"],
        "removed": removed,
        "removedEntries": removed_entries,
        "removedBytes": sum(int(entry["bytes"]) for entry in removed_entries),
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create or apply a fail-closed DocWen housekeeping plan.")
    parser.add_argument("--workspace-root", type=Path)
    parser.add_argument("--target", action="append", type=Path, default=[])
    parser.add_argument("--reason")
    parser.add_argument(
        "--clean-deps",
        action="append",
        type=Path,
        default=[],
        metavar="PATH",
        help="Explicitly plan one .venv or node_modules directory.",
    )
    parser.add_argument("--failure-ttl-hours", type=float, default=72.0)
    parser.add_argument(
        "--max-age-hours",
        dest="failure_ttl_hours",
        type=float,
        default=argparse.SUPPRESS,
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--failure-max-per-kind", type=int, default=2)
    parser.add_argument("--plan-output", type=Path)
    parser.add_argument("--apply-plan", type=Path)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.failure_ttl_hours < 0:
        print("workspace_cleanup_error:failure_ttl_hours_must_be_non_negative")
        return 2
    if args.failure_max_per_kind < 0:
        print("workspace_cleanup_error:failure_max_per_kind_must_be_non_negative")
        return 2
    if args.apply_plan is not None and (
        args.target or args.clean_deps or args.reason is not None or args.plan_output is not None
    ):
        print("workspace_cleanup_error:apply_plan_is_mutually_exclusive_with_plan_generation")
        return 2
    try:
        workspace_root = resolve_workspace_root(
            Path(__file__).resolve().parents[1],
            explicit=args.workspace_root,
        )
        if args.apply_plan is not None:
            result = apply_saved_plan(args.apply_plan, workspace_root=workspace_root)
        else:
            result = create_plan(
                workspace_root=workspace_root,
                explicit_targets=tuple(args.target),
                reason=args.reason,
                clean_dependency_targets=tuple(args.clean_deps),
                failure_ttl=timedelta(hours=args.failure_ttl_hours),
                failure_max_per_kind=args.failure_max_per_kind,
            )
            if args.plan_output is not None:
                saved = save_plan(result, args.plan_output)
                result = {**result, "savedPlanPath": str(saved)}
    except (OSError, HousekeepingError, ValueError, json.JSONDecodeError) as error:
        print(f"workspace_cleanup_error:{error}")
        return 2
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
