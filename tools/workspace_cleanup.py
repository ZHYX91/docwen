from __future__ import annotations

import argparse
import json
import os
import shutil
import stat
import sys
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

if __package__ in {None, ""}:
    _BOOTSTRAP_ROOT = Path(__file__).resolve().parents[1]
    if str(_BOOTSTRAP_ROOT) not in sys.path:
        sys.path.insert(0, str(_BOOTSTRAP_ROOT))

from tools.windows_short_path import ShortPathDriveError, unmount_short_drive
from tools.workspace_root import WORKSPACE_ROOT_ENV as _WORKSPACE_ROOT_ENV
from tools.workspace_root import resolve_workspace_root

LEASE_NAME = ".docwen-temp-lease.json"
WORKSPACE_ROOT_ENV = _WORKSPACE_ROOT_ENV


def _absolute(path: Path) -> Path:
    return Path(os.path.abspath(os.fspath(path)))


def _is_reparse(metadata: os.stat_result) -> bool:
    flag = int(getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0))
    return bool(flag and int(getattr(metadata, "st_file_attributes", 0)) & flag)


def _chain_is_plain(path: Path, *, boundary: Path) -> bool:
    current = _absolute(path)
    stop = _absolute(boundary)
    while True:
        metadata = current.lstat()
        if stat.S_ISLNK(metadata.st_mode) or _is_reparse(metadata):
            return False
        if current == stop:
            return True
        parent = current.parent
        if parent == current:
            return False
        current = parent


def _read_lease(marker: Path, *, temp_root: Path) -> tuple[Path, dict[str, Any]]:
    runtime_root = marker.parent
    if not _chain_is_plain(runtime_root, boundary=temp_root):
        raise ValueError("linked_or_reparse_path")
    payload = json.loads(marker.read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or payload.get("schemaVersion") != 1:
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


def _lease_markers(temp_root: Path) -> list[Path]:
    markers: list[Path] = []
    for current, directories, files in os.walk(temp_root, topdown=True, followlinks=False):
        current_path = Path(current)
        safe_directories: list[str] = []
        for name in directories:
            child = current_path / name
            try:
                metadata = child.lstat()
            except OSError:
                continue
            if not stat.S_ISLNK(metadata.st_mode) and not _is_reparse(metadata):
                safe_directories.append(name)
        directories[:] = safe_directories
        if LEASE_NAME in files:
            markers.append(current_path / LEASE_NAME)
            directories[:] = []
    return sorted(markers, key=str)


def _windows_extended_path(path: Path) -> str:
    absolute = os.path.abspath(os.fspath(path))
    if os.name != "nt" or absolute.startswith("\\\\?\\"):
        return absolute
    if absolute.startswith("\\\\"):
        return f"\\\\?\\UNC\\{absolute[2:]}"
    return f"\\\\?\\{absolute}"


def _remove_owned_readonly_path(function: object, path: str, error: BaseException) -> None:
    if not isinstance(error, PermissionError) or not callable(function):
        raise error
    os.chmod(path, stat.S_IWRITE)
    function(path)


def cleanup(*, workspace_root: Path, max_age: timedelta, apply: bool, now: datetime | None = None) -> dict[str, Any]:
    workspace = _absolute(workspace_root)
    temp_root = workspace / "temp"
    result: dict[str, Any] = {
        "schemaVersion": 1,
        "workspaceRoot": str(workspace),
        "tempRoot": str(temp_root),
        "mode": "apply" if apply else "dry-run",
        "eligible": [],
        "removed": [],
        "skipped": [],
    }
    if not temp_root.exists():
        return result
    if not temp_root.is_dir() or not _chain_is_plain(temp_root, boundary=workspace):
        raise ValueError(f"unsafe_temp_root:{temp_root}")
    current_time = (now or datetime.now(UTC)).astimezone(UTC)
    for marker in _lease_markers(temp_root):
        try:
            runtime_root, payload = _read_lease(marker, temp_root=temp_root)
            age = current_time - _created_at(payload)
            if age < max_age:
                result["skipped"].append({"path": str(runtime_root), "reason": "not_expired"})
                continue
            if _process_alive(payload.get("pid")):
                result["skipped"].append({"path": str(runtime_root), "reason": "owner_process_alive"})
                continue
            short_drive = payload.get("shortDrive")
            if short_drive is not None:
                if not isinstance(short_drive, str):
                    raise ValueError("invalid_short_drive")
                if apply:
                    try:
                        unmount_short_drive(short_drive, expected_target=runtime_root)
                    except ShortPathDriveError as error:
                        raise ValueError(str(error)) from error
            result["eligible"].append(str(runtime_root))
            if apply:
                shutil.rmtree(_windows_extended_path(runtime_root), onexc=_remove_owned_readonly_path)
                result["removed"].append(str(runtime_root))
        except (OSError, ValueError, json.JSONDecodeError) as error:
            result["skipped"].append({"path": str(marker.parent), "reason": str(error)})
    return result


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Safely preview or remove expired DocWen-owned temporary roots.")
    parser.add_argument("--workspace-root", type=Path)
    parser.add_argument("--max-age-hours", type=float, default=72.0)
    parser.add_argument("--apply", action="store_true", help="Remove eligible roots; the default is dry-run.")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.max_age_hours < 0:
        print("workspace_cleanup_error:max_age_hours_must_be_non_negative")
        return 2
    try:
        workspace_root = resolve_workspace_root(
            Path(__file__).resolve().parents[1],
            explicit=args.workspace_root,
        )
        result = cleanup(
            workspace_root=workspace_root,
            max_age=timedelta(hours=args.max_age_hours),
            apply=args.apply,
        )
    except (OSError, ValueError) as error:
        print(f"workspace_cleanup_error:{error}")
        return 2
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
