from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path

_DRIVE_PATTERN = re.compile(r"^[A-Z]:$")
_DRIVE_LETTERS = tuple(f"{letter}:" for letter in "ZYXWVUTSRQP")


class ShortPathDriveError(RuntimeError):
    """A temporary Windows drive could not be managed safely."""


def _is_windows() -> bool:
    return os.name == "nt"


def drive_root(drive: str) -> Path:
    normalized = drive.strip().upper()
    if _DRIVE_PATTERN.fullmatch(normalized) is None:
        raise ShortPathDriveError(f"invalid_short_drive:{drive}")
    return Path(f"{normalized}\\")


def _run_subst(*arguments: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["subst.exe", *arguments],
        check=False,
        capture_output=True,
        text=True,
        timeout=10,
    )


def _same_directory(first: Path, second: Path) -> bool:
    try:
        return os.path.samefile(first, second)
    except OSError:
        return False


def _drive_is_available(drive: str) -> bool:
    return not drive_root(drive).exists()


def mount_short_drive(target: Path) -> str:
    """Map one unused drive letter to an existing physical runtime root."""
    if not _is_windows():
        raise ShortPathDriveError("short_drive_requires_windows")

    physical = target.resolve(strict=True)
    if not physical.is_dir():
        raise ShortPathDriveError(f"short_drive_target_not_directory:{physical}")

    failures: list[str] = []
    for drive in _DRIVE_LETTERS:
        root = drive_root(drive)
        if not _drive_is_available(drive):
            continue
        try:
            result = _run_subst(drive, str(physical))
        except (OSError, subprocess.SubprocessError) as error:
            failures.append(f"{drive}:{error}")
            continue
        if result.returncode != 0:
            detail = (result.stderr or result.stdout).strip()
            failures.append(f"{drive}:{detail or result.returncode}")
            continue
        if _same_directory(root, physical):
            return drive
        rollback = _run_subst(drive, "/D")
        failures.append(f"{drive}:mapping_identity_mismatch:{rollback.returncode}")

    detail = ";".join(failures[-3:])
    raise ShortPathDriveError(f"no_safe_short_drive_available:{detail}")


def unmount_short_drive(drive: str, *, expected_target: Path) -> None:
    """Remove only a drive mapping that still resolves to the expected target."""
    if not _is_windows():
        raise ShortPathDriveError("short_drive_requires_windows")

    root = drive_root(drive)
    physical = expected_target.resolve(strict=True)
    if not root.exists():
        return
    if not _same_directory(root, physical):
        raise ShortPathDriveError(f"short_drive_target_mismatch:{drive}:{physical}")
    try:
        result = _run_subst(drive.strip().upper(), "/D")
    except (OSError, subprocess.SubprocessError) as error:
        raise ShortPathDriveError(f"short_drive_unmount_failed:{drive}:{error}") from error
    if result.returncode != 0:
        detail = (result.stderr or result.stdout).strip()
        raise ShortPathDriveError(f"short_drive_unmount_failed:{drive}:{detail or result.returncode}")
