"""Cross-platform open and reveal actions for GUI-owned paths."""

from __future__ import annotations

import logging
import subprocess
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices

from docwen_runtime.path_io import filesystem_path

logger = logging.getLogger(__name__)

_COMMAND_TIMEOUT_SECONDS = 5


@dataclass(frozen=True, slots=True)
class PathActionResult:
    """Describe whether a path action was exact or used a safe fallback."""

    success: bool
    error: str | None = None
    error_code: str | None = None
    precise: bool = False
    fallback_used: bool = False


def _success(*, precise: bool = False, fallback_used: bool = False) -> PathActionResult:
    return PathActionResult(success=True, precise=precise, fallback_used=fallback_used)


def _failure(error: str, *, error_code: str) -> PathActionResult:
    return PathActionResult(success=False, error=error, error_code=error_code)


def _platform_key() -> str:
    if sys.platform.startswith("win"):
        return "windows"
    if sys.platform == "darwin":
        return "macos"
    if sys.platform.startswith("linux"):
        return "linux"
    return "other"


def _coerce_path(target_path: str | Path) -> tuple[Path | None, PathActionResult | None]:
    try:
        return Path(target_path).expanduser(), None
    except Exception as exc:
        return None, _failure(str(exc), error_code="probe_failed")


def _run_command(command: Sequence[str]) -> PathActionResult:
    try:
        completed = subprocess.run(
            list(command),
            capture_output=True,
            check=False,
            text=True,
            timeout=_COMMAND_TIMEOUT_SECONDS,
        )
    except FileNotFoundError as exc:
        return _failure(str(exc), error_code="command_missing")
    except Exception as exc:
        return _failure(str(exc), error_code="launch_failed")

    if completed.returncode == 0:
        return _success()
    detail = (completed.stderr or completed.stdout or "").strip()
    if not detail:
        detail = f"command exited with code {completed.returncode}"
    return _failure(detail, error_code="command_failed")


def _open_with_desktop_services(target_path: Path) -> PathActionResult:
    try:
        if QDesktopServices.openUrl(QUrl.fromLocalFile(str(target_path))):
            return _success()
        return _failure(str(target_path), error_code="open_failed")
    except Exception as exc:
        return _failure(str(exc), error_code="launch_failed")


def linux_reveal_commands(target_path: Path) -> list[list[str]]:
    """Return ordered Linux selectors from desktop-neutral to desktop-specific."""
    file_uri = target_path.resolve(strict=False).as_uri()
    return [
        [
            "dbus-send",
            "--session",
            "--dest=org.freedesktop.FileManager1",
            "--type=method_call",
            "/org/freedesktop/FileManager1",
            "org.freedesktop.FileManager1.ShowItems",
            f"array:string:{file_uri}",
            "string:",
        ],
        ["nautilus", "--select", str(target_path)],
        ["dolphin", "--select", str(target_path)],
    ]


def open_path(target_path: str | Path) -> PathActionResult:
    """Open a file/directory, falling back to an existing parent when needed."""
    candidate, error = _coerce_path(target_path)
    if error is not None or candidate is None:
        return error or _failure("invalid path", error_code="probe_failed")

    try:
        io_candidate = filesystem_path(candidate)
        candidate_exists = io_candidate.exists()
        candidate_is_file = io_candidate.is_file() if candidate_exists else False
    except Exception as exc:
        return _failure(str(exc), error_code="probe_failed")

    if candidate_exists:
        opened = _open_with_desktop_services(candidate)
        if opened.success:
            logger.info("Opened path: %s", candidate)
            return opened
        if not candidate_is_file:
            return opened

    try:
        parent = candidate.parent
        if parent != candidate and filesystem_path(parent).exists():
            opened = _open_with_desktop_services(parent)
            if opened.success:
                logger.info("Opened parent fallback: %s", parent)
                return _success(fallback_used=True)
            return opened
    except Exception as exc:
        return _failure(str(exc), error_code="probe_failed")
    return _failure(str(candidate), error_code="missing_path" if not candidate_exists else "open_failed")


def reveal_path(target_path: str | Path) -> PathActionResult:
    """Select a file in its manager, with an explicit parent-directory fallback."""
    candidate, error = _coerce_path(target_path)
    if error is not None or candidate is None:
        return error or _failure("invalid path", error_code="probe_failed")

    try:
        io_candidate = filesystem_path(candidate)
        if not io_candidate.exists():
            return _failure(str(candidate), error_code="missing_path")
        if io_candidate.is_dir():
            return open_path(candidate)
    except Exception as exc:
        return _failure(str(exc), error_code="probe_failed")

    platform_key = _platform_key()
    if platform_key == "windows":
        result = _run_command(["explorer", "/select,", str(candidate)])
        if result.success:
            logger.info("Selected file in Windows Explorer: %s", candidate)
            return _success(precise=True)
        return _reveal_fallback(candidate, result)

    if platform_key == "macos":
        result = _run_command(["open", "-R", str(candidate)])
        if result.success:
            logger.info("Selected file in macOS Finder: %s", candidate)
            return _success(precise=True)
        return _reveal_fallback(candidate, result)

    if platform_key == "linux":
        last_error: PathActionResult | None = None
        for command in linux_reveal_commands(candidate):
            result = _run_command(command)
            if result.success:
                logger.info("Selected file in a Linux file manager: %s", candidate)
                return _success(precise=True)
            last_error = result
        return _reveal_fallback(candidate, last_error)

    return _reveal_fallback(candidate, None)


def _reveal_fallback(candidate: Path, precise_error: PathActionResult | None) -> PathActionResult:
    fallback = open_path(candidate.parent)
    if fallback.success:
        logger.info("Precise reveal unavailable; opened parent: %s", candidate.parent)
        return _success(fallback_used=True)
    return precise_error or fallback


__all__ = ["PathActionResult", "linux_reveal_commands", "open_path", "reveal_path"]
