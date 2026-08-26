"""CLI utility functions for paths, file admission, and progress callbacks."""

from __future__ import annotations

import os
import sys
from collections.abc import Callable
from pathlib import Path

from docwen_cli.file_admission_i18n import (
    render_file_inspection_message,
    render_file_inspection_warning,
)
from docwen_cli.i18n import cli_t
from docwen_core.detection import inspect_file
from docwen_core.models.file_inspection import AdmissionDecision, FileInspection
from docwen_core.paths import scan_input_directory

_GLOB_CHARS = frozenset({"*", "?", "["})


# ── Path expansion ──────────────────────────────────────────────────


def _has_glob(path: str) -> bool:
    return any(ch in path for ch in _GLOB_CHARS)


def _clean_powershell_drag(path: str) -> str:
    """Strip PowerShell drag-and-drop quoting: ``& 'path'`` or ``& \"path\"``."""
    stripped = path.strip()
    if stripped.startswith("& "):
        stripped = stripped[2:].strip()
    if (stripped.startswith("'") and stripped.endswith("'")) or (stripped.startswith('"') and stripped.endswith('"')):
        stripped = stripped[1:-1]
    return stripped


def expand_paths(raw_paths: list[str]) -> list[str]:
    """Expand a list of raw CLI file arguments into absolute paths.

    Handles:
    - PowerShell drag-and-drop quoting
    - Glob patterns (``*``, ``?``, ``[``)
    - Recursive directory expansion
    - Stable first-seen de-duplication across explicit, glob, and directory inputs
    """
    result: list[str] = []
    for raw in raw_paths:
        cleaned = _clean_powershell_drag(raw)
        p = Path(cleaned)

        if _has_glob(cleaned):
            # Expand glob
            base_dir = Path(".")
            pattern = cleaned
            if os.path.isabs(cleaned):
                parts = Path(cleaned).parts
                for idx, part in enumerate(parts):
                    if _has_glob(part):
                        base_dir = Path(*parts[:idx])
                        pattern = str(Path(*parts[idx:]))
                        break
            matched = sorted(str(m.resolve()) for m in base_dir.glob(pattern) if m.is_file())
            result.extend(matched)
        elif p.is_dir():
            scan = scan_input_directory(p)
            result.extend(str(entry.resolve()) for entry in scan.files)
        elif p.exists():
            result.append(str(p.resolve()))
        else:
            # File may not exist yet (e.g. output target); keep as-is
            result.append(os.path.abspath(cleaned))

    seen: set[str] = set()
    ordered: list[str] = []
    for expanded in result:
        normalized = os.path.normcase(os.path.normpath(str(Path(expanded).expanduser().resolve())))
        if normalized in seen:
            continue
        seen.add(normalized)
        ordered.append(expanded)

    return ordered


# ── File validation ─────────────────────────────────────────────────


def validate_files(
    files: list[str],
    *,
    use_detected_format: bool = False,
    inspection_cache: dict[str, FileInspection] | None = None,
) -> tuple[list[str], list[tuple[str, str]], list[tuple[str, str]]]:
    """Validate a list of file paths.

    Returns:
        (valid, invalid_with_reasons, warnings_with_reasons)

    *valid* contains absolute paths whose content has passed the shared Core
    admission policy. An unrecognised suffix can be accepted only when
    ``use_detected_format`` is explicitly enabled.

    *invalid_with_reasons* lists missing files, hard-blocked content, and
    confirmation-required mismatches that were not explicitly accepted.

    *warnings_with_reasons* lists valid-looking files with soft warnings
    such as extension/content mismatches.
    """
    valid: list[str] = []
    invalid: list[tuple[str, str]] = []
    warnings: list[tuple[str, str]] = []

    for f in files:
        abs_path = os.path.abspath(f)
        if not os.path.exists(abs_path):
            invalid.append((f, "文件不存在"))
            continue
        if not os.path.isfile(abs_path):
            invalid.append((f, "不是文件"))
            continue

        try:
            inspection = inspect_file(abs_path)
        except (OSError, ValueError) as exc:
            invalid.append((f, f"无法检查文件: {exc}"))
            continue
        if inspection_cache is not None:
            inspection_cache[abs_path] = inspection

        if inspection.decision is AdmissionDecision.BLOCK:
            reason = render_file_inspection_message(inspection, prefer_reason=True)
            invalid.append((f, f"[{inspection.reason_code or 'INVALID_INPUT'}] {reason}"))
            continue
        if inspection.requires_explicit_acceptance and not use_detected_format:
            reason = render_file_inspection_message(inspection)
            invalid.append((f, f"[{inspection.reason_code}] {reason}"))
            continue

        valid.append(abs_path)
        if inspection.warning_message:
            warnings.append((f, render_file_inspection_warning(inspection)))

    return valid, invalid, warnings


# ── Progress callback ───────────────────────────────────────────────


def create_progress_callback(
    *,
    quiet: bool = False,
    verbose: bool = False,
    json_mode: bool = False,
) -> Callable[[str], None]:
    """Create a progress callback for conversion operations.

    Quiet and JSON modes suppress output. Verbose mode adds the localized
    progress prefix; normal mode writes the message alone to stderr.
    """

    def callback(msg: str) -> None:
        if quiet or json_mode:
            return
        if verbose:
            prefix = cli_t("main_window.task_progress_prefix", "Progress:")
            print(f"{prefix} {msg}", file=sys.stderr, flush=True)
        else:
            print(msg, file=sys.stderr, flush=True)

    return callback
