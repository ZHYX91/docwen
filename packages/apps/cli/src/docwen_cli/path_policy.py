"""Public CLI path compatibility policy.

DocWen 0.9 intentionally rejects Windows paths that require extended-length
syntax.  Many conversion backends still receive ordinary filesystem paths and
do not consistently understand ``\\\\?\\`` names.  Failing at the CLI boundary
keeps the machine contract deterministic instead of leaking backend-specific
``NotImplementedError`` or ``OSError`` failures.
"""

from __future__ import annotations

import argparse
import ntpath
import os
from dataclasses import dataclass
from typing import Any

from docwen_runtime.path_io import windows_utf16_units
from docwen_runtime.templates import is_canonical_template_id

WINDOWS_PUBLIC_PATH_LIMIT = 259
_PATH_FIELDS = (
    "file",
    "files",
    "output",
    "output_dir",
    "report",
    "report_dir",
    "template",
)


@dataclass(frozen=True, slots=True)
class PathPolicyIssue:
    """One actionable public-path compatibility failure."""

    path: str
    message: str


def check_public_path(raw_path: str, *, platform_name: str | None = None) -> PathPolicyIssue | None:
    """Return a compatibility issue for one public CLI path, if any."""

    platform = os.name if platform_name is None else platform_name
    if platform != "nt":
        return None

    expanded = os.path.expanduser(raw_path)
    absolute = ntpath.abspath(expanded)

    def has_namespace_prefix(value: str) -> bool:
        windows_spelling = value.replace("/", "\\")
        return windows_spelling.startswith(("\\\\?\\", "\\\\.\\"))

    if has_namespace_prefix(expanded) or has_namespace_prefix(absolute):
        return PathPolicyIssue(
            path=raw_path,
            message=(
                "Windows extended-length path syntax is not supported by DocWen 0.9; "
                "use an ordinary absolute path no longer than 259 UTF-16 code units."
            ),
        )

    if windows_utf16_units(absolute) > WINDOWS_PUBLIC_PATH_LIMIT:
        return PathPolicyIssue(
            path=raw_path,
            message=(
                "Path exceeds DocWen 0.9's Windows compatibility limit of "
                f"{WINDOWS_PUBLIC_PATH_LIMIT} UTF-16 code units."
            ),
        )
    return None


def first_namespace_path_issue(args: argparse.Namespace) -> PathPolicyIssue | None:
    """Validate every public path-bearing field present in parsed arguments."""

    for field in _PATH_FIELDS:
        value: Any = getattr(args, field, None)
        values = value if isinstance(value, list) else [value]
        for item in values:
            if not isinstance(item, str) or not item:
                continue
            if field == "template" and is_canonical_template_id(item):
                continue
            issue = check_public_path(item)
            if issue is not None:
                return issue
    return None


__all__ = [
    "WINDOWS_PUBLIC_PATH_LIMIT",
    "PathPolicyIssue",
    "check_public_path",
    "first_namespace_path_issue",
]
