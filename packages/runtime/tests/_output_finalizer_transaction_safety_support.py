"""Adversarial transaction-safety contracts for ``OutputFinalizer``."""

from __future__ import annotations

import errno
import multiprocessing
import os
import sys
import threading
import time
from pathlib import Path
from typing import Any

import pytest

from docwen_core.cancellation import CancellationToken
from docwen_core.errors import CancellationRequested
from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
from docwen_core.models.request import OutputPolicy
from docwen_runtime.output.finalizer import OutputFinalizer

pytestmark = pytest.mark.integration


def _hold_process_lock(output_dir: str, ready: Any, release: Any) -> None:
    with OutputFinalizer._process_lock(output_dir, None):
        ready.set()
        if not release.wait(10.0):
            raise TimeoutError("parent did not release process-lock probe")


def _finalize_process(
    source: str,
    output_dir: str,
    task_id: str,
    start: Any,
    results: Any,
) -> None:
    start.wait(10.0)
    result = OutputFinalizer().finalize(
        task_id=task_id,
        artifacts=[_artifact(Path(source))],
        policy=OutputPolicy(output_dir=output_dir, overwrite_mode="rename"),
    )
    results.put(
        (
            result.success,
            Path(result.artifacts[0].staging_path).name if result.artifacts else "",
            result.error.message if result.error else "",
        )
    )


def _artifact(staging_path: Path, suggested_name: str = "report.md") -> ArtifactManifest:
    return ArtifactManifest(
        artifact_id="primary",
        kind=ARTIFACT_KIND_PRIMARY,
        staging_path=str(staging_path),
        suggested_name=suggested_name,
        is_primary=True,
    )


def _relative_file_name_with_length(length: int, *, suffix: str = ".md") -> str:
    """Build a nested relative filename with an exact Windows character count."""
    minimum = len(suffix) + 1
    if length < minimum:
        raise ValueError(f"Relative filename length must be at least {minimum}")
    parts: list[str] = []
    remaining = length
    while remaining > 84:
        parts.append("p" * 80)
        remaining -= 81  # one path separator plus the component
    parts.append(f"{'r' * (remaining - len(suffix))}{suffix}")
    relative = os.path.join(*parts)
    assert len(relative) == length
    return relative


def _relative_dir_name_with_length(length: int) -> str:
    """Build a nested relative directory name with an exact character count."""
    if length < 1:
        raise ValueError("Relative directory length must be positive")
    parts: list[str] = []
    remaining = length
    while remaining > 81:
        parts.append("d" * 80)
        remaining -= 81
    parts.append("d" * remaining)
    relative = os.path.join(*parts)
    assert len(relative) == length
    return relative


def _suggested_name_for_total_length(output_dir: Path, total_length: int) -> str:
    output_abs = os.path.abspath(output_dir)
    relative_length = total_length - len(output_abs) - 1
    if relative_length < 4:
        pytest.skip(f"pytest temp root is already too long for {total_length}-character boundary")
    return _relative_file_name_with_length(relative_length)


def _output_dir_with_total_length(tmp_path: Path, total_length: int) -> Path:
    root_abs = os.path.abspath(tmp_path)
    relative_length = total_length - len(root_abs) - 1
    if relative_length < 1:
        pytest.skip(f"pytest temp root is already too long for {total_length}-character output root")
    return tmp_path / _relative_dir_name_with_length(relative_length)


def _io_bytes(path: os.PathLike[str] | str) -> bytes:
    return OutputFinalizer._io_path(path).read_bytes()


def _io_names(parent: os.PathLike[str] | str) -> list[str]:
    with os.scandir(OutputFinalizer._io_path(parent)) as entries:
        return sorted(entry.name for entry in entries)


def _utf16_units(value: os.PathLike[str] | str) -> int:
    return len(os.fspath(value).encode("utf-16-le", errors="surrogatepass")) // 2


__all__ = (
    "Any",
    "ArtifactManifest",
    "CancellationRequested",
    "CancellationToken",
    "OutputFinalizer",
    "OutputPolicy",
    "Path",
    "_artifact",
    "_finalize_process",
    "_hold_process_lock",
    "_io_bytes",
    "_io_names",
    "_output_dir_with_total_length",
    "_suggested_name_for_total_length",
    "_utf16_units",
    "errno",
    "multiprocessing",
    "os",
    "pytest",
    "pytestmark",
    "sys",
    "threading",
    "time",
)
