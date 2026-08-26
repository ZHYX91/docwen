"""Tests for OutputFinalizer — the single choke-point for final output writes."""

from __future__ import annotations

import os
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from docwen_core.models.artifact import ARTIFACT_KIND_PRIMARY, ArtifactManifest
from docwen_core.models.request import OutputPolicy
from docwen_runtime.output.finalizer import OutputFinalizer

pytestmark = pytest.mark.integration


@pytest.fixture
def finalizer() -> OutputFinalizer:
    return OutputFinalizer()


def _create_staging_file(staging_dir: str, name: str, content: str = "test content") -> str:
    """Create a real file in staging to simulate plugin output."""
    path = os.path.join(staging_dir, name)
    with open(path, "w") as f:
        f.write(content)
    return path


__all__ = (
    "ARTIFACT_KIND_PRIMARY",
    "ArtifactManifest",
    "OutputFinalizer",
    "OutputPolicy",
    "Path",
    "ThreadPoolExecutor",
    "_create_staging_file",
    "finalizer",
    "os",
    "pytest",
    "pytestmark",
    "tempfile",
    "threading",
)
