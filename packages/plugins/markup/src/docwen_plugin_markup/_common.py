"""Shared utilities for the Markup plugin — kept small and dependency-free."""

from __future__ import annotations

import uuid
from pathlib import Path


def new_artifact_id() -> str:
    """Return a new unique artifact identifier."""
    return f"markup-{uuid.uuid4().hex[:12]}"


def file_size(path: str | Path) -> int:
    """Return file size in bytes, or 0 if the file does not exist."""
    try:
        return Path(path).stat().st_size
    except OSError:
        return 0
