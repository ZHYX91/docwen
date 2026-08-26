"""Shared utilities for the Invoice plugin — kept small and dependency-free."""

from __future__ import annotations

import uuid
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


def new_artifact_id() -> str:
    """Return a new unique artifact identifier."""
    return f"invoice-{uuid.uuid4().hex[:12]}"


def file_size(path: str | Path) -> int:
    """Return file size in bytes, or 0 if the file does not exist."""
    try:
        return Path(path).stat().st_size
    except OSError:
        return 0


def request_source_format(context: ConverterContext) -> str:
    """Return the concrete source format frozen at file admission."""
    refs = context.request.input_refs
    if not refs:
        return "unknown"
    return str(refs[0].format or "unknown").strip().lower()
