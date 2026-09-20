"""Freeze user-facing source identity before conversion changes physical inputs."""

from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models.document_node import ConversionIdentity
from docwen_core.models.file_ref import FileRef
from docwen_runtime.path_io import filesystem_path

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import CancellationTokenView


def conversion_identity(
    task_id: str,
    source: FileRef,
    *,
    cancellation: CancellationTokenView | None = None,
) -> ConversionIdentity:
    if cancellation is not None:
        cancellation.check()
    original = source.metadata.get("_docwen_preconversion_source")
    original_path = source.path
    source_format = source.format
    frozen_sha256 = ""
    created_at = None
    if isinstance(original, dict):
        original_path = str(original.get("path") or original_path)
        source_format = str(original.get("format") or source_format)
        frozen_sha256 = str(original.get("sha256") or "")
        if original.get("created_at"):
            created_at = datetime.fromisoformat(str(original["created_at"]))
    name = Path(source.logical_path or original_path).name
    stem = Path(name).stem or "document"
    return ConversionIdentity.create(
        task_id=task_id,
        source_stem=stem,
        source_format=source_format,
        source_name=name,
        source_sha256=frozen_sha256 or _source_sha256(original_path, cancellation),
        created_at=created_at,
    )


def _source_sha256(path: str, cancellation: CancellationTokenView | None) -> str:
    source = filesystem_path(path)
    if not source.is_file():
        return ""
    digest = hashlib.sha256()
    with source.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            if cancellation is not None:
                cancellation.check()
            digest.update(block)
    return digest.hexdigest()
