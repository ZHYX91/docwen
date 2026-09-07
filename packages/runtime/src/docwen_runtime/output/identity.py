"""Freeze user-facing source identity before conversion changes physical inputs."""

from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models.document_node import DOCUMENT_NODE_SCHEMA, ConversionIdentity
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
    # A filename alone never proves provenance. Only an intact artifact named by
    # its adjacent DocWen manifest can reuse the original label on a later pass.
    if name == Path(original_path).name:
        stem = _verified_source_stem(original_path, cancellation) or stem
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


def _verified_source_stem(input_path: str, cancellation: CancellationTokenView | None) -> str | None:
    path = filesystem_path(input_path)
    manifest = path.parent / "docwen-node.json"
    try:
        if manifest.is_symlink() or manifest.stat().st_size > 1024 * 1024:
            return None
        document = json.loads(manifest.read_text(encoding="utf-8"))
        if not isinstance(document, dict) or document.get("schema") != DOCUMENT_NODE_SCHEMA:
            return None
        if document.get("node_name") != path.parent.name:
            return None
        original = document.get("source")
        artifacts = document.get("artifacts")
        if not isinstance(original, dict) or not isinstance(artifacts, list):
            return None
        stem = original.get("stem")
        if not isinstance(stem, str) or not stem:
            return None
        logical = f"{path.parent.name}/{path.name}"
        matches = [item for item in artifacts if isinstance(item, dict) and item.get("logical_path") == logical]
        if len(matches) != 1:
            return None
        return stem if matches[0].get("sha256") == _source_sha256(input_path, cancellation) else None
    except (OSError, ValueError, TypeError):
        return None
