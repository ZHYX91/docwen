"""Shared native artifact fixture construction for document-node output contracts."""

from __future__ import annotations

from pathlib import Path

from docwen_core.models import ArtifactManifest, ArtifactOwnership


def _artifact(
    path: Path,
    *,
    artifact_id: str,
    suggested_name: str,
    media_type: str,
    primary: bool = False,
    metadata: dict | None = None,
    ownership: ArtifactOwnership | None = None,
) -> ArtifactManifest:
    return ArtifactManifest(
        artifact_id=artifact_id,
        kind="primary" if primary else "auxiliary",
        staging_path=str(path),
        suggested_name=suggested_name,
        media_type=media_type,
        metadata=metadata or {},
        is_primary=primary,
        ownership=ownership,
    )
