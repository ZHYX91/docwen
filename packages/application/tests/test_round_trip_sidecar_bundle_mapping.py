"""Consumer-neutral mapping for the public DOCX round-trip sidecar."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_application.bundle_mapping import BundleMappingError, build_bundle_draft
from docwen_core.models.artifact import ArtifactManifest
from docwen_core.round_trip_sidecar import (
    DOCX_MEDIA_TYPE,
    ROUND_TRIP_SIDECAR_MEDIA_TYPE,
    ROUND_TRIP_SIDECAR_OWNER_METADATA,
    ROUND_TRIP_SIDECAR_SCHEMA,
    ROUND_TRIP_SIDECAR_SCHEMA_METADATA,
)

pytestmark = pytest.mark.unit


def _artifacts(tmp_path: Path) -> tuple[ArtifactManifest, ArtifactManifest]:
    docx = tmp_path / "document.docx"
    sidecar = tmp_path / "document.docx.docwen"
    docx.write_bytes(b"docx")
    sidecar.write_bytes(b"sidecar")
    primary = ArtifactManifest(
        artifact_id="document.main",
        kind="primary",
        staging_path=str(docx),
        suggested_name="document.docx",
        media_type=DOCX_MEDIA_TYPE,
        is_primary=True,
    )
    companion = ArtifactManifest(
        artifact_id="document.sidecar",
        kind="auxiliary",
        staging_path=str(sidecar),
        suggested_name="document.docx.docwen",
        media_type=ROUND_TRIP_SIDECAR_MEDIA_TYPE,
        metadata={
            ROUND_TRIP_SIDECAR_SCHEMA_METADATA: ROUND_TRIP_SIDECAR_SCHEMA,
            ROUND_TRIP_SIDECAR_OWNER_METADATA: primary.artifact_id,
        },
    )
    return primary, companion


def test_round_trip_sidecar_maps_as_one_manifest_resource_of_docx(tmp_path: Path) -> None:
    primary, sidecar = _artifacts(tmp_path)

    draft = build_bundle_draft(
        profile="document_with_round_trip_sidecar",
        output_media_type=DOCX_MEDIA_TYPE,
        artifacts=(primary, sidecar),
    )

    assert [(item.artifact_id, item.kind) for item in draft.artifacts] == [
        (primary.artifact_id, "document"),
        (sidecar.artifact_id, "resource"),
    ]
    assert [(item.artifact_id, item.role, item.preferred) for item in draft.entries] == [
        (primary.artifact_id, "primary", True)
    ]
    assert [item.to_dict() for item in draft.relations] == [
        {
            "type": "resource_of",
            "source_artifact_id": sidecar.artifact_id,
            "target_artifact_id": primary.artifact_id,
            "role": "manifest",
            "ordinal": 0,
        }
    ]


@pytest.mark.parametrize("mutation", ["owner", "schema", "name", "media", "missing"])
def test_round_trip_sidecar_mapping_fails_closed_for_invalid_pair(tmp_path: Path, mutation: str) -> None:
    primary, sidecar = _artifacts(tmp_path)
    metadata = dict(sidecar.metadata)
    artifacts: tuple[ArtifactManifest, ...] = (primary, sidecar)
    if mutation == "owner":
        metadata[ROUND_TRIP_SIDECAR_OWNER_METADATA] = "other"
        artifacts = (primary, ArtifactManifest.from_dict({**sidecar.to_dict(), "metadata": metadata}))
    elif mutation == "schema":
        metadata[ROUND_TRIP_SIDECAR_SCHEMA_METADATA] = "foreign"
        artifacts = (primary, ArtifactManifest.from_dict({**sidecar.to_dict(), "metadata": metadata}))
    elif mutation == "name":
        artifacts = (
            primary,
            ArtifactManifest.from_dict({**sidecar.to_dict(), "suggested_name": "detached.docwen"}),
        )
    elif mutation == "media":
        artifacts = (
            ArtifactManifest.from_dict({**primary.to_dict(), "media_type": "application/pdf"}),
            sidecar,
        )
    else:
        artifacts = (primary,)

    with pytest.raises(BundleMappingError):
        build_bundle_draft(
            profile="document_with_round_trip_sidecar",
            output_media_type=DOCX_MEDIA_TYPE,
            artifacts=artifacts,
        )
