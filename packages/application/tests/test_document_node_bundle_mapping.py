from __future__ import annotations

from pathlib import Path

import pytest

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_core.models import ArtifactManifest, validate_artifact_bundle_draft

pytestmark = pytest.mark.contract


def test_gongwen_document_node_maps_attachment_and_manifest(tmp_path: Path) -> None:
    root_name = "notice_20260820_120000_fromDocx"
    root = tmp_path / root_name
    child_name = "notice_附件01_20260820_120000_fromDocx"
    child = root / child_name
    child.mkdir(parents=True)
    primary_path = root / f"{root_name}.md"
    attachment_path = child / f"{child_name}.md"
    manifest_path = root / "docwen-node.json"
    primary_path.write_text("# Notice\n", encoding="utf-8")
    attachment_path.write_text("# Attachment\n", encoding="utf-8")
    manifest_path.write_text("{}\n", encoding="utf-8")

    primary = ArtifactManifest(
        artifact_id="document.main",
        kind="primary",
        staging_path=str(primary_path),
        suggested_name=primary_path.name,
        media_type="text/markdown",
        is_primary=True,
        logical_path=f"{root_name}/{primary_path.name}",
    )
    attachment = ArtifactManifest(
        artifact_id="document.attachment.1",
        kind="auxiliary",
        staging_path=str(attachment_path),
        suggested_name=attachment_path.name,
        media_type="text/markdown",
        metadata={"source_kind": "gongwen_attachment", "attachment_ordinal": 1},
        logical_path=f"{root_name}/{child_name}/{attachment_path.name}",
    )
    manifest = ArtifactManifest(
        artifact_id="manifest.node",
        kind="manifest",
        staging_path=str(manifest_path),
        suggested_name=manifest_path.name,
        media_type="application/vnd.docwen.document-node+json",
        logical_path=f"{root_name}/docwen-node.json",
    )

    draft = build_bundle_draft(
        profile="document_with_resources",
        output_media_type="text/markdown",
        artifacts=(primary, attachment, manifest),
    )

    validate_artifact_bundle_draft(draft)
    assert [artifact.kind for artifact in draft.artifacts] == ["document", "document", "resource"]
    assert [(relation.type, relation.role) for relation in draft.relations] == [
        ("attachment_of", "attachment"),
        ("resource_of", "manifest"),
    ]
    assert [artifact.logical_path for artifact in draft.artifacts] == [
        primary.logical_path,
        attachment.logical_path,
        manifest.logical_path,
    ]
