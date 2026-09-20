from __future__ import annotations

from pathlib import Path

import pytest

from docwen_application.bundle_mapping import BundleMappingError, build_bundle_draft
from docwen_core.models import ArtifactManifest, validate_artifact_bundle_draft

pytestmark = pytest.mark.contract


@pytest.mark.parametrize(
    ("profile", "media_type", "suffix"),
    [
        ("single_document", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"),
        ("table_resources", "text/csv", ".csv"),
        ("report_resource", "text/plain", ".txt"),
    ],
)
def test_explicit_audit_maps_as_a_supplementary_resource_without_node_json(tmp_path, profile, media_type, suffix):
    primary = ArtifactManifest(
        "main",
        "primary",
        str(tmp_path / f"main{suffix}"),
        f"main{suffix}",
        media_type,
        metadata={"table_index": 0, "document_node_schema": "docwen.document_node.v1"},
        is_primary=True,
        logical_path=f"result/main{suffix}",
        size_bytes=4,
        sha256="a" * 64,
    )
    audit = ArtifactManifest(
        "audit",
        "manifest",
        str(tmp_path / "manifest.json"),
        "manifest.json",
        "application/json",
        metadata={"document_node_role": "audit", "document_node_schema": "docwen.document_node.v1"},
        logical_path="result/manifest.json",
        size_bytes=2,
        sha256="b" * 64,
    )
    draft = build_bundle_draft(profile=profile, output_media_type=media_type, artifacts=[primary, audit])
    validate_artifact_bundle_draft(draft)
    assert draft.layout_schema == "docwen.document_node.v1"
    assert len(draft.artifacts) == 2
    assert draft.artifacts[-1].kind == "resource"
    assert draft.artifacts[-1].expected_sha256 == audit.sha256
    assert draft.artifacts[-1].expected_size_bytes == audit.size_bytes
    assert draft.entries[-1].artifact_id == "audit" and draft.entries[-1].role == "supplementary"
    assert not draft.entries[-1].preferred
    assert not draft.relations


def test_gongwen_document_node_maps_attachment_without_node_json(tmp_path: Path) -> None:
    root_name = "notice_20260820_120000_fromDocx"
    root = tmp_path / root_name
    child_name = "notice_附件_20260820_120000_fromDocx"
    child = root / child_name
    child.mkdir(parents=True)
    primary_path = root / f"{root_name}.md"
    attachment_path = child / f"{child_name}.md"
    primary_path.write_text("# Notice\n", encoding="utf-8")
    attachment_path.write_text("# Attachment\n", encoding="utf-8")

    primary = ArtifactManifest(
        artifact_id="document.main",
        kind="primary",
        staging_path=str(primary_path),
        suggested_name=primary_path.name,
        media_type="text/markdown",
        is_primary=True,
        metadata={"document_node_schema": "docwen.document_node.v1"},
        logical_path=f"{root_name}/{primary_path.name}",
    )
    attachment = ArtifactManifest(
        artifact_id="document.attachment.1",
        kind="auxiliary",
        staging_path=str(attachment_path),
        suggested_name=attachment_path.name,
        media_type="text/markdown",
        metadata={
            "source_kind": "gongwen_attachment",
            "attachment_ordinal": 1,
            "document_node_schema": "docwen.document_node.v1",
        },
        logical_path=f"{root_name}/{child_name}/{attachment_path.name}",
    )
    draft = build_bundle_draft(
        profile="document_with_resources",
        output_media_type="text/markdown",
        artifacts=(primary, attachment),
    )

    validate_artifact_bundle_draft(draft)
    assert draft.layout_schema == "docwen.document_node.v1"
    assert [artifact.kind for artifact in draft.artifacts] == ["document", "document"]
    assert [(relation.type, relation.role) for relation in draft.relations] == [
        ("attachment_of", "attachment"),
    ]
    assert [artifact.logical_path for artifact in draft.artifacts] == [
        primary.logical_path,
        attachment.logical_path,
    ]


@pytest.mark.parametrize(
    ("profile", "code"),
    [("single_document", "unexpected_output_shape"), ("document_with_resources", "artifact_semantics_unknown")],
)
def test_obsolete_node_manifest_is_not_mapped_as_a_resource(tmp_path: Path, profile, code) -> None:
    primary = ArtifactManifest(
        "main", "primary", str(tmp_path / "main.md"), "main.md", "text/markdown", is_primary=True
    )
    obsolete = ArtifactManifest(
        "node",
        "manifest",
        str(tmp_path / "docwen-node.json"),
        "docwen-node.json",
        "application/vnd.docwen.document-node+json",
    )
    with pytest.raises(BundleMappingError) as caught:
        build_bundle_draft(profile=profile, output_media_type="text/markdown", artifacts=(primary, obsolete))
    assert caught.value.code == code
