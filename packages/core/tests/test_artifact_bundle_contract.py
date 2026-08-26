"""Typed Artifact Bundle v2 round-trip and semantic validation contracts."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docwen_core.models import (
    ArtifactBundle,
    ArtifactBundleValidationError,
    BundleDraft,
    BundleDraftArtifact,
    BundleEntry,
    BundleRelation,
    PageFragmentSemantics,
    validate_artifact_bundle,
    validate_artifact_bundle_draft,
)

pytestmark = pytest.mark.contract

CONTRACTS_ROOT = Path(__file__).resolve().parents[3] / "contracts"
CONFORMANCE_MANIFEST = json.loads((CONTRACTS_ROOT / "conformance-manifest.json").read_text(encoding="utf-8"))
VALID_BUNDLE_RECORDS = [
    record
    for record in CONFORMANCE_MANIFEST["fixtures"]
    if record["document_type"] == "bundle" and record["expect"] == "valid"
]
SEMANTIC_INVALID_BUNDLE_RECORDS = [
    record
    for record in CONFORMANCE_MANIFEST["fixtures"]
    if record["document_type"] == "bundle" and record["expect"] == "invalid_semantic"
]


@pytest.mark.parametrize(
    "record",
    VALID_BUNDLE_RECORDS,
    ids=lambda record: Path(record["path"]).stem,
)
def test_valid_conformance_bundle_round_trips_without_wire_drift(record: dict[str, str]) -> None:
    fixture = CONTRACTS_ROOT / record["path"]
    payload = json.loads(fixture.read_text(encoding="utf-8"))
    bundle = ArtifactBundle.from_dict(payload)

    validate_artifact_bundle(bundle)
    assert bundle.to_dict() == payload


def test_v1_bundle_is_rejected_without_inference_or_field_synthesis() -> None:
    fixture = CONTRACTS_ROOT / "fixtures" / "invalid" / "artifact-bundle.duplicate-id.json"
    payload = json.loads(fixture.read_text(encoding="utf-8"))
    payload["schema"] = "docwen.artifact_bundle.v1"
    bundle = ArtifactBundle.from_dict(payload)

    with pytest.raises(ArtifactBundleValidationError) as rejected:
        validate_artifact_bundle(bundle)

    assert rejected.value.code == "unsupported_bundle_schema"


@pytest.mark.parametrize(
    "record",
    SEMANTIC_INVALID_BUNDLE_RECORDS,
    ids=lambda record: Path(record["path"]).stem,
)
def test_typed_validator_matches_semantic_conformance_error(record: dict[str, str]) -> None:
    fixture = CONTRACTS_ROOT / record["path"]
    bundle = ArtifactBundle.from_dict(json.loads(fixture.read_text(encoding="utf-8")))

    with pytest.raises(ArtifactBundleValidationError) as rejected:
        validate_artifact_bundle(bundle)

    assert rejected.value.code == record["error_code"]


@pytest.mark.parametrize(
    ("field", "error_code"),
    [
        ("artifacts", "empty_artifacts"),
        ("entries", "empty_entries"),
    ],
)
def test_typed_validator_rejects_empty_schema_required_collections(field: str, error_code: str) -> None:
    fixture = CONTRACTS_ROOT / "fixtures" / "valid" / "artifact-bundle.gongwen.json"
    payload = json.loads(fixture.read_text(encoding="utf-8"))
    payload[field] = []
    bundle = ArtifactBundle.from_dict(payload)

    with pytest.raises(ArtifactBundleValidationError) as rejected:
        validate_artifact_bundle(bundle)

    assert rejected.value.code == error_code


def test_draft_page_semantics_are_validated_before_any_locator_or_file_state() -> None:
    draft = BundleDraft(
        artifacts=(
            BundleDraftArtifact("document.main", "document", "relative-missing.md", "main.md", "text/markdown"),
            BundleDraftArtifact("fragment.page.1", "fragment", "also-missing.md", "page-001.md", "text/markdown"),
        ),
        entries=(BundleEntry("document.main", "primary", 0, True),),
        relations=(
            BundleRelation(
                "fragment_of",
                "fragment.page.1",
                "document.main",
                "ocr_page",
                0,
                PageFragmentSemantics(0, 1, "success", 1),
            ),
        ),
    )

    with pytest.raises(ArtifactBundleValidationError) as rejected:
        validate_artifact_bundle_draft(draft)

    assert rejected.value.code == "invalid_page_range"


def test_page_fragments_must_target_the_primary_document_entry() -> None:
    fixture = CONTRACTS_ROOT / "fixtures" / "valid" / "artifact-bundle.ocr.json"
    payload = json.loads(fixture.read_text(encoding="utf-8"))
    primary_id = payload["entries"][0]["artifact_id"]
    secondary_id = "document.secondary"
    payload["artifacts"].append(
        {
            "artifact_id": secondary_id,
            "kind": "document",
            "locator": "secondary.md",
            "logical_path": "secondary.md",
            "suggested_name": "secondary.md",
            "media_type": "text/markdown",
            "size_bytes": 0,
            "sha256": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
        }
    )
    payload["relations"].append(
        {
            "type": "attachment_of",
            "source_artifact_id": secondary_id,
            "target_artifact_id": primary_id,
            "role": "attachment",
            "ordinal": 0,
        }
    )
    for relation in payload["relations"]:
        if relation["type"] == "fragment_of" and relation["role"] == "ocr_page":
            relation["target_artifact_id"] = secondary_id

    with pytest.raises(ArtifactBundleValidationError) as rejected:
        validate_artifact_bundle(ArtifactBundle.from_dict(payload))

    assert rejected.value.code == "unexpected_page_semantics"


def test_document_owned_page_resource_must_target_the_primary_document_entry() -> None:
    fixture = CONTRACTS_ROOT / "fixtures" / "valid" / "artifact-bundle.ocr.json"
    payload = json.loads(fixture.read_text(encoding="utf-8"))
    primary_id = payload["entries"][0]["artifact_id"]
    secondary_id = "document.secondary"
    payload["artifacts"].append(
        {
            "artifact_id": secondary_id,
            "kind": "document",
            "locator": "secondary.md",
            "logical_path": "secondary.md",
            "suggested_name": "secondary.md",
            "media_type": "text/markdown",
            "size_bytes": 0,
            "sha256": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
        }
    )
    payload["relations"].append(
        {
            "type": "attachment_of",
            "source_artifact_id": secondary_id,
            "target_artifact_id": primary_id,
            "role": "attachment",
            "ordinal": 0,
        }
    )
    unresolved = next(
        relation
        for relation in payload["relations"]
        if relation["type"] == "resource_of" and "page_resource" not in relation
    )
    unresolved["target_artifact_id"] = secondary_id
    unresolved["page_resource"] = {"source_page": 1}

    with pytest.raises(ArtifactBundleValidationError) as rejected:
        validate_artifact_bundle(ArtifactBundle.from_dict(payload))

    assert rejected.value.code == "resource_page_mismatch"
