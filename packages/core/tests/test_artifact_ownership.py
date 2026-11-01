from __future__ import annotations

from dataclasses import replace

import pytest

from docwen_core.models import ArtifactManifest, ArtifactOwnership, ConversionResult, validate_artifact_ownership

pytestmark = pytest.mark.contract


def _artifacts():
    return [
        ArtifactManifest("main", "primary", "main.md", "main.md", "text/markdown", is_primary=True),
        ArtifactManifest(
            "child",
            "auxiliary",
            "child.md",
            "child.md",
            "text/markdown",
            ownership=ArtifactOwnership("attachment_of", "main", "attachment", 0, "Supplement"),
        ),
        ArtifactManifest(
            "image",
            "image",
            "image.png",
            "image.png",
            "image/png",
            ownership=ArtifactOwnership("resource_of", "child", "image", 0),
        ),
    ]


def test_ownership_survives_worker_result_wire_roundtrip():
    artifacts = _artifacts()
    result = ConversionResult(task_id="test", success=True, artifacts=artifacts)
    restored = ConversionResult.from_dict(result.to_dict())
    assert [item.ownership for item in restored.artifacts] == [item.ownership for item in artifacts]
    assert "ownership" not in artifacts[0].to_dict()
    validate_artifact_ownership(restored.artifacts)


@pytest.mark.parametrize("case", ["attachment_under_fragment", "fragment_under_fragment", "physical_page"])
def test_ownership_shorthand_cannot_publish_bundle_incompatible_structure(case):
    artifacts = _artifacts()
    if case == "physical_page":
        artifacts[1].ownership = ArtifactOwnership("fragment_of", "main", "ocr_page", 0)
    else:
        artifacts[1].ownership = ArtifactOwnership("fragment_of", "main", "ocr_text", 0)
        nested = ArtifactManifest("nested", "auxiliary", "nested.md", "nested.md", "text/markdown")
        nested.ownership = (
            ArtifactOwnership("attachment_of", "child", "attachment", 0)
            if case == "attachment_under_fragment"
            else ArtifactOwnership("fragment_of", "child", "ocr_text", 0)
        )
        artifacts.append(nested)
    with pytest.raises(ValueError):
        validate_artifact_ownership(artifacts)


@pytest.mark.parametrize(
    "metadata,media_type",
    [
        ({"source_kind": "embedded_resource"}, "text/markdown"),
        (
            {"source_format": "pptx", "payload": "chart_workbook"},
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ),
        ({}, "text/markdown"),
    ],
)
def test_legacy_resource_or_unknown_auxiliary_cannot_become_a_resource_owner(metadata, media_type):
    artifacts = _artifacts()
    artifacts[1].ownership = None
    artifacts[1].metadata = metadata
    artifacts[1].media_type = media_type
    with pytest.raises(ValueError):
        validate_artifact_ownership(artifacts)


@pytest.mark.parametrize(
    "case", ["missing", "self", "cycle", "duplicate", "primary", "resource_owner", "ordinal", "image_type"]
)
def test_invalid_ownership_is_refused(case):
    artifacts = _artifacts()
    if case == "missing":
        artifacts[2].ownership = ArtifactOwnership("resource_of", "missing", "image", 0)
    elif case == "self":
        artifacts[1].ownership = ArtifactOwnership("attachment_of", "child", "attachment", 0)
    elif case == "cycle":
        artifacts[0].is_primary = False
        artifacts[0].kind = "auxiliary"
        artifacts[0].ownership = ArtifactOwnership("attachment_of", "child", "attachment", 0)
    elif case == "duplicate":
        artifacts.append(replace(artifacts[2]))
    elif case == "primary":
        artifacts[1].is_primary = True
    elif case == "resource_owner":
        artifacts[1].ownership = ArtifactOwnership("attachment_of", "image", "attachment", 0)
    elif case == "ordinal":
        artifacts[1].ownership = ArtifactOwnership("attachment_of", "main", "attachment", True)
    else:
        artifacts[2].media_type = "application/octet-stream"
    with pytest.raises(ValueError):
        validate_artifact_ownership(artifacts)
