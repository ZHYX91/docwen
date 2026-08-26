"""Consumer-neutral Bundle mapping for physical-page OCR artifacts."""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_application.bundle_mapping import (
    BundleMappingError,
    build_bundle_draft,
    validate_physical_page_diagnostics,
)
from docwen_core.models import ArtifactManifest, ConversionDiagnostic, validate_artifact_bundle_draft

pytestmark = pytest.mark.contract


def _artifact(
    root: Path,
    artifact_id: str,
    kind: str,
    suffix: str,
    *,
    metadata: dict[str, object] | None = None,
    primary: bool = False,
) -> ArtifactManifest:
    path = root / f"{artifact_id}{suffix}"
    path.write_bytes(b"")
    return ArtifactManifest(
        artifact_id=artifact_id,
        kind=kind,
        staging_path=str(path),
        suggested_name=path.name,
        media_type="text/markdown" if suffix == ".md" else "image/png",
        metadata=dict(metadata or {}),
        is_primary=primary,
    )


@pytest.mark.parametrize(
    ("recognize_text", "preserve_resources"),
    [(False, False), (False, True), (True, False), (True, True)],
)
def test_physical_page_profile_freezes_all_four_fidelity_bundle_shapes(
    tmp_path: Path,
    recognize_text: bool,
    preserve_resources: bool,
) -> None:
    page_count = 2
    resource_count = 3
    primary = _artifact(
        tmp_path,
        "document.main",
        "primary",
        ".md",
        metadata={
            "physical_page_count": page_count,
            "ocr_enabled": recognize_text,
            "keep_images": preserve_resources,
        },
        primary=True,
    )
    pages = (
        [
            _artifact(
                tmp_path,
                f"fragment.page.{page}",
                "auxiliary",
                ".md",
                metadata={
                    "fragment_kind": "page",
                    "page_index": page,
                    "page_count": page_count,
                    "source_page": page,
                    "ocr_status": "success",
                },
            )
            for page in range(1, page_count + 1)
        ]
        if recognize_text
        else []
    )
    resources = (
        [
            _artifact(
                tmp_path,
                f"resource.{index}",
                "image",
                ".png",
                metadata={"source_page": 1 if index == 1 else 2},
            )
            for index in range(1, resource_count + 1)
        ]
        if preserve_resources
        else []
    )

    draft = build_bundle_draft(
        profile="physical_page_ocr",
        output_media_type="text/markdown",
        artifacts=[primary, *pages, *resources],
    )
    validate_artifact_bundle_draft(draft)

    assert [(entry.role, entry.ordinal, entry.preferred) for entry in draft.entries] == [("primary", 0, True)]
    assert [artifact.kind for artifact in draft.artifacts].count("document") == 1
    assert [artifact.kind for artifact in draft.artifacts].count("fragment") == (page_count if recognize_text else 0)
    assert [artifact.kind for artifact in draft.artifacts].count("resource") == (
        resource_count if preserve_resources else 0
    )
    assert [relation.type for relation in draft.relations].count("fragment_of") == (page_count if recognize_text else 0)
    assert [relation.type for relation in draft.relations].count("resource_of") == (
        resource_count if preserve_resources else 0
    )
    assert all(entry.artifact_id == primary.artifact_id for entry in draft.entries)
    assert all(resource.artifact_id not in {entry.artifact_id for entry in draft.entries} for resource in resources)
    if recognize_text and preserve_resources:
        fragment_ids = {page.artifact_id for page in pages}
        assert all(
            relation.target_artifact_id in fragment_ids
            for relation in draft.relations
            if relation.type == "resource_of"
        )
    elif preserve_resources:
        assert all(
            relation.target_artifact_id == primary.artifact_id
            for relation in draft.relations
            if relation.type == "resource_of"
        )


@pytest.mark.parametrize(
    ("recognize_text", "preserve_resources"),
    [(False, False), (False, True), (True, False), (True, True)],
)
def test_physical_page_document_node_commits_all_fidelity_shapes(
    tmp_path: Path,
    recognize_text: bool,
    preserve_resources: bool,
) -> None:
    from docwen_core.models import OutputPolicy
    from docwen_runtime.output.artifact_bundle import ArtifactBundleCommitter
    from docwen_runtime.output.finalizer import OutputFinalizer

    source = tmp_path / "source.pdf"
    source.write_bytes(b"pdf")
    producer = tmp_path / "producer"
    producer.mkdir()
    staging = tmp_path / "staging"
    primary = _artifact(
        producer,
        "document.main",
        "primary",
        ".md",
        metadata={
            "physical_page_count": 2,
            "ocr_enabled": recognize_text,
            "keep_images": preserve_resources,
        },
        primary=True,
    )
    pages = (
        [
            _artifact(
                producer,
                f"fragment.page.{page}",
                "auxiliary",
                ".md",
                metadata={
                    "fragment_kind": "page",
                    "page_index": page,
                    "page_count": 2,
                    "source_page": page,
                    "ocr_status": "success",
                },
            )
            for page in range(1, 3)
        ]
        if recognize_text
        else []
    )
    resources = (
        [
            _artifact(
                producer,
                f"resource.{page}",
                "image",
                ".png",
                metadata={"source_page": page},
            )
            for page in range(1, 3)
        ]
        if preserve_resources
        else []
    )

    result = OutputFinalizer().finalize(
        "task.physical.node",
        [primary, *pages, *resources],
        OutputPolicy(output_dir=str(staging), overwrite_mode="error"),
        input_path=str(source),
    )
    assert result.success is True
    draft = build_bundle_draft(
        profile="physical_page_ocr",
        output_media_type="text/markdown",
        artifacts=result.artifacts,
    )

    bundle = ArtifactBundleCommitter().commit(
        task_id=result.task_id,
        staging_root=str(staging),
        draft=draft,
    )

    assert bundle.layout_schema == "docwen.document_node.v1"
    assert any(relation.role == "manifest" for relation in bundle.relations)


def test_physical_page_profile_maps_fragments_and_page_resources(tmp_path: Path) -> None:
    primary = _artifact(
        tmp_path,
        "document.main",
        "primary",
        ".md",
        metadata={"physical_page_count": 4, "ocr_enabled": True, "keep_images": True},
        primary=True,
    )
    page_statuses = ("success", "no_text", "recognition_failed", "success")
    pages = [
        _artifact(
            tmp_path,
            f"fragment.page.{page}",
            "auxiliary",
            ".md",
            metadata={
                "fragment_kind": "page",
                "page_index": page,
                "page_count": 4,
                "source_page": page,
                "ocr_status": page_statuses[page - 1],
            },
        )
        for page in range(1, 5)
    ]
    resources = [
        _artifact(
            tmp_path,
            f"resource.{page}",
            "image",
            ".png",
            metadata={"source_page": page},
        )
        for page in range(1, 5)
    ]
    unresolved = _artifact(tmp_path, "resource.unresolved", "image", ".png")

    draft = build_bundle_draft(
        profile="physical_page_ocr",
        output_media_type="text/markdown",
        artifacts=[primary, *pages, *resources, unresolved],
    )

    validate_artifact_bundle_draft(draft)
    page_relations = [relation for relation in draft.relations if relation.role == "ocr_page"]
    assert [relation.page_fragment.ocr_status for relation in page_relations if relation.page_fragment] == list(
        page_statuses
    )
    resource_relations = [relation for relation in draft.relations if relation.type == "resource_of"]
    assert [relation.target_artifact_id for relation in resource_relations[:4]] == [
        f"fragment.page.{page}" for page in range(1, 5)
    ]
    assert [relation.page_resource.source_page for relation in resource_relations[:4] if relation.page_resource] == [
        1,
        2,
        3,
        4,
    ]
    assert resource_relations[-1].target_artifact_id == primary.artifact_id
    assert resource_relations[-1].page_resource is None


def test_physical_page_profile_keeps_proven_resources_on_primary_when_ocr_is_off(tmp_path: Path) -> None:
    primary = _artifact(
        tmp_path,
        "document.main",
        "primary",
        ".md",
        metadata={"physical_page_count": 2, "ocr_enabled": False, "keep_images": True},
        primary=True,
    )
    resources = [
        _artifact(tmp_path, f"resource.{page}", "image", ".png", metadata={"source_page": page}) for page in range(1, 3)
    ]

    draft = build_bundle_draft(
        profile="physical_page_ocr",
        output_media_type="text/markdown",
        artifacts=[primary, *resources],
    )

    validate_artifact_bundle_draft(draft)
    assert all(relation.target_artifact_id == primary.artifact_id for relation in draft.relations)
    assert [relation.page_resource.source_page for relation in draft.relations if relation.page_resource] == [1, 2]


def test_physical_page_profile_rejects_duplicate_ids_before_preferred_shape(tmp_path: Path) -> None:
    first = _artifact(
        tmp_path,
        "document.same",
        "primary",
        ".md",
        metadata={"physical_page_count": 1, "ocr_enabled": False, "keep_images": False},
        primary=True,
    )
    second = _artifact(
        tmp_path,
        "document.same",
        "primary",
        ".md",
        metadata={"physical_page_count": 1, "ocr_enabled": False, "keep_images": False},
        primary=True,
    )

    with pytest.raises(BundleMappingError) as exc_info:
        build_bundle_draft(profile="physical_page_ocr", output_media_type="text/markdown", artifacts=[first, second])

    assert exc_info.value.code == "duplicate_artifact_id"


def test_physical_page_profile_preserves_typed_page_error_order(tmp_path: Path) -> None:
    primary = _artifact(
        tmp_path,
        "document.main",
        "primary",
        ".md",
        metadata={"physical_page_count": 2, "ocr_enabled": True, "keep_images": True},
        primary=True,
    )
    invalid_page = _artifact(
        tmp_path,
        "fragment.page.1",
        "auxiliary",
        ".md",
        metadata={
            "fragment_kind": "page",
            "page_index": 1,
            "page_count": 2,
            "source_page": 1,
            "ocr_status": [],
        },
    )
    unmatched = _artifact(tmp_path, "resource.2", "image", ".png", metadata={"source_page": 2})

    with pytest.raises(BundleMappingError) as exc_info:
        build_bundle_draft(
            profile="physical_page_ocr",
            output_media_type="text/markdown",
            artifacts=[primary, invalid_page, unmatched],
        )

    assert exc_info.value.code == "invalid_page_range"


def test_physical_page_profile_cross_checks_primary_count_after_typed_semantics(tmp_path: Path) -> None:
    primary = _artifact(
        tmp_path,
        "document.main",
        "primary",
        ".md",
        metadata={"physical_page_count": 4, "ocr_enabled": True, "keep_images": False},
        primary=True,
    )
    pages = [
        _artifact(
            tmp_path,
            f"fragment.page.{page}",
            "auxiliary",
            ".md",
            metadata={
                "fragment_kind": "page",
                "page_index": page,
                "page_count": 3,
                "source_page": page,
                "ocr_status": "success",
            },
        )
        for page in range(1, 4)
    ]

    with pytest.raises(BundleMappingError) as exc_info:
        build_bundle_draft(
            profile="physical_page_ocr",
            output_media_type="text/markdown",
            artifacts=[primary, *pages],
        )

    assert exc_info.value.code == "incomplete_page_sequence"


def test_physical_page_profile_rejects_ocr_off_resource_past_primary_count(tmp_path: Path) -> None:
    primary = _artifact(
        tmp_path,
        "document.main",
        "primary",
        ".md",
        metadata={"physical_page_count": 2, "ocr_enabled": False, "keep_images": True},
        primary=True,
    )
    resource = _artifact(tmp_path, "resource.3", "image", ".png", metadata={"source_page": 3})

    with pytest.raises(BundleMappingError) as exc_info:
        build_bundle_draft(
            profile="physical_page_ocr",
            output_media_type="text/markdown",
            artifacts=[primary, resource],
        )

    assert exc_info.value.code == "invalid_page_range"


def test_physical_page_diagnostics_require_exact_one_bound_warning(tmp_path: Path) -> None:
    primary = _artifact(
        tmp_path,
        "document.main",
        "primary",
        ".md",
        metadata={"physical_page_count": 1, "ocr_enabled": False, "keep_images": True},
        primary=True,
    )
    unresolved = _artifact(tmp_path, "resource.unresolved", "image", ".png")
    draft = build_bundle_draft(
        profile="physical_page_ocr",
        output_media_type="text/markdown",
        artifacts=[primary, unresolved],
    )
    exact = ConversionDiagnostic(
        level="warning",
        message="unresolved",
        code="resource_page_unresolved",
        artifact_id=unresolved.artifact_id,
    )

    validate_physical_page_diagnostics(draft, [exact])
    for invalid in (
        [],
        [ConversionDiagnostic(level="warning", message="unbound", code="resource_page_unresolved")],
        [exact, exact],
        [
            ConversionDiagnostic(
                level="warning",
                message="wrong",
                code="resource_page_unresolved",
                artifact_id=primary.artifact_id,
            )
        ],
    ):
        with pytest.raises(BundleMappingError) as exc_info:
            validate_physical_page_diagnostics(draft, invalid)
        assert exc_info.value.code == "resource_page_diagnostic_mismatch"
