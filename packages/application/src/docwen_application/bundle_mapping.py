"""Route-family policies that map technical runtime artifacts to Bundle v2."""

from __future__ import annotations

from collections import Counter
from collections.abc import Sequence
from typing import Any, Literal, cast

from docwen_core.models import (
    ArtifactBundleValidationError,
    ArtifactManifest,
    BundleDraft,
    BundleDraftArtifact,
    BundleEntry,
    BundleRelation,
    ConversionDiagnostic,
    PageFragmentSemantics,
    PageOcrStatus,
    PageResourceSemantics,
)
from docwen_core.round_trip_sidecar import (
    DOCX_MEDIA_TYPE,
    ROUND_TRIP_SIDECAR_MEDIA_TYPE,
    ROUND_TRIP_SIDECAR_OWNER_METADATA,
    ROUND_TRIP_SIDECAR_SCHEMA,
    ROUND_TRIP_SIDECAR_SCHEMA_METADATA,
)

MARKDOWN_MEDIA_TYPE = "text/markdown"
DOCUMENT_NODE_MANIFEST_MEDIA_TYPE = "application/vnd.docwen.document-node+json"
BundleProfile = Literal[
    "single_document",
    "document_with_round_trip_sidecar",
    "document_with_resources",
    "image_to_markdown",
    "physical_page_ocr",
    "table_resources",
    "frame_images",
    "worksheet_resources",
    "page_images",
    "section_documents",
    "partition_documents",
    "report_resource",
    "image_resource",
]


class BundleMappingError(ValueError):
    """A route returned technical artifacts without valid public semantics."""

    def __init__(
        self,
        category: str,
        code: str,
        message: str,
        *,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.category = category
        self.code = code
        self.details = dict(details or {})


def validate_physical_page_diagnostics(
    draft: BundleDraft,
    diagnostics: Sequence[ConversionDiagnostic],
) -> None:
    """Require exact artifact-bound coverage for unresolved page resources."""
    unresolved_resource_ids = {
        relation.source_artifact_id
        for relation in draft.relations
        if relation.type == "resource_of"
        and relation.role in {"image", "original", "preview"}
        and relation.page_resource is None
    }
    unresolved_diagnostics = [diagnostic for diagnostic in diagnostics if diagnostic.code == "resource_page_unresolved"]
    unbound = any(diagnostic.artifact_id is None for diagnostic in unresolved_diagnostics)
    diagnostic_ids = [
        diagnostic.artifact_id for diagnostic in unresolved_diagnostics if diagnostic.artifact_id is not None
    ]
    counts = Counter(diagnostic_ids)
    if unbound or unresolved_resource_ids != set(diagnostic_ids) or any(count != 1 for count in counts.values()):
        raise BundleMappingError(
            "conversion_failed",
            "resource_page_diagnostic_mismatch",
            "unresolved physical-page resources must have exact artifact-bound diagnostics",
            details={
                "unresolved_artifact_ids": sorted(unresolved_resource_ids),
                "diagnostic_artifact_ids": sorted(diagnostic_ids),
            },
        )


def build_bundle_draft(
    *,
    profile: BundleProfile,
    output_media_type: str,
    artifacts: Sequence[ArtifactManifest],
) -> BundleDraft:
    """Apply one explicit route-family mapping without inferring product storage."""
    artifacts = tuple(artifacts)
    node_manifests = tuple(
        artifact for artifact in artifacts if artifact.media_type == DOCUMENT_NODE_MANIFEST_MEDIA_TYPE
    )
    if len(node_manifests) > 1:
        raise BundleMappingError(
            "conversion_failed",
            "document_node_manifest_count",
            "a conversion may expose at most one document-node manifest",
        )
    node_manifest = node_manifests[0] if node_manifests else None
    artifacts = tuple(artifact for artifact in artifacts if artifact is not node_manifest)

    def finish(draft: BundleDraft) -> BundleDraft:
        return _attach_document_node_manifest(draft, node_manifest)

    if profile == "physical_page_ocr":
        artifact_ids = [artifact.artifact_id for artifact in artifacts]
        if len(artifact_ids) != len(set(artifact_ids)):
            raise BundleMappingError(
                "conversion_failed",
                "duplicate_artifact_id",
                "physical-page capability produced duplicate artifact_id values",
            )
    if profile == "worksheet_resources":
        return finish(
            _build_ordered_entries(
                output_media_type,
                artifacts,
                technical_kinds=("primary", "auxiliary"),
                ordinal_metadata="sheet_index",
                entry_role="worksheet",
                bundle_kind="resource",
            )
        )
    if profile == "page_images":
        return finish(
            _build_ordered_entries(
                output_media_type,
                artifacts,
                technical_kinds=("image",),
                ordinal_metadata="page",
                entry_role="image",
                bundle_kind="resource",
                one_based_ordinal=True,
            )
        )
    if profile == "frame_images":
        return finish(
            _build_ordered_entries(
                output_media_type,
                artifacts,
                technical_kinds=("primary", "auxiliary"),
                ordinal_metadata="page_index",
                entry_role="image",
                bundle_kind="resource",
            )
        )
    if profile == "table_resources":
        return finish(
            _build_ordered_entries(
                output_media_type,
                artifacts,
                technical_kinds=("primary",),
                ordinal_metadata="table_index",
                entry_role="supplementary",
                bundle_kind="resource",
            )
        )
    if profile == "section_documents":
        return finish(
            _build_ordered_entries(
                output_media_type,
                artifacts,
                technical_kinds=("primary",),
                ordinal_metadata="page",
                entry_role="section",
                bundle_kind="document",
                one_based_ordinal=True,
            )
        )
    if profile == "partition_documents":
        return finish(_build_partition_documents(output_media_type, artifacts))

    preferred_artifacts = [artifact for artifact in artifacts if artifact.is_primary]
    if len(preferred_artifacts) != 1:
        raise BundleMappingError(
            "conversion_failed",
            "preferred_output_count",
            "capability output must identify exactly one preferred artifact",
            details={"preferred_artifact_count": len(preferred_artifacts)},
        )
    preferred = preferred_artifacts[0]
    if preferred.kind != "primary" or preferred.media_type != output_media_type:
        raise BundleMappingError(
            "conversion_failed",
            "preferred_output_semantics_invalid",
            "preferred artifact kind or media type does not match the capability contract",
        )
    if profile == "report_resource":
        if len(artifacts) != 1:
            raise BundleMappingError(
                "conversion_failed",
                "unexpected_output_shape",
                "report capability must produce exactly one deliverable artifact",
                details={"artifact_count": len(artifacts)},
            )
        return finish(
            BundleDraft(
                artifacts=(_draft_artifact(preferred, "resource"),),
                entries=(BundleEntry(preferred.artifact_id, "supplementary", 0, True),),
            )
        )
    if profile == "image_resource":
        if len(artifacts) != 1:
            raise BundleMappingError(
                "conversion_failed",
                "unexpected_output_shape",
                "image merge capability must produce exactly one deliverable artifact",
                details={"artifact_count": len(artifacts)},
            )
        return finish(
            BundleDraft(
                artifacts=(_draft_artifact(preferred, "resource"),),
                entries=(BundleEntry(preferred.artifact_id, "image", 0, True),),
            )
        )
    if profile == "single_document":
        if len(artifacts) != 1:
            raise BundleMappingError(
                "conversion_failed",
                "unexpected_output_shape",
                "single-document capability must produce exactly one deliverable artifact",
                details={"artifact_count": len(artifacts)},
            )
        return finish(
            BundleDraft(
                artifacts=(_draft_artifact(preferred, "document"),),
                entries=(BundleEntry(preferred.artifact_id, "primary", 0, True),),
            )
        )
    if profile == "document_with_round_trip_sidecar":
        return finish(_build_document_with_round_trip_sidecar(preferred, artifacts))
    if profile == "image_to_markdown":
        return finish(_build_image_to_markdown(preferred, artifacts))
    if profile == "physical_page_ocr":
        return finish(_build_physical_page_ocr(preferred, artifacts))
    if profile != "document_with_resources":
        raise BundleMappingError(
            "internal",
            "bundle_profile_unknown",
            "capability uses an unknown Artifact Bundle mapping profile",
        )
    return finish(_build_document_with_resources(preferred, artifacts))


def _validate_mapped_draft(draft: BundleDraft) -> None:
    try:
        from docwen_core.models import validate_artifact_bundle_draft

        validate_artifact_bundle_draft(draft)
    except ArtifactBundleValidationError as exc:
        raise BundleMappingError("conversion_failed", exc.code, str(exc)) from exc


def _build_physical_page_ocr(
    preferred: ArtifactManifest,
    artifacts: tuple[ArtifactManifest, ...],
) -> BundleDraft:
    physical_page_count = preferred.metadata.get("physical_page_count")
    ocr_enabled = preferred.metadata.get("ocr_enabled")
    keep_images = preferred.metadata.get("keep_images")
    page_fragments = [
        artifact
        for artifact in artifacts
        if artifact.kind == "auxiliary"
        and artifact.media_type == MARKDOWN_MEDIA_TYPE
        and artifact.metadata.get("fragment_kind") == "page"
    ]
    resources = [
        artifact for artifact in artifacts if artifact.kind == "image" and artifact.media_type.startswith("image/")
    ]
    recognized_objects = {id(preferred), *(id(artifact) for artifact in page_fragments), *(id(a) for a in resources)}
    if len(recognized_objects) != len(artifacts):
        unknown = sorted(artifact.artifact_id for artifact in artifacts if id(artifact) not in recognized_objects)
        raise BundleMappingError(
            "conversion_failed",
            "artifact_semantics_unknown",
            "physical-page capability produced an artifact without closed page semantics",
            details={"artifact_ids": unknown},
        )

    draft_artifacts: list[BundleDraftArtifact] = [_draft_artifact(preferred, "document")]
    relations: list[BundleRelation] = []
    for fragment in page_fragments:
        page_index = fragment.metadata.get("page_index")
        fragment_page_count = fragment.metadata.get("page_count")
        source_page = fragment.metadata.get("source_page")
        ocr_status = fragment.metadata.get("ocr_status")
        draft_artifacts.append(_draft_artifact(fragment, "fragment"))
        relations.append(
            BundleRelation(
                type="fragment_of",
                source_artifact_id=fragment.artifact_id,
                target_artifact_id=preferred.artifact_id,
                role="ocr_page",
                ordinal=page_index - 1 if isinstance(page_index, int) and not isinstance(page_index, bool) else 0,
                page_fragment=PageFragmentSemantics(
                    page_index=cast(int, page_index),
                    page_count=cast(int, fragment_page_count),
                    ocr_status=cast(PageOcrStatus, ocr_status),
                    source_page=cast(int, source_page),
                ),
            )
        )

    page_artifact_by_source = {
        source_page: fragment
        for fragment in page_fragments
        if isinstance((source_page := fragment.metadata.get("source_page")), int) and not isinstance(source_page, bool)
    }

    for resource in resources:
        draft_artifacts.append(_draft_artifact(resource, "resource"))
        raw_source_page = resource.metadata.get("source_page")
        page_resource = None
        target_artifact_id = preferred.artifact_id
        if raw_source_page is not None:
            source_page = cast(int, raw_source_page)
            page_resource = PageResourceSemantics(source_page=source_page)
            if page_fragments and isinstance(raw_source_page, int) and not isinstance(raw_source_page, bool):
                target = page_artifact_by_source.get(raw_source_page)
                if target is not None:
                    target_artifact_id = target.artifact_id
        relations.append(
            BundleRelation(
                type="resource_of",
                source_artifact_id=resource.artifact_id,
                target_artifact_id=target_artifact_id,
                role="image",
                page_resource=page_resource,
            )
        )

    draft = BundleDraft(
        artifacts=tuple(draft_artifacts),
        entries=(BundleEntry(preferred.artifact_id, "primary", 0, True),),
        relations=tuple(relations),
    )
    _validate_mapped_draft(draft)
    if (
        not isinstance(physical_page_count, int)
        or isinstance(physical_page_count, bool)
        or physical_page_count < 1
        or not isinstance(ocr_enabled, bool)
        or not isinstance(keep_images, bool)
    ):
        raise BundleMappingError(
            "conversion_failed",
            "physical_page_metadata_invalid",
            "physical-page primary artifact must declare positive physical_page_count and boolean ocr_enabled/keep_images",
        )
    if ocr_enabled and len(page_fragments) != physical_page_count:
        raise BundleMappingError(
            "conversion_failed",
            "incomplete_page_sequence",
            "OCR-enabled physical-page output must contain one fragment per page",
        )
    if not ocr_enabled and page_fragments:
        raise BundleMappingError(
            "conversion_failed",
            "unexpected_page_semantics",
            "OCR-disabled physical-page output must not contain page fragments",
        )
    if not keep_images and resources:
        raise BundleMappingError(
            "conversion_failed",
            "unexpected_page_semantics",
            "image-export-disabled physical-page output must not contain image resources",
        )
    for resource in resources:
        source_page = resource.metadata.get("source_page")
        if isinstance(source_page, int) and not isinstance(source_page, bool) and source_page > physical_page_count:
            raise BundleMappingError(
                "conversion_failed",
                "invalid_page_range",
                "physical-page resource source_page exceeds the primary page count",
            )
    return draft


def _build_image_to_markdown(
    preferred: ArtifactManifest,
    artifacts: tuple[ArtifactManifest, ...],
) -> BundleDraft:
    retained_images = [
        artifact for artifact in artifacts if artifact.kind == "image" and artifact.media_type.startswith("image/")
    ]
    ocr_fragments = [
        artifact
        for artifact in artifacts
        if artifact.kind == "auxiliary"
        and artifact.media_type == MARKDOWN_MEDIA_TYPE
        and artifact.metadata.get("ocr") is True
    ]
    if len(artifacts) != 3 or len(retained_images) != 1 or len(ocr_fragments) != 1:
        raise BundleMappingError(
            "conversion_failed",
            "image_markdown_output_shape_invalid",
            "image OCR capability must produce one document, one retained image, and one OCR fragment",
        )
    retained_image = retained_images[0]
    ocr_fragment = ocr_fragments[0]
    return BundleDraft(
        artifacts=(
            _draft_artifact(preferred, "document"),
            _draft_artifact(retained_image, "resource"),
            _draft_artifact(ocr_fragment, "fragment"),
        ),
        entries=(BundleEntry(preferred.artifact_id, "primary", 0, True),),
        relations=(
            BundleRelation(
                type="resource_of",
                source_artifact_id=retained_image.artifact_id,
                target_artifact_id=preferred.artifact_id,
                role="original",
                ordinal=0,
            ),
            BundleRelation(
                type="fragment_of",
                source_artifact_id=ocr_fragment.artifact_id,
                target_artifact_id=preferred.artifact_id,
                role="ocr_text",
                ordinal=0,
            ),
            BundleRelation(
                type="derived_from",
                source_artifact_id=ocr_fragment.artifact_id,
                target_artifact_id=retained_image.artifact_id,
                role="source",
            ),
        ),
    )


def _build_document_with_resources(
    preferred: ArtifactManifest,
    artifacts: tuple[ArtifactManifest, ...],
) -> BundleDraft:
    draft_artifacts = [_draft_artifact(preferred, "document")]
    relations: list[BundleRelation] = []
    attachment_ordinal = 0
    fragment_ordinal = 0
    resource_ordinal = 0
    for artifact in artifacts:
        if artifact is preferred:
            continue
        if artifact.kind == "image" and artifact.media_type.startswith("image/"):
            draft_artifacts.append(_draft_artifact(artifact, "resource"))
            relations.append(
                BundleRelation(
                    type="resource_of",
                    source_artifact_id=artifact.artifact_id,
                    target_artifact_id=preferred.artifact_id,
                    role="image",
                    ordinal=resource_ordinal,
                )
            )
            resource_ordinal += 1
            continue
        if (
            artifact.kind == "auxiliary"
            and artifact.media_type == MARKDOWN_MEDIA_TYPE
            and artifact.metadata.get("source_kind") == "gongwen_attachment"
        ):
            raw_ordinal = artifact.metadata.get("attachment_ordinal")
            ordinal = raw_ordinal - 1 if isinstance(raw_ordinal, int) and raw_ordinal > 0 else attachment_ordinal
            draft_artifacts.append(_draft_artifact(artifact, "document"))
            relations.append(
                BundleRelation(
                    type="attachment_of",
                    source_artifact_id=artifact.artifact_id,
                    target_artifact_id=preferred.artifact_id,
                    role="attachment",
                    ordinal=ordinal,
                )
            )
            attachment_ordinal = max(attachment_ordinal, ordinal + 1)
            continue
        if (
            artifact.kind == "auxiliary"
            and artifact.media_type == MARKDOWN_MEDIA_TYPE
            and artifact.metadata.get("ocr") is True
        ):
            draft_artifacts.append(_draft_artifact(artifact, "fragment"))
            relations.append(
                BundleRelation(
                    type="fragment_of",
                    source_artifact_id=artifact.artifact_id,
                    target_artifact_id=preferred.artifact_id,
                    role="ocr_text",
                    ordinal=fragment_ordinal,
                )
            )
            fragment_ordinal += 1
            continue
        raise BundleMappingError(
            "conversion_failed",
            "artifact_semantics_unknown",
            "runtime produced an artifact that the capability has not assigned consumer-neutral semantics",
            details={
                "artifact_id": artifact.artifact_id,
                "technical_kind": artifact.kind,
                "media_type": artifact.media_type,
            },
        )
    return BundleDraft(
        artifacts=tuple(draft_artifacts),
        entries=(BundleEntry(preferred.artifact_id, "primary", 0, True),),
        relations=tuple(relations),
    )


def _build_document_with_round_trip_sidecar(
    preferred: ArtifactManifest,
    artifacts: tuple[ArtifactManifest, ...],
) -> BundleDraft:
    sidecars = [
        artifact
        for artifact in artifacts
        if artifact.kind == "auxiliary" and artifact.media_type == ROUND_TRIP_SIDECAR_MEDIA_TYPE
    ]
    if len(artifacts) != 2 or len(sidecars) != 1:
        raise BundleMappingError(
            "conversion_failed",
            "round_trip_sidecar_output_shape_invalid",
            "resolved Markdown to DOCX must produce one document and one round-trip sidecar",
            details={"artifact_count": len(artifacts), "sidecar_count": len(sidecars)},
        )
    sidecar = sidecars[0]
    if (
        sidecar.metadata.get(ROUND_TRIP_SIDECAR_SCHEMA_METADATA) != ROUND_TRIP_SIDECAR_SCHEMA
        or sidecar.metadata.get(ROUND_TRIP_SIDECAR_OWNER_METADATA) != preferred.artifact_id
        or sidecar.suggested_name != f"{preferred.suggested_name}.docwen"
        or preferred.media_type != DOCX_MEDIA_TYPE
    ):
        raise BundleMappingError(
            "conversion_failed",
            "round_trip_sidecar_owner_invalid",
            "round-trip sidecar ownership metadata or suggested name is invalid",
        )
    draft = BundleDraft(
        artifacts=(
            _draft_artifact(preferred, "document"),
            _draft_artifact(sidecar, "resource"),
        ),
        entries=(BundleEntry(preferred.artifact_id, "primary", 0, True),),
        relations=(
            BundleRelation(
                type="resource_of",
                source_artifact_id=sidecar.artifact_id,
                target_artifact_id=preferred.artifact_id,
                role="manifest",
                ordinal=0,
            ),
        ),
    )
    _validate_mapped_draft(draft)
    return draft


def _build_ordered_entries(
    output_media_type: str,
    artifacts: tuple[ArtifactManifest, ...],
    *,
    technical_kinds: tuple[str, ...],
    ordinal_metadata: str,
    entry_role: Literal["supplementary", "worksheet", "image", "section"],
    bundle_kind: Literal["document", "resource"],
    one_based_ordinal: bool = False,
) -> BundleDraft:
    if not artifacts:
        raise BundleMappingError(
            "conversion_failed",
            "empty_output_bundle",
            "multi-output capability did not produce an artifact",
        )
    entries: list[BundleEntry] = []
    draft_artifacts: list[BundleDraftArtifact] = []
    preferred_count = 0
    ordinals: set[int] = set()
    for artifact in artifacts:
        raw_ordinal = artifact.metadata.get(ordinal_metadata)
        if not isinstance(raw_ordinal, int) or isinstance(raw_ordinal, bool):
            raise BundleMappingError(
                "conversion_failed",
                "artifact_order_missing",
                "ordered multi-output artifact is missing its route-defined ordinal",
                details={"artifact_id": artifact.artifact_id, "metadata_key": ordinal_metadata},
            )
        ordinal = raw_ordinal - 1 if one_based_ordinal else raw_ordinal
        if (
            ordinal < 0
            or ordinal in ordinals
            or artifact.kind not in technical_kinds
            or artifact.media_type != output_media_type
        ):
            raise BundleMappingError(
                "conversion_failed",
                "artifact_order_or_type_invalid",
                "ordered multi-output artifact violates the capability contract",
                details={"artifact_id": artifact.artifact_id},
            )
        ordinals.add(ordinal)
        preferred_count += int(artifact.is_primary)
        draft_artifacts.append(_draft_artifact(artifact, bundle_kind))
        entries.append(BundleEntry(artifact.artifact_id, entry_role, ordinal, artifact.is_primary))
    if ordinals != set(range(len(artifacts))) or preferred_count != 1:
        raise BundleMappingError(
            "conversion_failed",
            "artifact_sequence_invalid",
            "ordered multi-output artifacts must form one contiguous sequence with one preferred entry",
        )
    ordered = sorted(zip(entries, draft_artifacts, strict=True), key=lambda pair: pair[0].ordinal)
    return BundleDraft(
        artifacts=tuple(artifact for _, artifact in ordered),
        entries=tuple(entry for entry, _ in ordered),
    )


def _build_partition_documents(
    output_media_type: str,
    artifacts: tuple[ArtifactManifest, ...],
) -> BundleDraft:
    if not artifacts:
        raise BundleMappingError(
            "conversion_failed",
            "empty_output_bundle",
            "PDF partition capability did not produce an artifact",
        )
    entries: list[BundleEntry] = []
    draft_artifacts: list[BundleDraftArtifact] = []
    preferred_count = 0
    seen_pages: set[int] = set()
    for ordinal, artifact in enumerate(artifacts):
        pages = artifact.metadata.get("pages")
        if (
            artifact.kind != "primary"
            or artifact.media_type != output_media_type
            or not isinstance(pages, list)
            or not pages
            or any(not isinstance(page, int) or isinstance(page, bool) or page < 1 for page in pages)
            or seen_pages.intersection(pages)
        ):
            raise BundleMappingError(
                "conversion_failed",
                "partition_output_shape_invalid",
                "PDF partition artifacts must declare disjoint one-based page sets",
                details={"artifact_id": artifact.artifact_id},
            )
        seen_pages.update(pages)
        preferred_count += int(artifact.is_primary)
        draft_artifacts.append(_draft_artifact(artifact, "document"))
        entries.append(BundleEntry(artifact.artifact_id, "section", ordinal, artifact.is_primary))
    if preferred_count != 1:
        raise BundleMappingError(
            "conversion_failed",
            "preferred_output_count",
            "PDF partition capability must identify exactly one preferred artifact",
        )
    return BundleDraft(artifacts=tuple(draft_artifacts), entries=tuple(entries))


def _attach_document_node_manifest(
    draft: BundleDraft,
    manifest: ArtifactManifest | None,
) -> BundleDraft:
    if manifest is None:
        return draft
    preferred = [entry for entry in draft.entries if entry.preferred]
    if len(preferred) != 1:
        raise BundleMappingError(
            "conversion_failed",
            "document_node_manifest_owner",
            "document-node manifest requires exactly one preferred owner",
        )
    target_id = preferred[0].artifact_id
    used_ordinals = {
        relation.ordinal
        for relation in draft.relations
        if relation.type == "resource_of" and relation.target_artifact_id == target_id and relation.ordinal is not None
    }
    ordinal = 0
    while ordinal in used_ordinals:
        ordinal += 1
    return BundleDraft(
        artifacts=(*draft.artifacts, _draft_artifact(manifest, "resource")),
        entries=draft.entries,
        relations=(
            *draft.relations,
            BundleRelation(
                type="resource_of",
                source_artifact_id=manifest.artifact_id,
                target_artifact_id=target_id,
                role="manifest",
                ordinal=ordinal,
            ),
        ),
    )


def _draft_artifact(
    artifact: ArtifactManifest,
    kind: Literal["document", "fragment", "resource"],
) -> BundleDraftArtifact:
    return BundleDraftArtifact(
        artifact_id=artifact.artifact_id,
        kind=kind,
        path=artifact.staging_path,
        suggested_name=artifact.suggested_name,
        media_type=artifact.media_type,
        logical_path=artifact.logical_path,
    )


__all__ = ["BundleMappingError", "BundleProfile", "build_bundle_draft"]
