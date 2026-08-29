"""Permanent contracts for the plan-first DocWen ConversionService."""

from __future__ import annotations

import hashlib
import json
import os
from dataclasses import replace
from pathlib import Path
from typing import Any, Literal

import pytest

from docwen_application.conversion_service import (
    CSV_MEDIA_TYPE,
    DOCX_MEDIA_TYPE,
    DOCX_TO_MARKDOWN_CAPABILITY_ID,
    IMAGES_MERGE_TO_TIFF_CAPABILITY_ID,
    JSON_MEDIA_TYPE,
    MARKDOWN_MEDIA_TYPE,
    MARKDOWN_NUMBERING_CAPABILITY_ID,
    MARKDOWN_TABLES_TO_CSV_CAPABILITY_ID,
    MARKDOWN_TO_DOCX_CAPABILITY_ID,
    MARKDOWN_TO_XLSX_CAPABILITY_ID,
    MARKDOWN_VALIDATE_CAPABILITY_ID,
    OFD_MEDIA_TYPE,
    OFD_TO_MARKDOWN_CAPABILITY_ID,
    PDF_MEDIA_TYPE,
    PDF_MERGE_CAPABILITY_ID,
    PDF_SPLIT_CUSTOM_CAPABILITY_ID,
    PDF_SPLIT_EVERY_PAGE_CAPABILITY_ID,
    PDF_TO_MARKDOWN_CAPABILITY_ID,
    PDF_TO_PNG_CAPABILITY_ID,
    PNG_MEDIA_TYPE,
    PNG_TO_OCR_MARKDOWN_CAPABILITY_ID,
    SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE,
    TIFF_FRAMES_TO_PNG_CAPABILITY_ID,
    TIFF_MEDIA_TYPE,
    TIFF_TO_MARKDOWN_CAPABILITY_ID,
    XLSX_MEDIA_TYPE,
    XLSX_MERGE_TABLES_CAPABILITY_ID,
    XLSX_TO_CSV_CAPABILITY_ID,
    XLSX_TO_MARKDOWN_CAPABILITY_ID,
    XPS_MEDIA_TYPE,
    XPS_TO_MARKDOWN_CAPABILITY_ID,
    ConversionPlanRequest,
    ConversionService,
    ConversionServiceError,
    LocalInputHandle,
    StagingOutputTarget,
)
from docwen_core.models import (
    NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
    RESOLVED_DOCUMENT_MEDIA_TYPE,
    ArtifactBundle,
    ArtifactManifest,
    BundleArtifact,
    BundleProducer,
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
    canonicalize_numbering_plan,
)
from docwen_core.paths import filesystem_path
from docwen_core.round_trip_sidecar import (
    ROUND_TRIP_SIDECAR_MEDIA_TYPE,
    ROUND_TRIP_SIDECAR_OWNER_METADATA,
    ROUND_TRIP_SIDECAR_SCHEMA,
    ROUND_TRIP_SIDECAR_SCHEMA_METADATA,
)

pytestmark = pytest.mark.integration

_EXPECTED_DOCUMENT_SEMANTICS_MACHINE_LIMITATIONS = (
    {
        "severity": "warning",
        "code": "document_semantics.citation_processor_unavailable",
        "message": (
            "DocWen does not run a CSL citation processor or accept citation_style inputs in Machine v1; "
            "Markdown citation keys remain literal."
        ),
    },
    {
        "severity": "warning",
        "code": "document_semantics.v1_scope",
        "message": (
            "Document semantics v1 excludes CSL processing, composite or range citation semantics, "
            "custom citation display, and PDF semantic round trips."
        ),
    },
)

_EXPECTED_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS = (
    {
        "severity": "warning",
        "code": "resolved_document.provider_owned_semantics",
        "message": (
            "DocWen consumes already-resolved targets, citations, resources, and numbering facts; it does not "
            "scan a Workspace, run a citation resolver, or infer numbering from authored text."
        ),
    },
)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _input_kind(media_type: str) -> Literal["document", "resource"]:
    return "document" if media_type in {MARKDOWN_MEDIA_TYPE, DOCX_MEDIA_TYPE} else "resource"


def _resolved_numbering_inputs(source: Path) -> tuple[Path, Path]:
    authored_markdown = source.read_text(encoding="utf-8")
    plan_value = {"heading_definitions": [], "heading_instances": [], "targets": []}
    plan_sha256 = hashlib.sha256(canonicalize_numbering_plan(plan_value)).hexdigest()
    source_sha256 = hashlib.sha256(authored_markdown.encode()).hexdigest()
    document_value = {
        "$schema": "urn:docwen:schema:resolved-document:v1",
        "schema": "docwen.resolved_document.v1",
        "input_id": "document-1",
        "source_sha256": source_sha256,
        "plan_sha256": plan_sha256,
        "document": {
            "authored_markdown": authored_markdown,
            "targets": [],
            "references": [],
            "resource_occurrences": [],
            "citations": [],
            "resources": [],
        },
    }
    plan_envelope = {
        "$schema": "urn:docwen:schema:numbering-export-plan:v1",
        "schema": "docwen.numbering_export_plan.v1",
        "input_id": "document-1",
        "source_sha256": source_sha256,
        "plan_sha256": plan_sha256,
        "plan": plan_value,
    }
    neutral_path = source.with_name(f"{source.name}.resolved.json")
    plan_path = source.with_name(f"{source.name}.numbering.json")
    neutral_path.write_text(json.dumps(document_value, separators=(",", ":")), encoding="utf-8")
    plan_path.write_text(json.dumps(plan_envelope, separators=(",", ":")), encoding="utf-8")
    return neutral_path, plan_path


def _refingerprint(handle: LocalInputHandle) -> LocalInputHandle:
    path = Path(handle.path)
    return replace(handle, size_bytes=path.stat().st_size, sha256=_sha256(path))


class _Controller:
    has_runtime = True

    _ROUTE_IDS = (
        "docwen_plugin_markdown:markdown:docx:convert",
        "docwen_plugin_markdown:markdown:xlsx:convert",
        "docwen_plugin_document:docx:md:convert",
        "docwen_plugin_layout:pdf:md:convert",
        "docwen_plugin_layout:ofd:md:convert",
        "docwen_plugin_layout:xps:md:convert",
        "docwen_plugin_spreadsheet:xlsx:csv:convert",
        "docwen_plugin_spreadsheet:xlsx:md:convert",
        "docwen_plugin_layout:pdf:png:convert",
        "docwen_plugin_layout:pdf:pdf:split_pdf",
        "docwen_plugin_image:image:md:convert",
        "docwen_plugin_markdown:markdown:csv:convert",
        "docwen_plugin_image:image:png:convert",
        "docwen_plugin_proofread:markdown:markdown:validate",
        "docwen_plugin_markdown:markdown:md:process_md_numbering",
        "docwen_plugin_layout:pdf:pdf:merge_pdfs",
        "docwen_plugin_spreadsheet:spreadsheet:xlsx:merge_tables",
        "docwen_plugin_image:image:tif:merge_images_to_tiff",
    )

    def __init__(self) -> None:
        self.cancelled: list[str] = []
        self.released: list[str] = []
        self.requests: list[Any] = []

    def prepare_execution_cancellation(self, request: Any, *, batch: bool = False) -> object:
        assert batch is False
        return object()

    def describe_runtime_capabilities(self) -> dict[str, Any]:
        return {
            "gates": [],
            "sources": [
                {
                    "routes": [
                        {
                            "id": route_id,
                            "available": True,
                            "required_capabilities": [],
                            "optional_capabilities": [],
                            "missing_required_capabilities": [],
                            "missing_optional_capabilities": [],
                            "limitations": [],
                        }
                        for route_id in self._ROUTE_IDS
                    ]
                }
            ],
        }

    def release_execution_cancellation(self, task_id: str, reservation: object) -> None:
        self.released.append(task_id)

    def execute_single(self, request: Any) -> ConversionResult:
        self.requests.append(request)
        output_contract = {
            "md": (".md", MARKDOWN_MEDIA_TYPE),
            "markdown": (".json", JSON_MEDIA_TYPE),
            "docx": (".docx", DOCX_MEDIA_TYPE),
            "xlsx": (".xlsx", XLSX_MEDIA_TYPE),
            "pdf": (".pdf", PDF_MEDIA_TYPE),
            "tif": (".tiff", TIFF_MEDIA_TYPE),
        }
        suffix, media_type = output_contract[request.target_format]
        output = Path(request.output_policy.output_dir) / f"converted{suffix}"
        output.write_bytes(b"# markdown fixture\n" if suffix == ".md" else b"fixture")
        primary = ArtifactManifest(
            artifact_id="artifact.primary",
            kind="primary",
            staging_path=str(output),
            suggested_name=output.name,
            media_type=media_type,
            is_primary=True,
        )
        artifacts = [primary]
        if request.target_format == "docx":
            sidecar = Path(f"{output}.docwen")
            sidecar.write_bytes(b"sidecar fixture")
            artifacts.append(
                ArtifactManifest(
                    artifact_id="artifact.sidecar",
                    kind="auxiliary",
                    staging_path=str(sidecar),
                    suggested_name=f"{output.name}.docwen",
                    media_type=ROUND_TRIP_SIDECAR_MEDIA_TYPE,
                    metadata={
                        ROUND_TRIP_SIDECAR_SCHEMA_METADATA: ROUND_TRIP_SIDECAR_SCHEMA,
                        ROUND_TRIP_SIDECAR_OWNER_METADATA: primary.artifact_id,
                    },
                )
            )
        return ConversionResult(
            task_id=request.request_id,
            success=True,
            artifacts=artifacts,
            metrics=ConversionMetrics(
                input_bytes=request.input_refs[0].size_bytes,
                output_bytes=sum(Path(item.staging_path).stat().st_size for item in artifacts),
            ),
        )

    def execute_aggregate(self, request: Any, action_name: str) -> ConversionResult:
        assert action_name in {"merge_pdfs", "merge_tables", "merge_images_to_tiff"}
        return self.execute_single(request)

    def cancel(self, task_id: str) -> None:
        self.cancelled.append(task_id)


class _Committer:
    def commit(self, *, task_id: str, staging_root: str, draft: Any) -> ArtifactBundle:
        return ArtifactBundle(
            bundle_id="bundle.test",
            task_id=task_id,
            producer=BundleProducer(product_version="0.9.0"),
            artifacts=tuple(
                BundleArtifact(
                    artifact_id=artifact.artifact_id,
                    kind=artifact.kind,
                    locator=source.relative_to(staging_root).as_posix(),
                    suggested_name=artifact.suggested_name,
                    media_type=artifact.media_type,
                    size_bytes=source.stat().st_size,
                    sha256=_sha256(source),
                )
                for artifact in draft.artifacts
                for source in (Path(artifact.path),)
            ),
            entries=draft.entries,
            relations=draft.relations,
        )

    def discard(self, *, staging_root: str, artifact_paths: list[str]) -> None:
        for artifact_path in artifact_paths:
            path = Path(artifact_path)
            if path.is_file() and path.is_relative_to(staging_root):
                path.unlink()


def _request(
    input_path: Path,
    staging_root: Path,
    *,
    capability_id: str = MARKDOWN_TO_DOCX_CAPABILITY_ID,
    media_type: str = MARKDOWN_MEDIA_TYPE,
    options: dict[str, Any] | None = None,
) -> ConversionPlanRequest:
    if capability_id == MARKDOWN_TO_DOCX_CAPABILITY_ID:
        neutral_path, plan_path = _resolved_numbering_inputs(input_path)
        return ConversionPlanRequest(
            capability_id=capability_id,
            inputs=(
                LocalInputHandle(
                    input_id="input.neutral",
                    path=str(neutral_path),
                    kind="document",
                    role="neutral_document",
                    logical_path="document.resolved.json",
                    media_type=RESOLVED_DOCUMENT_MEDIA_TYPE,
                    size_bytes=neutral_path.stat().st_size,
                    sha256=_sha256(neutral_path),
                ),
                LocalInputHandle(
                    input_id="input.plan",
                    path=str(plan_path),
                    kind="resource",
                    role="numbering_export_plan",
                    logical_path="numbering-plan.json",
                    media_type=NUMBERING_EXPORT_PLAN_MEDIA_TYPE,
                    size_bytes=plan_path.stat().st_size,
                    sha256=_sha256(plan_path),
                ),
            ),
            output=StagingOutputTarget(staging_root=str(staging_root)),
            options=dict(options or {}),
        )
    return ConversionPlanRequest(
        capability_id=capability_id,
        inputs=(
            LocalInputHandle(
                input_id="input.1",
                path=str(input_path),
                kind=_input_kind(media_type),
                role="source",
                logical_path=input_path.name,
                media_type=media_type,
                size_bytes=input_path.stat().st_size,
                sha256=_sha256(input_path),
            ),
        ),
        output=StagingOutputTarget(staging_root=str(staging_root)),
        options=dict(options or {}),
    )


def _request_many(
    input_paths: list[Path],
    staging_root: Path,
    *,
    capability_id: str,
    media_types: list[str],
    options: dict[str, Any] | None = None,
) -> ConversionPlanRequest:
    return ConversionPlanRequest(
        capability_id=capability_id,
        inputs=tuple(
            LocalInputHandle(
                input_id=f"input.{index}",
                path=str(input_path),
                kind=_input_kind(media_type),
                role="source",
                logical_path=f"{index}-{input_path.name}",
                media_type=media_type,
                size_bytes=input_path.stat().st_size,
                sha256=_sha256(input_path),
            )
            for index, (input_path, media_type) in enumerate(zip(input_paths, media_types, strict=True), start=1)
        ),
        output=StagingOutputTarget(staging_root=str(staging_root)),
        options=dict(options or {}),
    )


__all__ = (
    "CSV_MEDIA_TYPE",
    "DOCX_MEDIA_TYPE",
    "DOCX_TO_MARKDOWN_CAPABILITY_ID",
    "IMAGES_MERGE_TO_TIFF_CAPABILITY_ID",
    "MARKDOWN_MEDIA_TYPE",
    "MARKDOWN_NUMBERING_CAPABILITY_ID",
    "MARKDOWN_TABLES_TO_CSV_CAPABILITY_ID",
    "MARKDOWN_TO_DOCX_CAPABILITY_ID",
    "MARKDOWN_TO_XLSX_CAPABILITY_ID",
    "MARKDOWN_VALIDATE_CAPABILITY_ID",
    "NUMBERING_EXPORT_PLAN_MEDIA_TYPE",
    "OFD_MEDIA_TYPE",
    "OFD_TO_MARKDOWN_CAPABILITY_ID",
    "PDF_MEDIA_TYPE",
    "PDF_MERGE_CAPABILITY_ID",
    "PDF_SPLIT_CUSTOM_CAPABILITY_ID",
    "PDF_SPLIT_EVERY_PAGE_CAPABILITY_ID",
    "PDF_TO_MARKDOWN_CAPABILITY_ID",
    "PDF_TO_PNG_CAPABILITY_ID",
    "PNG_MEDIA_TYPE",
    "PNG_TO_OCR_MARKDOWN_CAPABILITY_ID",
    "RESOLVED_DOCUMENT_MEDIA_TYPE",
    "SEMANTIC_BIBLIOGRAPHY_MEDIA_TYPE",
    "TIFF_FRAMES_TO_PNG_CAPABILITY_ID",
    "TIFF_MEDIA_TYPE",
    "TIFF_TO_MARKDOWN_CAPABILITY_ID",
    "XLSX_MEDIA_TYPE",
    "XLSX_MERGE_TABLES_CAPABILITY_ID",
    "XLSX_TO_CSV_CAPABILITY_ID",
    "XLSX_TO_MARKDOWN_CAPABILITY_ID",
    "XPS_MEDIA_TYPE",
    "XPS_TO_MARKDOWN_CAPABILITY_ID",
    "_EXPECTED_DOCUMENT_SEMANTICS_MACHINE_LIMITATIONS",
    "_EXPECTED_RESOLVED_DOCUMENT_MACHINE_LIMITATIONS",
    "Any",
    "ArtifactBundle",
    "ArtifactManifest",
    "ConversionDiagnostic",
    "ConversionErrorInfo",
    "ConversionPlanRequest",
    "ConversionResult",
    "ConversionService",
    "ConversionServiceError",
    "LocalInputHandle",
    "Path",
    "StagingOutputTarget",
    "_Committer",
    "_Controller",
    "_refingerprint",
    "_request",
    "_request_many",
    "_sha256",
    "canonicalize_numbering_plan",
    "filesystem_path",
    "hashlib",
    "json",
    "os",
    "pytest",
    "pytestmark",
    "replace",
)
