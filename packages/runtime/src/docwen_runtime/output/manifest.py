"""Deterministic, privacy-bounded conversion sidecar manifests."""

from __future__ import annotations

import hashlib
import json
import tempfile
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

from docwen_core.models.artifact import ARTIFACT_KIND_MANIFEST, ArtifactManifest
from docwen_core.models.conversion_manifest import ConversionManifestContext
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_core.models.result import ConversionDiagnostic, ConversionResult
from docwen_runtime.output.finalizer import OutputFinalizer

_MANIFEST_SCHEMA_VERSION = "1.0"
_MANIFEST_MEDIA_TYPE = "application/json"
_WRITE_FAILURE_CODE = "OUTPUT_MANIFEST_WRITE_FAILED"


@dataclass(frozen=True, slots=True)
class OutputManifestDocument:
    """Allowlisted terminal facts serialized to one canonical JSON document."""

    request_id_sha256: str
    status: str
    target_format: str
    action_name: str
    inputs: tuple[tuple[str, str, str], ...]
    preconversion: tuple[tuple[int, str, str, str, str, str], ...]
    artifacts: tuple[tuple[str, str, str, bool], ...]
    diagnostics: tuple[tuple[str, str], ...]
    error: tuple[str, str] | None

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": _MANIFEST_SCHEMA_VERSION,
            "request": {
                "id_sha256": self.request_id_sha256,
                "target_format": self.target_format,
                "action_name": self.action_name,
            },
            "status": self.status,
            "inputs": [
                {"path": path, "format": format_name, "category": category}
                for path, format_name, category in self.inputs
            ],
            "preconversion": [
                {
                    "input_index": input_index,
                    "source_format": source_format,
                    "target_format": target_format,
                    "status": status,
                    "backend": backend,
                    "diagnostic_code": diagnostic_code,
                }
                for input_index, source_format, target_format, status, backend, diagnostic_code in self.preconversion
            ],
            "artifacts": [
                {
                    "kind": kind,
                    "name": name,
                    "media_type": media_type,
                    "is_primary": is_primary,
                }
                for kind, name, media_type, is_primary in self.artifacts
            ],
            "diagnostics": [{"level": level, "code": code} for level, code in self.diagnostics],
            "error": ({"type": self.error[0], "diagnostic_code": self.error[1]} if self.error is not None else None),
        }


def canonical_manifest_bytes(document: OutputManifestDocument) -> bytes:
    """Serialize *document* deterministically as UTF-8 without a BOM."""
    payload = json.dumps(
        document.to_dict(),
        ensure_ascii=False,
        sort_keys=True,
        indent=2,
        allow_nan=False,
    )
    return f"{payload}\n".encode()


class OutputManifestWriter:
    """Publish optional manifest sidecars through the sole output finalizer."""

    def __init__(self, finalizer: OutputFinalizer) -> None:
        self._finalizer = finalizer

    def persist(self, request: ConversionRequest, result: Any) -> Any:
        """Return *result* with manifest artifacts added when policy enables them."""
        if isinstance(result, list):
            return [self._persist_list_item(request, item, index) for index, item in enumerate(result)]
        return self._persist_one(request, result, self._context_for_request(request))

    def _persist_list_item(self, request: ConversionRequest, result: Any, index: int) -> Any:
        context = self._context_for_request(request)
        child_context = context.for_input(index) if index < len(context.inputs) else context
        return self._persist_one(request, result, child_context)

    @staticmethod
    def _context_for_request(request: ConversionRequest) -> ConversionManifestContext:
        return request.manifest_context or ConversionManifestContext.from_request_inputs(
            request.input_refs,
            request.config_snapshot,
        )

    def _persist_one(
        self,
        request: ConversionRequest,
        result: Any,
        context: ConversionManifestContext,
    ) -> Any:
        if not isinstance(result, ConversionResult):
            return result
        if not context.policy.save_to_output or not request.output_policy.write_artifacts:
            return result
        if any(artifact.kind == ARTIFACT_KIND_MANIFEST for artifact in result.artifacts):
            return result
        if result.error is not None and result.error.error_type == "cancelled":
            return result
        if not context.inputs:
            return self._with_write_failure(result)

        document = self._build_document(request, result, context)
        suggested_name = self._suggested_name(result, context)
        digest = _request_digest(result.task_id)
        try:
            output_dir = self._finalizer.resolve_output_dir(
                request.output_policy,
                context.inputs[0].path,
            )
            with tempfile.TemporaryDirectory(prefix="docwen_manifest_", ignore_cleanup_errors=True) as staging_dir:
                staging_path = Path(staging_dir) / suggested_name
                staging_path.write_bytes(canonical_manifest_bytes(document))
                manifest_artifact = ArtifactManifest(
                    artifact_id=f"manifest-{digest}",
                    kind=ARTIFACT_KIND_MANIFEST,
                    staging_path=str(staging_path),
                    suggested_name=suggested_name,
                    media_type=_MANIFEST_MEDIA_TYPE,
                    is_primary=False,
                )
                published = self._finalizer.finalize(
                    result.task_id,
                    [manifest_artifact],
                    OutputPolicy(output_dir=output_dir, overwrite_mode="rename"),
                    input_path=context.inputs[0].path,
                )
        except Exception:
            return self._with_write_failure(result)

        if not published.success or len(published.artifacts) != 1:
            return self._with_write_failure(result)
        return replace(
            result,
            artifacts=[*result.artifacts, published.artifacts[0]],
            metrics=replace(
                result.metrics,
                output_bytes=result.metrics.output_bytes + published.metrics.output_bytes,
            ),
        )

    @staticmethod
    def _build_document(
        request: ConversionRequest,
        result: ConversionResult,
        context: ConversionManifestContext,
    ) -> OutputManifestDocument:
        mask = context.policy.mask_input_path
        inputs = tuple(
            (
                _masked_path(item.path) if mask else item.path,
                item.format,
                item.category,
            )
            for item in context.inputs
        )
        preconversion = tuple(
            (
                step.input_index,
                step.source_format,
                step.target_format,
                step.status,
                step.backend,
                step.diagnostic_code,
            )
            for step in context.preconversion_steps
        )
        artifacts = tuple(
            (
                artifact.kind,
                _portable_basename(artifact.staging_path) or _portable_basename(artifact.suggested_name),
                artifact.media_type,
                artifact.is_primary,
            )
            for artifact in result.artifacts
            if artifact.kind != ARTIFACT_KIND_MANIFEST
        )
        diagnostics = tuple((diagnostic.level, diagnostic.code) for diagnostic in result.diagnostics)
        error = (result.error.error_type, result.error.diagnostic_code) if result.error is not None else None
        return OutputManifestDocument(
            request_id_sha256=_request_digest(result.task_id, full=True),
            status="success" if result.success else "failed",
            target_format=request.target_format,
            action_name=request.action_name,
            inputs=inputs,
            preconversion=preconversion,
            artifacts=artifacts,
            diagnostics=diagnostics,
            error=error,
        )

    @staticmethod
    def _suggested_name(result: ConversionResult, context: ConversionManifestContext) -> str:
        digest = _request_digest(result.task_id)
        if not result.success:
            return f"manifest_failed_{digest}.json"
        if context.batch_child:
            return f"manifest_{digest}.json"
        return "manifest.json"

    @staticmethod
    def _with_write_failure(result: ConversionResult) -> ConversionResult:
        if any(diagnostic.code == _WRITE_FAILURE_CODE for diagnostic in result.diagnostics):
            return result
        return replace(
            result,
            diagnostics=[
                *result.diagnostics,
                ConversionDiagnostic(
                    level="warning",
                    code=_WRITE_FAILURE_CODE,
                    message="The configured output manifest could not be written.",
                ),
            ],
        )


def _request_digest(task_id: str, *, full: bool = False) -> str:
    digest = hashlib.sha256(str(task_id).encode("utf-8", errors="replace")).hexdigest()
    return digest if full else digest[:16]


def _portable_basename(path: str) -> str:
    normalized = str(path).replace("\\", "/").rstrip("/")
    return normalized.rsplit("/", 1)[-1] if normalized else ""


def _masked_path(path: str) -> str:
    return f"<redacted>/{_portable_basename(path)}"


__all__ = ["OutputManifestDocument", "OutputManifestWriter", "canonical_manifest_bytes"]
