"""Shared PDF result builders — kept small and dependency-free."""

from __future__ import annotations

import uuid
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ConverterContext


# ── tiny internal utils (avoid a full _common module) ────────────────


def _new_artifact_id() -> str:
    return f"print-{uuid.uuid4().hex[:12]}"


def _file_size(path: str | Path) -> int:
    try:
        return Path(path).stat().st_size
    except OSError:
        return 0


# ── result builders ─────────────────────────────────────────────────


def _build_pdf_result(
    task_id: str,
    output_path: str,
    stem: str,
    source_format: str,
    context: ConverterContext,
) -> ConversionResult:
    """Build a successful ConversionResult for PDF output."""
    from docwen_core.models.artifact import (
        ARTIFACT_KIND_PRIMARY,
        ArtifactManifest,
    )
    from docwen_core.models.result import (
        ConversionDiagnostic,
        ConversionMetrics,
        ConversionResult,
    )

    out_bytes = _file_size(output_path)
    artifact = ArtifactManifest(
        artifact_id=_new_artifact_id(),
        kind=ARTIFACT_KIND_PRIMARY,
        staging_path=output_path,
        suggested_name=f"{stem}.pdf",
        media_type="application/pdf",
        metadata={"source_format": source_format},
        is_primary=True,
    )
    context.workspace.add_artifact(artifact)
    context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)

    return ConversionResult(
        task_id=task_id,
        success=True,
        artifacts=[artifact],
        diagnostics=[
            ConversionDiagnostic(
                level="info",
                message=f"Converted {stem} ({source_format}) to PDF ({out_bytes} bytes)",
                code="PDF-CONVERT-OK",
            )
        ],
        metrics=ConversionMetrics(
            input_bytes=_file_size(context.workspace.input_path),
            output_bytes=out_bytes,
            extra={"source_format": source_format},
        ),
    )


def _error_result(task_id: str, message: str, code: str, exc: Exception | None = None) -> ConversionResult:
    """Build an error ConversionResult."""
    from docwen_core.models.result import (
        ConversionDiagnostic,
        ConversionErrorInfo,
        ConversionResult,
    )

    return ConversionResult(
        task_id=task_id,
        success=False,
        error=ConversionErrorInfo(
            error_type="conversion_failed",
            message=message,
            diagnostic_code=code,
        ),
        diagnostics=[ConversionDiagnostic(level="error", message=message, code=code)],
    )
