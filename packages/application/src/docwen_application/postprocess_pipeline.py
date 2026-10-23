"""Compose rendering and proofreading through Application-owned Runtime stages."""

from __future__ import annotations

import tempfile
from collections.abc import Callable
from dataclasses import replace
from pathlib import Path
from typing import Any

from docwen_application.postprocessing import postprocess_proofread_options


def _pipeline_conversion_identity(request: Any, task_id: str) -> Any:
    """Freeze the original Markdown identity once for every pipeline stage."""

    from docwen_core.models import FILE_INSPECTION_METADATA_KEY
    from docwen_core.models.document_node import ConversionIdentity
    from docwen_core.models.file_ref import source_presentation_name

    source = next((ref for ref in request.input_refs if ref.input_role == "source"), request.input_refs[0])
    inspection = source.metadata.get(FILE_INSPECTION_METADATA_KEY)
    source_sha256 = str(inspection.get("content_sha256") or "") if isinstance(inspection, dict) else ""
    source_name = source_presentation_name(source)
    return ConversionIdentity.create(
        task_id=task_id,
        source_stem=Path(source_name).stem or "document",
        source_format=source.format,
        source_name=source_name,
        source_sha256=source_sha256,
    )


def execute_postprocessed_request(
    request: Any,
    task_id: str,
    execute_stage: Callable[[Any, str], Any],
) -> Any:
    """Execute a request, composing DOCX proofreading without plugin coupling."""

    proofread_options = postprocess_proofread_options(request)
    if proofread_options is None:
        return execute_stage(request, task_id)
    from docwen_core.models.request import POSTPROCESS_PROOFREAD_OPTION
    from docwen_core.models.result import ConversionDiagnostic, ConversionResult
    from docwen_core.options.proofread import resolve_proofread_switches

    switches = resolve_proofread_switches(request.config_snapshot, proofread_options)
    proofread_options = {**proofread_options, **switches}
    if not any(switches.values()):
        options = dict(request.options)
        options.pop(POSTPROCESS_PROOFREAD_OPTION)
        result = execute_stage(replace(request, options=options), task_id)
        if isinstance(result, ConversionResult) and result.success:
            result = replace(
                result,
                diagnostics=[
                    *result.diagnostics,
                    ConversionDiagnostic(
                        level="info",
                        code="POSTPROCESS_PROOFREAD_SKIPPED",
                        message="All proofreading checks are disabled; the generated document was published without proofreading.",
                    ),
                ],
            )
        return result
    return _execute_docx_proofread_pipeline(
        request,
        task_id,
        execute_stage,
        proofread_options=proofread_options,
    )


def _execute_docx_proofread_pipeline(
    request: Any,
    task_id: str,
    execute_stage: Callable[[Any, str], Any],
    *,
    proofread_options: dict[str, bool],
) -> Any:
    """Render Markdown privately, then run the existing DOCX validator.

    The Markdown and proofread plugins remain unaware of each other.  The
    first stage publishes only inside an Application-owned temporary
    directory; only the validated second stage reaches the caller's output
    policy.
    """

    from docwen_core.formats import get_category, get_media_type
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import (
        POSTPROCESS_PROOFREAD_OPTION,
        ConversionRequest,
        OutputPolicy,
    )
    from docwen_core.models.result import (
        ConversionDiagnostic,
        ConversionErrorInfo,
        ConversionMetrics,
        ConversionResult,
    )

    source = next((ref for ref in request.input_refs if ref.input_role == "source"), None)
    if source is None or source.format not in {"md", "markdown", "txt"}:
        raise ValueError("Post-conversion proofreading requires an admitted Markdown source")

    identity = request.conversion_identity or _pipeline_conversion_identity(request, task_id)
    original_target = str(request.target_format).lower()
    final_is_docx = original_target == "docx"
    render_options = dict(request.options)
    render_options.pop(POSTPROCESS_PROOFREAD_OPTION, None)

    with tempfile.TemporaryDirectory(prefix="docwen_postprocess_", ignore_cleanup_errors=True) as private_output:
        # tempfile may use macOS's /var alias. This directory is ours; pass its
        # physical path to the next admitted stage without relaxing input checks.
        private_output = str(Path(private_output).resolve(strict=True))
        render_task_id = f"{task_id}-render"
        render_request = replace(
            request,
            request_id=render_task_id,
            target_format="docx",
            options=render_options,
            output_policy=OutputPolicy(
                output_dir=private_output,
                overwrite_mode="error",
                write_artifacts=True,
                group_outputs=True,
                open_after_done=False,
            ),
            conversion_identity=replace(identity, task_id=render_task_id),
        )
        render_result = execute_stage(render_request, render_task_id)
        if not isinstance(render_result, ConversionResult):
            return render_result

        private_codes = {"FINALIZER_DONE", "DOCUMENT_NODE_REUSED", "OUTPUT_MANIFEST_WRITE_FAILED"}
        render_diagnostics = [item for item in render_result.diagnostics if item.code not in private_codes]
        if not render_result.success or render_result.error is not None:
            return replace(
                render_result,
                task_id=task_id,
                artifacts=[],
                diagnostics=render_diagnostics,
                metrics=ConversionMetrics(
                    duration_ms=render_result.metrics.duration_ms,
                    input_bytes=source.size_bytes,
                    output_bytes=0,
                    extra={"proofread_postprocess": True, "stage": "render"},
                ),
            )

        primary_docx = [
            artifact
            for artifact in render_result.artifacts
            if artifact.is_primary
            and artifact.media_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ]
        if len(primary_docx) != 1:
            message = "Markdown rendering did not produce exactly one primary DOCX for proofreading."
            return ConversionResult(
                task_id=task_id,
                success=False,
                diagnostics=[
                    *render_diagnostics,
                    ConversionDiagnostic(
                        level="error",
                        message=message,
                        code="POSTPROCESS_PROOFREAD_INPUT_INVALID",
                    ),
                ],
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=message,
                    diagnostic_code="POSTPROCESS_PROOFREAD_INPUT_INVALID",
                ),
                metrics=ConversionMetrics(
                    duration_ms=render_result.metrics.duration_ms,
                    input_bytes=source.size_bytes,
                    output_bytes=0,
                    extra={"proofread_postprocess": True, "stage": "render"},
                ),
            )

        artifact = primary_docx[0]
        proofread_ref = FileRef(
            path=artifact.staging_path,
            format="docx",
            category=get_category("docx"),
            input_kind="document",
            input_role="source",
            logical_path=artifact.suggested_name or f"{request.source_stem}.docx",
            media_type=get_media_type("docx"),
            size_bytes=artifact.size_bytes or 0,
        )
        # The original DOCX workflow publishes the validated document directly.
        # A different target needs a private validated DOCX until conversion succeeds.
        proofread_output = Path(private_output) / "validated" if not final_is_docx else None
        if proofread_output is not None:
            proofread_output.mkdir(parents=True, exist_ok=True)
        proofread_output_dir = (
            str(proofread_output)
            if proofread_output is not None
            else request.output_policy.output_dir
            if request.output_policy.output_dir or request.output_policy.output_path
            else str(Path(source.path).parent)
        )
        proofread_request = ConversionRequest(
            request_id=task_id,
            input_refs=[proofread_ref],
            target_format="docx",
            action_name="validate",
            options={
                **proofread_options,
                **({"locale": request.options["locale"]} if "locale" in request.options else {}),
            },
            output_policy=replace(
                request.output_policy,
                group_outputs=True,
                output_dir=proofread_output_dir,
            ),
            config_snapshot=dict(request.config_snapshot),
            manifest_context=request.manifest_context,
            conversion_identity=identity,
        )
        proofread_result = execute_stage(proofread_request, task_id)
        if not isinstance(proofread_result, ConversionResult):
            return proofread_result

        diagnostics = [
            *render_diagnostics,
            *proofread_result.diagnostics,
        ]
        if proofread_result.success:
            diagnostics.append(
                ConversionDiagnostic(
                    level="info",
                    message=(
                        "Generated DOCX was proofread before publication."
                        if final_is_docx
                        else "Generated DOCX was proofread before final document conversion."
                    ),
                    code="POSTPROCESS_PROOFREAD_APPLIED",
                )
            )
        if not final_is_docx:
            if not proofread_result.success:
                return replace(
                    proofread_result,
                    task_id=task_id,
                    artifacts=[],
                    diagnostics=diagnostics,
                    metrics=ConversionMetrics(
                        duration_ms=render_result.metrics.duration_ms + proofread_result.metrics.duration_ms,
                        input_bytes=source.size_bytes,
                        output_bytes=0,
                        extra={"proofread_postprocess": True, "stage": "proofread"},
                    ),
                )
            from docwen_application.preconversion.intermediate_policy import (
                build_intermediate_record_if_enabled,
            )
            from docwen_core.models.request import PRECONVERSION_INTERMEDIATES_OPTION

            checked = [
                item
                for item in proofread_result.artifacts
                if item.is_primary
                and item.media_type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ]
            if len(checked) != 1:
                message = "Proofreading did not produce exactly one validated DOCX."
                return ConversionResult(
                    task_id=task_id,
                    success=False,
                    artifacts=[],
                    diagnostics=[
                        *diagnostics,
                        ConversionDiagnostic(
                            level="error",
                            message=message,
                            code="POSTPROCESS_PROOFREAD_INPUT_INVALID",
                        ),
                    ],
                    error=ConversionErrorInfo(
                        error_type="conversion_failed",
                        message=message,
                        diagnostic_code="POSTPROCESS_PROOFREAD_INPUT_INVALID",
                    ),
                )
            checked_docx = checked[0]
            checked_ref = FileRef(
                path=checked_docx.staging_path,
                format="docx",
                category=get_category("docx"),
                input_kind="document",
                input_role="source",
                logical_path=checked_docx.suggested_name or f"{request.source_stem}.docx",
                media_type=get_media_type("docx"),
                size_bytes=checked_docx.size_bytes or Path(checked_docx.staging_path).stat().st_size,
            )
            retained = build_intermediate_record_if_enabled(
                checked_ref.path,
                identity.source_stem,
                source.format,
                target_format="docx",
                backend="proofread",
                config_snapshot=request.config_snapshot,
            )
            final_options: dict[str, Any] = {}
            if retained is not None:
                retained["applies_to_input_path"] = checked_ref.path
                final_options[PRECONVERSION_INTERMEDIATES_OPTION] = [retained]
            final_request = ConversionRequest(
                request_id=task_id,
                input_refs=[checked_ref],
                target_format=original_target,
                action_name="",
                options=final_options,
                output_policy=replace(
                    request.output_policy,
                    group_outputs=True,
                    output_dir=request.output_policy.output_dir or str(Path(source.path).parent),
                ),
                config_snapshot=dict(request.config_snapshot),
                manifest_context=request.manifest_context,
                conversion_identity=identity,
            )
            final_result = execute_stage(final_request, task_id)
            if not isinstance(final_result, ConversionResult):
                return final_result
            final_diagnostics = [
                *render_diagnostics,
                *[item for item in proofread_result.diagnostics if item.code not in private_codes],
                *final_result.diagnostics,
            ]
            if final_result.success:
                final_diagnostics.append(
                    ConversionDiagnostic(
                        level="info",
                        message="Validated DOCX was converted to the selected document format.",
                        code="POSTPROCESS_PROOFREAD_APPLIED",
                    )
                )
            return replace(
                final_result,
                task_id=task_id,
                diagnostics=final_diagnostics,
                metrics=ConversionMetrics(
                    duration_ms=(
                        render_result.metrics.duration_ms
                        + proofread_result.metrics.duration_ms
                        + final_result.metrics.duration_ms
                    ),
                    input_bytes=source.size_bytes,
                    output_bytes=final_result.metrics.output_bytes,
                    extra={**final_result.metrics.extra, "proofread_postprocess": True},
                ),
            )
        return replace(
            proofread_result,
            task_id=task_id,
            diagnostics=diagnostics,
            metrics=ConversionMetrics(
                duration_ms=render_result.metrics.duration_ms + proofread_result.metrics.duration_ms,
                input_bytes=source.size_bytes,
                output_bytes=proofread_result.metrics.output_bytes,
                extra={
                    **proofread_result.metrics.extra,
                    "proofread_postprocess": True,
                },
            ),
        )
