"""SmartConverter: document-format interconversion hub."""

from __future__ import annotations

import os
import shutil
import time
import uuid
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ConverterContext

from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.office_bridge import BridgeCandidate, convert_with_backend_priority

# Source/target formats covered by SmartConverter.
_SMARTCONV_FORMATS: frozenset[str] = frozenset({"docx", "doc", "odt", "rtf", "wps"})
_DOCX_SAVE_FORMAT = 12
_DOC_SAVE_FORMAT = 0
_ODT_SAVE_FORMAT = 23
_RTF_SAVE_FORMAT = 6
_DEFAULT_WORD_PRIORITY = ("wps_writer", "msoffice_word", "libreoffice")
_DEFAULT_ODT_PRIORITY = ("msoffice_word", "libreoffice")

_BEST_EFFORT_LOSS_CLASSES: dict[str, str] = {
    "doc": "fields, revisions/comments, inline-object identities, and layout",
    "rtf": "fields, revisions/comments, inline-object identities, and layout",
    "odt": "paragraphs, tables, fields, revisions, shapes, sections, and pagination",
}

_MEDIA_TYPES: dict[str, str] = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "doc": "application/msword",
    "odt": "application/vnd.oasis.opendocument.text",
    "rtf": "application/rtf",
    "wps": "application/msword",
}


def _word_candidates(save_format: int, *, source_format: str, target_format: str) -> dict[str, BridgeCandidate]:
    if source_format == "odt" or target_format == "odt":
        return {
            "msoffice_word": BridgeCandidate(
                "Microsoft Word",
                "Word.Application",
                save_format,
                "word",
                suppress_new_revisions=True,
            )
        }
    return {
        "wps_writer": BridgeCandidate(
            "WPS Writer",
            "Kwps.Application",
            save_format,
            "word",
            suppress_new_revisions=True,
        ),
        "msoffice_word": BridgeCandidate(
            "Microsoft Word",
            "Word.Application",
            save_format,
            "word",
            suppress_new_revisions=True,
        ),
    }


def _configured_priority(context: ConverterContext, *, source_format: str, target_format: str) -> list[str]:
    uses_odt = source_format == "odt" or target_format == "odt"
    if uses_odt:
        key = "software.special_conversions.odt"
        default = _DEFAULT_ODT_PRIORITY
    else:
        key = "software.default_priority.word_processors"
        default = _DEFAULT_WORD_PRIORITY
    configured = context.config.get(key, list(default))
    priority: list[str]
    if isinstance(configured, (list, tuple)):
        priority = [str(item) for item in configured if isinstance(item, str)]
    else:
        priority = list(default)

    request_source_format = context.request.input_refs[0].format if context.request.input_refs else ""
    if (
        request_source_format == "docx"
        and source_format == "docx"
        and target_format in _BEST_EFFORT_LOSS_CLASSES
        and "msoffice_word" in priority
    ):
        return ["msoffice_word", *(backend for backend in priority if backend != "msoffice_word")]
    return priority


def _target_spec(target_format: str) -> tuple[str | None, int | None]:
    if target_format == "docx":
        return "docx", _DOCX_SAVE_FORMAT
    if target_format == "doc":
        return "doc", _DOC_SAVE_FORMAT
    if target_format == "odt":
        return "odt", _ODT_SAVE_FORMAT
    if target_format == "rtf":
        return "rtf", _RTF_SAVE_FORMAT
    if target_format == "wps":
        return "doc", _DOC_SAVE_FORMAT
    return None, None


def _build_error(
    *,
    task_id: str,
    code: str,
    error_type: str,
    message: str,
) -> ConversionResult:
    return ConversionResult(
        task_id=task_id,
        success=False,
        error=ConversionErrorInfo(
            error_type=error_type,
            message=message,
            diagnostic_code=code,
            recoverable=True,
        ),
        diagnostics=[ConversionDiagnostic(level="error", message=message, code=code)],
    )


class SmartDocConverter:
    """Document-format interconversion hub backed by COM / LibreOffice."""

    # Convenience set for plugin.py dispatch guards.
    HANDLED_FORMATS: frozenset[str] = _SMARTCONV_FORMATS

    def convert(self, context: ConverterContext) -> ConversionResult:
        t_start = time.monotonic()
        task_id = context.request.request_id
        if not context.request.input_refs:
            return _build_error(
                task_id=task_id,
                code="DOCX-SMARTDOC-NO-INPUT",
                error_type="invalid_input",
                message="SmartDocConverter requires one input file.",
            )

        source = context.request.input_refs[0].format
        target = context.request.target_format
        input_path = context.workspace.input_path
        if not os.path.exists(input_path) or os.path.getsize(input_path) == 0:
            return _build_error(
                task_id=task_id,
                code="DOCX-SMARTDOC-INVALID-INPUT",
                error_type="invalid_input",
                message=f"SmartDoc input for {source}->{target} is empty or missing.",
            )
        context.progress.report_progress(0.0, f"Starting {source.upper()}→{target.upper()} conversion")
        context.cancellation.check()

        if source == target:
            output_path = context.workspace.create_artifact_path("primary", f".{target}")
            shutil.copy2(input_path, output_path)
            return self._success_result(
                context=context,
                source=source,
                target=target,
                output_path=output_path,
                backend="copy",
                started_at=t_start,
            )

        if source == "docx" or target == "docx":
            result = self._convert_one_hop(
                context=context,
                input_path=input_path,
                source=source,
                target=target,
                output_path=context.workspace.create_artifact_path("primary", f".{target}"),
                cancel=context.cancellation,
            )
            if not result.success or not result.output_path:
                return _build_error(
                    task_id=task_id,
                    code="DOCX-SMARTDOC-BACKEND",
                    error_type="dependency_missing",
                    message=result.message,
                )
            return self._success_result(
                context=context,
                source=source,
                target=target,
                output_path=result.output_path,
                backend=result.backend,
                started_at=t_start,
            )

        hub_docx = context.workspace.create_artifact_path("auxiliary", ".docx")
        first_leg = self._convert_one_hop(
            context=context,
            input_path=input_path,
            source=source,
            target="docx",
            output_path=hub_docx,
            cancel=context.cancellation,
        )
        if not first_leg.success or not first_leg.output_path:
            return _build_error(
                task_id=task_id,
                code="DOCX-SMARTDOC-BACKEND",
                error_type="dependency_missing",
                message=first_leg.message,
            )

        context.progress.report_progress(55.0, "Intermediate DOCX generated")
        final_output = context.workspace.create_artifact_path("primary", f".{target}")
        second_leg = self._convert_one_hop(
            context=context,
            input_path=first_leg.output_path,
            source="docx",
            target=target,
            output_path=final_output,
            cancel=context.cancellation,
        )
        if not second_leg.success or not second_leg.output_path:
            return _build_error(
                task_id=task_id,
                code="DOCX-SMARTDOC-BACKEND",
                error_type="dependency_missing",
                message=second_leg.message,
            )

        return self._success_result(
            context=context,
            source=source,
            target=target,
            output_path=second_leg.output_path,
            backend=f"{first_leg.backend} -> {second_leg.backend}",
            started_at=t_start,
        )

    def _convert_one_hop(
        self,
        *,
        context: ConverterContext,
        input_path: str,
        source: str,
        target: str,
        output_path: str,
        cancel: object | None = None,
    ):
        libreoffice_format, save_format = _target_spec(target)
        if save_format is None:
            return convert_with_backend_priority(
                input_path,
                output_path,
                source_format=source,
                backend_priority=[],
                com_candidates={},
                libreoffice_format=None,
                cancel=cancel,
                failure_subject=f"Configured SmartDoc {source.upper()}→{target.upper()} backends",
            )

        bridge_output = Path(output_path)
        actual_output = output_path
        if target == "wps":
            actual_output = str(bridge_output.with_suffix(".doc"))

        result = convert_with_backend_priority(
            input_path,
            actual_output,
            source_format=source,
            backend_priority=_configured_priority(
                context,
                source_format=source,
                target_format=target,
            ),
            com_candidates=_word_candidates(save_format, source_format=source, target_format=target),
            libreoffice_format=libreoffice_format,
            cancel=cancel,
            failure_subject=f"Configured SmartDoc {source.upper()}→{target.upper()} backends",
        )
        if not result.success or not result.output_path:
            return result

        produced = Path(result.output_path)
        if target == "wps":
            produced.replace(bridge_output)
            result.output_path = str(bridge_output)
        return result

    def _success_result(
        self,
        *,
        context: ConverterContext,
        source: str,
        target: str,
        output_path: str,
        backend: str,
        started_at: float,
    ) -> ConversionResult:
        artifact = ArtifactManifest(
            artifact_id=str(uuid.uuid4()),
            kind="primary",
            staging_path=output_path,
            suggested_name=f"{Path(context.workspace.input_path).stem}.{target}",
            media_type=_MEDIA_TYPES.get(target, "application/octet-stream"),
            metadata={
                "source_format": source,
                "target_format": target,
                "backend": backend,
            },
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)
        context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)
        context.progress.report_progress(100.0, "Document conversion complete")
        context.logger.info(f"SmartDoc {source}->{target} completed via {backend}")
        diagnostics = [
            ConversionDiagnostic(
                level="info",
                message=f"Converted {source.upper()} to {target.upper()} via {backend}.",
                code="DOCX-SMARTDOC-OK",
            )
        ]
        loss_classes = _BEST_EFFORT_LOSS_CLASSES.get(target) if source == "docx" else None
        if loss_classes is not None:
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        f"Best-effort DOCX to {target.upper()} conversion via {backend} may change or lose "
                        f"{loss_classes}; the source file was not modified. Review the output against the source."
                    ),
                    code="DOCX-SMARTDOC-BEST-EFFORT-LOSS",
                )
            )

        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[artifact],
            diagnostics=diagnostics,
            metrics=ConversionMetrics(
                duration_ms=(time.monotonic() - started_at) * 1000.0,
                input_bytes=os.path.getsize(context.workspace.input_path)
                if os.path.exists(context.workspace.input_path)
                else 0,
                output_bytes=os.path.getsize(output_path) if os.path.exists(output_path) else 0,
                extra={"backend": backend},
            ),
        )
