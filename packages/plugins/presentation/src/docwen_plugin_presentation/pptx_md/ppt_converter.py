"""Legacy PPT -> PPTX -> Markdown bridge."""

from __future__ import annotations

import os
from typing import TYPE_CHECKING

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest
from docwen_core.models.result import ConversionDiagnostic, ConversionErrorInfo, ConversionResult
from docwen_core.office_bridge import BridgeCandidate, convert_with_backend_priority
from docwen_core.protocols.hub_context import HubConversionContext, HubWorkspaceHandle

from .converter import PptxToMarkdownConverter
from .request_policy import build_presentation_markdown_request_policy

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


_DEFAULT_PRESENTATION_PRIORITY = ("wps_presentation", "msoffice_powerpoint", "libreoffice")
_PRESENTATION_CANDIDATES = {
    "wps_presentation": BridgeCandidate("WPS Presentation", "Kwpp.Application", 24, "powerpoint"),
    "msoffice_powerpoint": BridgeCandidate(
        "Microsoft PowerPoint",
        "PowerPoint.Application",
        24,
        "powerpoint",
    ),
}


def _configured_priority(context: ConverterContext) -> list[str]:
    configured = context.config.get(
        "software.default_priority.presentation_processors",
        list(_DEFAULT_PRESENTATION_PRIORITY),
    )
    if isinstance(configured, (list, tuple)):
        return [str(item) for item in configured if isinstance(item, str)]
    return list(_DEFAULT_PRESENTATION_PRIORITY)


class PptToMarkdownConverter:
    def __init__(self) -> None:
        self._delegate = PptxToMarkdownConverter()

    def convert(self, context: ConverterContext) -> ConversionResult:
        if not context.request.input_refs:
            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="invalid_input",
                    message="PPT→MD requires one input file.",
                    diagnostic_code="PPT2MD-NO-INPUT",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message="PPT→MD requires one input file.",
                        code="PPT2MD-NO-INPUT",
                    )
                ],
            )

        input_ref = context.request.input_refs[0]
        if not os.path.exists(input_ref.path) or os.path.getsize(input_ref.path) == 0:
            message = "PPT input is empty or missing."
            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="invalid_input",
                    message=message,
                    diagnostic_code="PPT2MD-INVALID-INPUT",
                ),
                diagnostics=[ConversionDiagnostic(level="error", message=message, code="PPT2MD-INVALID-INPUT")],
            )

        request_policy = build_presentation_markdown_request_policy(
            context,
            context.request.options or {},
        )
        pptx_path = context.workspace.create_artifact_path("auxiliary", ".pptx")
        result = convert_with_backend_priority(
            input_ref.path,
            pptx_path,
            source_format=str(input_ref.format or "").strip().lower(),
            backend_priority=_configured_priority(context),
            com_candidates=_PRESENTATION_CANDIDATES,
            libreoffice_format="pptx",
            cancel=context.cancellation,
            failure_subject="Configured presentation bridge backends",
        )
        if not result.success or not result.output_path:
            message = (
                f"PPT→PPTX preprocessing failed. {result.message or 'Install Microsoft Office/WPS or LibreOffice.'}"
            )
            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="dependency_missing",
                    message=message,
                    diagnostic_code="PPT2MD-BACKEND",
                    recoverable=True,
                ),
                diagnostics=[ConversionDiagnostic(level="error", message=message, code="PPT2MD-BACKEND")],
            )

        proxy_request = ConversionRequest(
            request_id=context.request.request_id,
            input_refs=[
                FileRef(
                    path=result.output_path,
                    format="pptx",
                    category=input_ref.category,
                    encoding=input_ref.encoding,
                    size_bytes=input_ref.size_bytes,
                    metadata=dict(input_ref.metadata),
                )
            ],
            target_format="md",
            action_name=context.request.action_name,
            options=dict(context.request.options),
            output_policy=context.request.output_policy,
            config_snapshot=dict(context.request.config_snapshot),
        )
        proxy_context = HubConversionContext(
            base=context,
            request=proxy_request,
            workspace=HubWorkspaceHandle(context.workspace, result.output_path),
        )
        proxy_context.logger.info(f"PPT→MD preprocessing completed via {result.backend}")
        return self._delegate.convert(
            proxy_context,
            source_path_for_naming=input_ref.path,
            request_policy=request_policy,
        )
