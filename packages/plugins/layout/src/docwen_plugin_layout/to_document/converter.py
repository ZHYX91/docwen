"""Layout → DOCX / DOC / ODT / RTF converters backed by Office/pdf2docx."""

from __future__ import annotations

from contextlib import suppress
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models.result import (
    ArtifactManifest,
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.office_bridge import BridgeCandidate, BridgeResult, convert_with_backend_priority

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import CancellationTokenView, ConverterContext


_MEDIA_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "doc": "application/msword",
    "odt": "application/vnd.oasis.opendocument.text",
    "rtf": "application/rtf",
}

_WORD_SAVE_FORMATS = {
    "docx": 16,
    "doc": 0,
    "odt": 23,
    "rtf": 6,
}

_PDF_TO_OFFICE_DEFAULT_PRIORITY = ("msoffice_word", "libreoffice")
_WORD_PROCESSOR_DEFAULT_PRIORITY = ("wps_writer", "msoffice_word", "libreoffice")
_ODT_DEFAULT_PRIORITY = ("msoffice_word", "libreoffice")

_PDF_TO_DOCX_COM_CANDIDATES = {
    "msoffice_word": BridgeCandidate(
        name="Microsoft Word",
        prog_id="Word.Application",
        save_format=_WORD_SAVE_FORMATS["docx"],
        app_type="word",
    ),
}


class LayoutToDocumentConverter:
    """Converter for layout→document routes.

    DOC/DOCX/ODT/RTF generally use the shared Office bridge.  For PDF→DOCX,
    preserve the old priority semantics: try configured external Office
    backends first, then fall back to the route-local pure-Python pdf2docx
    backend when none of them succeeds.
    """

    def __init__(self, target_format: str) -> None:
        if target_format not in _MEDIA_TYPES:
            raise ValueError(f"Unsupported target format: {target_format}")
        self._target_format = target_format

    def convert(
        self,
        context: ConverterContext,
        *,
        input_path_override: str | None = None,
        source_format_override: str | None = None,
    ) -> ConversionResult:
        context.cancellation.check()
        input_ref = context.request.input_refs[0]
        input_path = Path(input_path_override or input_ref.path)
        source_stem = Path(input_ref.path).stem
        source_format = str(source_format_override or input_ref.format or "unknown").strip().lower()
        output_path = context.workspace.create_artifact_path(source_stem, f".{self._target_format}")

        if self._target_format == "docx" and source_format == "pdf":
            office_result = _convert_pdf_with_configured_office_priority(context, input_path, Path(output_path))
            if office_result.success and office_result.output_path is not None:
                return self._success_result(context, source_stem, Path(office_result.output_path), office_result)

            pdf2docx_result = _convert_pdf_with_pdf2docx(
                input_path,
                Path(output_path),
                cancellation=context.cancellation,
            )
            if pdf2docx_result.success and pdf2docx_result.output_path is not None:
                return self._success_result(context, source_stem, Path(pdf2docx_result.output_path), pdf2docx_result)

            message = _combine_pdf_to_docx_failure_messages(office_result, pdf2docx_result)
            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=message,
                    diagnostic_code="LAYOUT-PDF2DOCX-FAILED",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=message,
                        code="LAYOUT-PDF2DOCX-FAILED",
                    )
                ],
            )

        bridge_result = convert_with_backend_priority(
            str(input_path),
            str(output_path),
            source_format=source_format,
            backend_priority=_document_priority(context, self._target_format),
            com_candidates=_document_candidates(self._target_format),
            libreoffice_format=self._target_format,
            cancel=context.cancellation,
            failure_subject=f"Configured Layout→{self._target_format.upper()} backends",
        )

        if not bridge_result.success or bridge_result.output_path is None:
            message = bridge_result.message or (
                f"Layout→{self._target_format.upper()} conversion failed via Office bridge."
            )
            return ConversionResult(
                task_id=context.request.request_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=message,
                    diagnostic_code="LAYOUT-OFFICE-BRIDGE-FAILED",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error",
                        message=message,
                        code="LAYOUT-OFFICE-BRIDGE-FAILED",
                    )
                ],
            )

        return self._success_result(context, source_stem, Path(bridge_result.output_path), bridge_result)

    def _success_result(
        self,
        context: ConverterContext,
        source_stem: str,
        output_path: Path,
        bridge_result: BridgeResult,
    ) -> ConversionResult:
        return ConversionResult(
            task_id=context.request.request_id,
            success=True,
            artifacts=[
                ArtifactManifest(
                    artifact_id=f"layout-{self._target_format}",
                    kind="primary",
                    staging_path=str(output_path),
                    suggested_name=f"{source_stem}.{self._target_format}",
                    media_type=_MEDIA_TYPES[self._target_format],
                    is_primary=True,
                )
            ],
            metrics=ConversionMetrics(
                extra={
                    "engine": "pdf2docx" if bridge_result.backend == "pdf2docx" else "office_bridge",
                    "backend": bridge_result.backend,
                }
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="info",
                    message=(
                        f"Converted layout to {self._target_format.upper()} "
                        f"via {bridge_result.backend or 'office bridge'}."
                    ),
                    code="LAYOUT-OFFICE-BRIDGE-OK",
                )
            ],
        )


def _convert_pdf_with_pdf2docx(
    input_path: Path,
    output_path: Path,
    *,
    cancellation: CancellationTokenView | None = None,
) -> BridgeResult:
    """Convert PDF bytes to DOCX with the pure-Python pdf2docx fallback.

    ``pdf2docx.Converter`` otherwise reopens a filename and lets PyMuPDF infer
    its parser from the suffix.  Its stream API is the content-first boundary.
    The one in-memory copy is unavoidable with the supported pdf2docx API;
    cancellation is checked on both sides of that potentially large read.
    """
    try:
        from pdf2docx import Converter
    except ImportError:
        return BridgeResult(
            False,
            message=(
                "PDF→DOCX reached the pdf2docx fallback, but the pdf2docx Python "
                "package is not installed. Reinstall DocWen with the layout plugin "
                "dependencies."
            ),
        )

    if cancellation is not None:
        cancellation.check()
    try:
        pdf_bytes = input_path.read_bytes()
    except Exception as exc:
        return BridgeResult(False, message=f"PDF→DOCX could not read the admitted PDF bytes: {exc}")
    if cancellation is not None:
        cancellation.check()

    converter = None
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        converter = Converter(stream=pdf_bytes)
        converter.convert(str(output_path))
    except Exception as exc:
        return BridgeResult(False, message=f"PDF→DOCX conversion failed via pdf2docx: {exc}")
    finally:
        if converter is not None:
            with suppress(Exception):
                converter.close()

    if cancellation is not None:
        cancellation.check()

    if output_path.exists():
        return BridgeResult(True, output_path=str(output_path), backend="pdf2docx")

    return BridgeResult(False, message="pdf2docx completed but did not create the DOCX output file.")


def _convert_pdf_with_configured_office_priority(
    context: ConverterContext,
    input_path: Path,
    output_path: Path,
) -> BridgeResult:
    """Try configured external PDF→Office backends before pdf2docx fallback."""
    return convert_with_backend_priority(
        str(input_path),
        str(output_path),
        source_format="pdf",
        backend_priority=_pdf_to_office_priority(context),
        com_candidates=_PDF_TO_DOCX_COM_CANDIDATES,
        libreoffice_format="docx",
        cancel=context.cancellation,
        failure_subject="Configured Office/LibreOffice PDF→DOCX backends",
    )


def _document_candidates(target_format: str) -> dict[str, BridgeCandidate]:
    candidates = {
        "wps_writer": BridgeCandidate(
            name="WPS Writer",
            prog_id="KWPS.Application",
            save_format=_WORD_SAVE_FORMATS[target_format],
            app_type="word",
        ),
        "msoffice_word": BridgeCandidate(
            name="Microsoft Word",
            prog_id="Word.Application",
            save_format=_WORD_SAVE_FORMATS[target_format],
            app_type="word",
        ),
    }
    if target_format == "odt":
        return {"msoffice_word": candidates["msoffice_word"]}
    return candidates


def _document_priority(context: ConverterContext, target_format: str) -> list[str]:
    if target_format == "odt":
        key = "software.special_conversions.odt"
        default = _ODT_DEFAULT_PRIORITY
    else:
        key = "software.default_priority.word_processors"
        default = _WORD_PROCESSOR_DEFAULT_PRIORITY
    configured = context.config.get(key, list(default))
    if isinstance(configured, (list, tuple)):
        return [str(item) for item in configured if isinstance(item, str)]
    return list(default)


def _pdf_to_office_priority(context: ConverterContext) -> list[str]:
    configured = context.config.get(
        "software.special_conversions.pdf_to_office",
        list(_PDF_TO_OFFICE_DEFAULT_PRIORITY),
    )
    if isinstance(configured, (list, tuple)):
        return [str(item) for item in configured if isinstance(item, str)]
    return list(_PDF_TO_OFFICE_DEFAULT_PRIORITY)


def _combine_pdf_to_docx_failure_messages(office_result: BridgeResult, pdf2docx_result: BridgeResult) -> str:
    messages = [
        message
        for message in (
            office_result.message,
            pdf2docx_result.message,
        )
        if message
    ]
    if messages:
        return "PDF→DOCX conversion failed. " + " ".join(messages)
    return "PDF→DOCX conversion failed via Office/LibreOffice and pdf2docx fallback."
