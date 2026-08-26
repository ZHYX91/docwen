"""LayoutPlugin — entry point for docwen_plugin_layout."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docwen_plugin_layout.manifest import build_manifest
from docwen_plugin_layout.operations.converter import PdfMerger, PdfSplitter
from docwen_plugin_layout.preprocess import PreprocessResult, preprocess_layout_input
from docwen_plugin_layout.to_document.converter import LayoutToDocumentConverter
from docwen_plugin_layout.to_image.converter import LayoutToImageConverter
from docwen_plugin_layout.to_markdown.converter import LayoutToMarkdownConverter
from docwen_plugin_layout.to_pdf.converter import (
    LayoutToPdfConverter,
    OfdToPdfConverter,
    XpsToPdfConverter,
)

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import PluginExecutionContext


class LayoutPlugin:
    """Plugin for layout conversions, PDF operations, and Markdown conversion."""

    plugin_id: str
    _manifest: PluginManifest | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_layout"
        self._manifest = None

    @property
    def manifest(self) -> PluginManifest:
        if self._manifest is None:
            self._manifest = build_manifest()
        return self._manifest

    def can_handle(self, source_format: str, target_format: str, action_name: str = "") -> bool:
        for route in self.manifest.routes:
            if (
                route.source_format == source_format
                and route.target_format == target_format
                and route.action_name == action_name
            ):
                return True
        return False

    def convert(self, context: PluginExecutionContext) -> ConversionResult:
        """Dispatch to the appropriate converter based on source/target."""
        source = context.request.input_refs[0].format if context.request.input_refs else ""
        target = context.request.target_format

        # Named PDF operations must win over the generic pdf→pdf route.
        if context.request.action_name == "merge_pdfs":
            return PdfMerger().convert(context)
        if context.request.action_name == "split_pdf":
            return PdfSplitter().convert(context)

        # ── Layout → Markdown ─────────────────────────────────────────
        if source in ("pdf", "ofd", "xps") and target == "md":
            pre = preprocess_layout_input(
                context.workspace.input_path,
                context.workspace.staging_dir,
                source,
            )
            if pre.error_type:
                return self._preprocess_error(context, pre)
            return LayoutToMarkdownConverter().convert(
                context,
                input_path_override=pre.effective_input_path,
                source_format_override=pre.effective_source_format,
            )

        # ── Layout → Images (PNG / JPG / TIF) ───────────────────────
        if source in ("pdf", "ofd", "xps") and target in (
            "png",
            "jpg",
            "tif",
        ):
            pre = preprocess_layout_input(
                context.workspace.input_path,
                context.workspace.staging_dir,
                source,
            )
            if pre.error_type:
                return self._preprocess_error(context, pre)
            return LayoutToImageConverter(target).convert(
                context,
                input_path_override=pre.effective_input_path,
                source_format_override=pre.effective_source_format,
            )

        # ── Layout → Documents (DOCX / DOC / ODT / RTF) ─────────────
        if source in ("pdf", "ofd", "xps") and target in (
            "docx",
            "doc",
            "odt",
            "rtf",
        ):
            pre = preprocess_layout_input(
                context.workspace.input_path,
                context.workspace.staging_dir,
                source,
            )
            if pre.error_type:
                return self._preprocess_error(context, pre)
            return LayoutToDocumentConverter(target).convert(
                context,
                input_path_override=pre.effective_input_path,
                source_format_override=pre.effective_source_format,
            )

        # ── OFD → PDF ────────────────────────────────────────────────
        if source == "ofd" and target == "pdf":
            return OfdToPdfConverter().convert(context)

        # ── XPS → PDF ────────────────────────────────────────────────
        if source == "xps" and target == "pdf":
            return XpsToPdfConverter().convert(context)

        # ── Layout → PDF ─────────────────────────────────────────────
        if source == "pdf" and target == "pdf":
            return LayoutToPdfConverter().convert(context)

        # ── Fallback ─────────────────────────────────────────────────
        return self._unsupported_route(context, f"{source}→{target}")

    @staticmethod
    def _unsupported_route(context: PluginExecutionContext, route_label: str) -> ConversionResult:
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        msg = f"{route_label} is not an executable layout route."
        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="unsupported_route",
                message=msg,
                diagnostic_code="LAYOUT-UNSUPPORTED-ROUTE",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=msg,
                    code="LAYOUT-UNSUPPORTED-ROUTE",
                )
            ],
        )

    @staticmethod
    def _preprocess_error(context: PluginExecutionContext, pre: PreprocessResult) -> ConversionResult:
        """Convert a *PreprocessResult* error into a *ConversionResult*."""
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )

        return ConversionResult(
            task_id=context.request.request_id,
            success=False,
            error=ConversionErrorInfo(
                error_type=pre.error_type or "conversion_failed",
                message=pre.error_message or "Preprocessing failed",
                diagnostic_code=pre.diagnostic_code or "LAYOUT-PREPROCESS-ERROR",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=pre.error_message or "Preprocessing failed",
                    code=pre.diagnostic_code or "LAYOUT-PREPROCESS-ERROR",
                )
            ],
        )
