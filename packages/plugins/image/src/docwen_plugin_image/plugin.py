"""ImagePlugin — entry point for docwen_plugin_image."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from docwen_plugin_image.format_conversion.converter import ImageFormatConverter
from docwen_plugin_image.manifest import build_manifest
from docwen_plugin_image.merge.converter import ImageToTiffMerger
from docwen_plugin_image.to_markdown.converter import ImageToMarkdownConverter
from docwen_plugin_image.to_pdf.converter import ImageToPdfConverter

if TYPE_CHECKING:
    from docwen_core.models.manifest import PluginManifest
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ConverterContext, PluginExecutionContext


class ImagePlugin:
    """Plugin for image conversions and image merge operations."""

    plugin_id: str
    _manifest: PluginManifest | None

    def __init__(self) -> None:
        self.plugin_id = "docwen_plugin_image"
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
        from docwen_core.models.file_ref import FileRef
        from docwen_core.models.request import ConversionRequest
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionResult,
        )
        from docwen_core.protocols.hub_context import HubConversionContext, HubWorkspaceHandle
        from docwen_plugin_image._common import is_heic_format, preconvert_heic_to_png, source_format_from_context

        input_path = context.workspace.input_path
        converter_context: ConverterContext = context
        action = context.request.action_name

        if action == "merge_images_to_tiff":
            return ImageToTiffMerger().convert(context)

        source_format = source_format_from_context(context)
        if is_heic_format(source_format):
            try:
                png_path = preconvert_heic_to_png(input_path, context.workspace.staging_dir)
            except RuntimeError as exc:
                if exc.__cause__ is not None and isinstance(exc.__cause__, ImportError):
                    error_type = "dependency_missing"
                    code = "IMG-HEIC-DEPENDENCY-MISSING"
                else:
                    error_type = "conversion_failed"
                    code = "IMG-HEIC-PREPROCESS-ERROR"
                msg = str(exc)
                return ConversionResult(
                    task_id=context.request.request_id,
                    success=False,
                    error=ConversionErrorInfo(error_type=error_type, message=msg, diagnostic_code=code),
                    diagnostics=[ConversionDiagnostic(level="error", message=msg, code=code)],
                )
            source_ref = context.request.input_refs[0]
            proxy_request = ConversionRequest(
                request_id=context.request.request_id,
                input_refs=[
                    FileRef(
                        path=png_path,
                        format="png",
                        category="image",
                        encoding=source_ref.encoding,
                        warning_message=source_ref.warning_message,
                        size_bytes=Path(png_path).stat().st_size,
                        metadata=dict(source_ref.metadata),
                    )
                ],
                target_format=context.request.target_format,
                action_name=context.request.action_name,
                options=dict(context.request.options),
                output_policy=context.request.output_policy,
                config_snapshot=dict(context.request.config_snapshot),
            )
            converter_context = HubConversionContext(
                base=context,
                request=proxy_request,
                workspace=HubWorkspaceHandle(context.workspace, png_path),
            )

        target = converter_context.request.target_format
        if target == "md":
            return ImageToMarkdownConverter().convert(converter_context)
        if target == "pdf":
            return ImageToPdfConverter().convert(converter_context)
        return ImageFormatConverter().convert(converter_context)
