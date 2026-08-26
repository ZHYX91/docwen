"""Image to PDF converter."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import TYPE_CHECKING

from PIL import Image

from docwen_core.paths import input_stem
from docwen_plugin_image._common import (
    file_size,
    new_artifact_id,
    save_image_with_options,
    source_format_from_context,
)

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


def _layout_fun(quality_mode: str, is_landscape: bool):
    import img2pdf

    if quality_mode == "original":
        return None
    if quality_mode == "a4":
        pagesize = (
            (img2pdf.mm_to_pt(297), img2pdf.mm_to_pt(210))
            if is_landscape
            else (img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297))
        )
        return img2pdf.get_layout_fun(pagesize)
    if quality_mode == "a3":
        pagesize = (
            (img2pdf.mm_to_pt(420), img2pdf.mm_to_pt(297))
            if is_landscape
            else (img2pdf.mm_to_pt(297), img2pdf.mm_to_pt(420))
        )
        return img2pdf.get_layout_fun(pagesize)
    return None


class ImageToPdfConverter:
    """Convert an image to PDF, writing only to staging."""

    def convert(self, context: ConverterContext):
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionMetrics,
            ConversionResult,
        )

        task_id = context.request.request_id
        input_path = context.workspace.input_path
        options = context.request.options
        quality_mode = options.get("quality_mode") or "original"

        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting image to PDF conversion")

        mpo_frame_count = 0
        try:
            import img2pdf

            with tempfile.TemporaryDirectory() as temp_dir:
                source_for_pdf = input_path
                source_format = source_format_from_context(context)
                if source_format == "bmp":
                    source_for_pdf = str(Path(temp_dir) / f"{input_stem(input_path)}.png")
                    with Image.open(input_path) as img:
                        save_image_with_options(img, source_for_pdf, "png", {"compress_mode": "lossless"})

                with Image.open(source_for_pdf) as img:
                    width, height = img.size
                    is_landscape = width > height
                    if img.format == "MPO":
                        mpo_frame_count = int(getattr(img, "n_frames", 1))

                layout = _layout_fun(str(quality_mode), is_landscape)
                if layout is None:
                    pdf_bytes = img2pdf.convert(source_for_pdf)
                else:
                    pdf_bytes = img2pdf.convert(source_for_pdf, layout_fun=layout)
                if pdf_bytes is None:
                    raise RuntimeError("img2pdf did not return PDF bytes")

            output_path = context.workspace.create_artifact_path("primary", ".pdf")
            Path(output_path).write_bytes(pdf_bytes)
        except Exception as exc:
            context.logger.error(f"Image to PDF failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed", message=str(exc), diagnostic_code="IMG2PDF-ERROR"
                ),
                diagnostics=[
                    ConversionDiagnostic(level="error", message=f"Image to PDF failed: {exc}", code="IMG2PDF-ERROR")
                ],
            )

        artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind="primary",
            staging_path=output_path,
            suggested_name=f"{input_stem(input_path)}.pdf",
            media_type="application/pdf",
            metadata={"quality_mode": quality_mode},
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)
        context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)
        context.progress.report_progress(100.0, "Image to PDF complete")

        diagnostics = [ConversionDiagnostic(level="info", message="Converted image to PDF", code="IMG2PDF-OK")]
        if mpo_frame_count > 1:
            diagnostics.append(
                ConversionDiagnostic(
                    level="warning",
                    message=(
                        f"Delivered all {mpo_frame_count} MPO frames as PDF pages. "
                        "Auxiliary, gain-map, or secondary frames may look different from the primary photo; "
                        "review every page."
                    ),
                    code="IMG2PDF-MPO-AUXILIARY-FRAMES",
                )
            )

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact],
            diagnostics=diagnostics,
            metrics=ConversionMetrics(
                input_bytes=file_size(input_path),
                output_bytes=file_size(output_path),
                extra={"quality_mode": quality_mode, "mpo_frame_count": mpo_frame_count},
            ),
        )
