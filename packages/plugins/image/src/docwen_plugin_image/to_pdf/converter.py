"""Image to PDF converter."""

from __future__ import annotations

from contextlib import suppress
from pathlib import Path
from typing import TYPE_CHECKING

from PIL import Image

from docwen_core.errors import CancellationRequested
from docwen_plugin_image._common import file_size, new_artifact_id

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


def _layout_fun(quality_mode: str):
    import img2pdf

    paper = {"a4": (210, 297), "a3": (297, 420)}.get(quality_mode)
    if paper is None:
        return img2pdf.default_layout_fun
    width, height = paper
    return img2pdf.get_layout_fun((img2pdf.mm_to_pt(width), img2pdf.mm_to_pt(height)), auto_orient=True)


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
        output_path = ""
        complete = False
        try:
            import img2pdf

            with Image.open(input_path) as img:
                if img.format == "MPO":
                    mpo_frame_count = int(getattr(img, "n_frames", 1))

            # Let img2pdf read the original image's physical resolution. Its
            # layout callback is evaluated for each page of a multiframe input.
            output_path = context.workspace.create_artifact_path("primary", ".pdf")
            with Path(output_path).open("wb") as destination:
                img2pdf.convert(input_path, layout_fun=_layout_fun(str(quality_mode)), outputstream=destination)
            context.cancellation.check()
            complete = True
        except CancellationRequested:
            raise
        except Exception:
            context.logger.error("Image to PDF failed")
            message = "Image to PDF conversion failed."
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed", message=message, diagnostic_code="IMG2PDF-ERROR"
                ),
                diagnostics=[ConversionDiagnostic(level="error", message=message, code="IMG2PDF-ERROR")],
            )

        finally:
            if output_path and not complete:
                with suppress(OSError):
                    Path(output_path).unlink(missing_ok=True)

        artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind="primary",
            staging_path=output_path,
            suggested_name=f"{context.request.source_stem}.pdf",
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
