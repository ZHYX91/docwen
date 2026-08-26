"""Layout → PNG / JPG / TIF converters using PyMuPDF rendering."""

from __future__ import annotations

import contextlib
import time
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models.artifact import (
    ARTIFACT_KIND_AUXILIARY,
    ARTIFACT_KIND_IMAGE,
    ARTIFACT_KIND_PRIMARY,
    ArtifactManifest,
)
from docwen_core.models.result import (
    ConversionDiagnostic,
    ConversionErrorInfo,
    ConversionMetrics,
    ConversionResult,
)
from docwen_core.paths import input_stem
from docwen_plugin_layout._common import file_size, new_artifact_id, request_source_format

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


# ── Media types ────────────────────────────────────────────────────────

MEDIA_TYPES: dict[str, str] = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "tif": "image/tiff",
}


class LayoutToImageConverter:
    """Renders a layout document (PDF) to images using PyMuPDF.

    Each page is rendered as a separate image file. For TIF output,
    all pages are combined into a single multi-page TIFF.
    """

    def __init__(self, target_format: str) -> None:
        if target_format not in ("png", "jpg", "tif"):
            raise ValueError(f"Unsupported target format: {target_format}")
        self._target_format: str = target_format

    def convert(
        self,
        context: ConverterContext,
        *,
        input_path_override: str | None = None,
        source_format_override: str | None = None,
    ) -> ConversionResult:
        task_id = context.request.request_id
        start_time = time.monotonic()
        input_file = input_path_override or context.workspace.input_path
        # The override is an internal preprocessed PDF. User-visible artifact
        # names must remain owned by the original request input, not leak the
        # random staging stem (for example ``_preprocess_xps_<uuid>``).
        stem = input_stem(context.workspace.input_path)

        options = context.request.options or {}
        dpi = int(options.get("render_dpi", 150))

        # ── Cancellation check ──────────────────────────────────────
        try:
            context.cancellation.check()
        except Exception:
            return self._cancelled(task_id)

        # ── Import PyMuPDF ──────────────────────────────────────────
        try:
            import fitz  # PyMuPDF
        except ImportError:
            return self._dependency_missing(
                task_id,
                "PyMuPDF (fitz) is required for layout→image conversion. Install with: pip install PyMuPDF",
            )

        # ── Determine / validate input format ────────────────────────
        input_fmt = str(source_format_override or request_source_format(context)).strip().lower()
        if input_fmt not in ("pdf",):
            # Non-PDF layout formats (OFD/XPS) should be preprocessed
            # to PDF before reaching this converter.
            return self._invalid_input(
                task_id,
                f"layout ({input_fmt})→{self._target_format}",
                f"expected a preprocessed PDF, received {input_fmt}",
            )

        try:
            doc = fitz.open(input_file, filetype="pdf")
        except Exception as exc:
            return self._conversion_error(
                task_id,
                "LAYOUT2IMG-OPEN-ERROR",
                f"Failed to open {input_fmt.upper()} file: {exc}",
            )

        total_pages = len(doc)
        input_bytes = file_size(input_file)

        try:
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)

            if self._target_format in ("png", "jpg"):
                # ── Per-page image output ──────────────────────────
                artifacts: list[ArtifactManifest] = []
                width = max(len(str(total_pages)), 2)
                total_output_bytes = 0
                use_alpha = self._target_format == "png"

                for page_num in range(total_pages):
                    try:
                        context.cancellation.check()
                    except Exception:
                        doc.close()
                        return self._cancelled(task_id)

                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(matrix=mat, alpha=use_alpha)

                    image_filename = f"{stem}_page_{str(page_num + 1).zfill(width)}.{self._target_format}"
                    staging_path = context.workspace.create_artifact_path(
                        ARTIFACT_KIND_IMAGE, f".{self._target_format}"
                    )

                    pix.save(staging_path)
                    output_bytes = file_size(staging_path)
                    total_output_bytes += output_bytes

                    artifact = ArtifactManifest(
                        artifact_id=new_artifact_id(),
                        kind=ARTIFACT_KIND_IMAGE,
                        staging_path=staging_path,
                        suggested_name=image_filename,
                        media_type=MEDIA_TYPES[self._target_format],
                        metadata={
                            "page": page_num + 1,
                            "total_pages": total_pages,
                            "width": pix.width,
                            "height": pix.height,
                            "dpi": dpi,
                        },
                        is_primary=(page_num == 0),
                    )
                    context.workspace.add_artifact(artifact)
                    artifacts.append(artifact)

                    context.progress.report_progress(
                        (page_num + 1) / total_pages * 100.0,
                        f"Rendered page {page_num + 1}/{total_pages}",
                    )

                doc.close()
                elapsed_ms = (time.monotonic() - start_time) * 1000.0
                return ConversionResult(
                    task_id=task_id,
                    success=True,
                    artifacts=artifacts,
                    diagnostics=[
                        ConversionDiagnostic(
                            level="info",
                            message=f"Rendered {total_pages} page(s) to {self._target_format.upper()} at {dpi} DPI",
                            code="LAYOUT2IMG-RENDERED",
                        )
                    ],
                    metrics=ConversionMetrics(
                        duration_ms=elapsed_ms,
                        input_bytes=input_bytes,
                        output_bytes=total_output_bytes,
                        extra={"page_count": total_pages, "dpi": dpi},
                    ),
                )

            else:  # tif — multi-page TIFF
                from PIL import Image

                page_images: list[str] = []

                for page_num in range(total_pages):
                    try:
                        context.cancellation.check()
                    except Exception:
                        doc.close()
                        # Clean up temp page images
                        for p in page_images:
                            with contextlib.suppress(OSError):
                                Path(p).unlink(missing_ok=True)
                        return self._cancelled(task_id)

                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(matrix=mat, alpha=False)

                    # Save intermediate PNG for multi-page TIFF assembly
                    inter_path = context.workspace.create_artifact_path(ARTIFACT_KIND_AUXILIARY, ".png")
                    pix.save(inter_path)
                    page_images.append(inter_path)

                doc.close()

                # Combine into multi-page TIFF
                tif_staging = context.workspace.create_artifact_path(ARTIFACT_KIND_PRIMARY, ".tif")

                if total_pages == 1:
                    with Image.open(page_images[0]) as img:
                        img.save(tif_staging, format="TIFF", compression="tiff_lzw")
                else:
                    opened: list[Image.Image] = []
                    try:
                        first = Image.open(page_images[0])
                        opened.append(first)
                        rest = [Image.open(p) for p in page_images[1:]]
                        opened.extend(rest)
                        first.save(
                            tif_staging,
                            format="TIFF",
                            save_all=True,
                            append_images=rest,
                            compression="tiff_lzw",
                        )
                    finally:
                        for im in opened:
                            with contextlib.suppress(Exception):
                                im.close()

                # Clean up intermediate PNGs
                for p in page_images:
                    with contextlib.suppress(OSError):
                        Path(p).unlink(missing_ok=True)

                output_bytes = file_size(tif_staging)
                artifact = ArtifactManifest(
                    artifact_id=new_artifact_id(),
                    kind=ARTIFACT_KIND_PRIMARY,
                    staging_path=tif_staging,
                    suggested_name=f"{stem}.tif",
                    media_type=MEDIA_TYPES["tif"],
                    metadata={
                        "page_count": total_pages,
                        "dpi": dpi,
                    },
                    is_primary=True,
                )
                context.workspace.add_artifact(artifact)

                elapsed_ms = (time.monotonic() - start_time) * 1000.0
                return ConversionResult(
                    task_id=task_id,
                    success=True,
                    artifacts=[artifact],
                    diagnostics=[
                        ConversionDiagnostic(
                            level="info",
                            message=f"Rendered {total_pages} page(s) to multi-page TIF at {dpi} DPI",
                            code="LAYOUT2IMG-RENDERED",
                        )
                    ],
                    metrics=ConversionMetrics(
                        duration_ms=elapsed_ms,
                        input_bytes=input_bytes,
                        output_bytes=output_bytes,
                        extra={"page_count": total_pages, "dpi": dpi},
                    ),
                )

        except Exception as exc:
            with contextlib.suppress(Exception):
                doc.close()
            return self._conversion_error(
                task_id,
                "LAYOUT2IMG-RENDER-ERROR",
                f"Failed to render pages: {exc}",
            )

    # ── Helpers ──────────────────────────────────────────────────────

    @staticmethod
    def _cancelled(task_id: str) -> ConversionResult:
        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="cancelled",
                message="Conversion was cancelled by the user.",
                diagnostic_code="LAYOUT2IMG-CANCELLED",
            ),
        )

    @staticmethod
    def _dependency_missing(task_id: str, message: str) -> ConversionResult:
        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="dependency_missing",
                message=message,
                diagnostic_code="LAYOUT2IMG-DEPENDENCY-MISSING",
            ),
        )

    @staticmethod
    def _invalid_input(task_id: str, route_label: str, reason: str) -> ConversionResult:
        msg = f"Invalid input for {route_label}: {reason}."
        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="invalid_input",
                message=msg,
                diagnostic_code="LAYOUT2IMG-INVALID-INPUT",
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=msg,
                    code="LAYOUT2IMG-INVALID-INPUT",
                )
            ],
        )

    @staticmethod
    def _conversion_error(task_id: str, code: str, message: str) -> ConversionResult:
        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message=message,
                diagnostic_code=code,
            ),
            diagnostics=[
                ConversionDiagnostic(
                    level="error",
                    message=message,
                    code=code,
                )
            ],
        )
