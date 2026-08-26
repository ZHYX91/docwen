"""Merge images into a multipage TIFF."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from PIL import Image

from docwen_plugin_image._common import file_size, has_alpha, new_artifact_id, paste_on_white

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


def _load_image(path: str) -> Image.Image:
    try:
        with Image.open(path) as img:
            copy = img.copy()
            copy.load()
            return copy
    except Exception as exc:
        raise RuntimeError(f"Failed to load image '{Path(path).name}': {exc}") from exc


def _to_rgb_with_white(img: Image.Image) -> Image.Image:
    if has_alpha(img):
        return paste_on_white(img)
    if img.mode != "RGB":
        return img.convert("RGB")
    return img.copy()


class ImageToTiffMerger:
    """Merge request input_refs into a single multipage TIFF artifact."""

    def convert(self, context: ConverterContext):
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionMetrics,
            ConversionResult,
        )

        task_id = context.request.request_id
        input_refs = context.request.input_refs
        input_paths = [ref.path for ref in input_refs]
        options = context.request.options
        mode = str(options.get("mode") or "smart").lower()
        if options.get("keep_alpha") is False:
            mode = "rgb"

        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting image merge to TIFF")

        if not input_paths:
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="invalid_input", message="No images to merge", diagnostic_code="IMG2TIFF-NO-INPUT"
                ),
                diagnostics=[
                    ConversionDiagnostic(level="error", message="No images to merge", code="IMG2TIFF-NO-INPUT")
                ],
            )

        loaded: list[Image.Image] = []
        converted: list[Image.Image] = []
        try:
            for idx, path in enumerate(input_paths, 1):
                context.cancellation.check()
                context.progress.report_progress(
                    20.0 * idx / max(len(input_paths), 1), f"Loading image {idx}/{len(input_paths)}"
                )
                loaded.append(_load_image(path))

            target_mode = "RGBA" if mode == "smart" and all(has_alpha(img) for img in loaded) else "RGB"
            for img in loaded:
                if target_mode == "RGBA":
                    converted.append(img.convert("RGBA") if img.mode != "RGBA" else img.copy())
                else:
                    converted.append(_to_rgb_with_white(img))

            output_path = context.workspace.create_artifact_path("primary", ".tif")
            converted[0].save(
                output_path,
                format="TIFF",
                save_all=True,
                append_images=converted[1:],
                compression="tiff_lzw",
            )
        except Exception as exc:
            context.logger.error(f"Image merge to TIFF failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed", message=str(exc), diagnostic_code="IMG2TIFF-ERROR"
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error", message=f"Image merge to TIFF failed: {exc}", code="IMG2TIFF-ERROR"
                    )
                ],
            )
        finally:
            seen: set[int] = set()
            for collection in (loaded, converted):
                for img in collection:
                    if id(img) in seen:
                        continue
                    seen.add(id(img))
                    img.close()

        primary_stem = Path(input_paths[0]).stem if input_paths else "merged_images"
        artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind="primary",
            staging_path=output_path,
            suggested_name=f"{primary_stem}_merged.tif",
            media_type="image/tiff",
            metadata={"image_count": len(input_paths), "mode": mode},
            is_primary=True,
        )
        context.workspace.add_artifact(artifact)
        context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)
        context.progress.report_progress(100.0, "Image merge to TIFF complete")

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=[artifact],
            diagnostics=[
                ConversionDiagnostic(
                    level="info", message=f"Merged {len(input_paths)} images to TIFF", code="IMG2TIFF-OK"
                )
            ],
            metrics=ConversionMetrics(
                input_bytes=sum(file_size(p) for p in input_paths),
                output_bytes=file_size(output_path),
                extra={"image_count": len(input_paths), "mode": mode},
            ),
        )
