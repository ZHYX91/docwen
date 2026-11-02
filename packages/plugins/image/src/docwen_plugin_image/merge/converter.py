"""Merge images into a multipage TIFF."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import TYPE_CHECKING

from PIL import Image

from docwen_core.errors import CancellationRequested
from docwen_core.models.file_ref import (
    MANAGED_INPUT_SHA256_METADATA_KEY,
    MANAGED_INPUT_SIZE_BYTES_METADATA_KEY,
    source_presentation_name,
)
from docwen_plugin_image._common import (
    file_size,
    has_alpha,
    is_heic_format,
    new_artifact_id,
    paste_on_white,
    register_heic_decoder,
)

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


MAX_MERGE_FRAMES = 1000
MAX_MERGE_PIXELS = 64 * 1024 * 1024


class ImageMergeLimitError(ValueError):
    """The complete merge cannot fit the bounded decoded-frame budget."""


def _load_image_frames(path: str, frames: list[Image.Image], pixels: int, cancel_check: Callable[[], None]) -> int:
    """Append each physical frame, checking the aggregate budget before decoding."""
    try:
        with Image.open(path) as img:
            index = 0
            while True:
                cancel_check()
                try:
                    img.seek(index)
                except EOFError:
                    break
                if len(frames) >= MAX_MERGE_FRAMES:
                    raise ImageMergeLimitError(
                        f"Image merge exceeds {MAX_MERGE_FRAMES} frames. Merge smaller groups of images."
                    )
                pixels += img.width * img.height
                if pixels > MAX_MERGE_PIXELS:
                    raise ImageMergeLimitError(
                        f"Image merge exceeds {MAX_MERGE_PIXELS} total decoded pixels. Merge smaller groups of images."
                    )
                frame = img.copy()
                frames.append(frame)
                frame.load()
                index += 1
        return pixels
    except (CancellationRequested, ImageMergeLimitError):
        raise
    except Exception as exc:
        raise RuntimeError("Image input could not be decoded.") from exc


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
        failed_input_name: str | None = None
        pixels = 0
        try:
            if any(is_heic_format(ref.format) for ref in input_refs):
                register_heic_decoder()
            for idx, path in enumerate(input_paths, 1):
                context.cancellation.check()
                context.progress.report_progress(
                    20.0 * idx / max(len(input_paths), 1), f"Loading image {idx}/{len(input_paths)}"
                )
                ref = input_refs[idx - 1]
                managed = (
                    MANAGED_INPUT_SHA256_METADATA_KEY in ref.metadata
                    and MANAGED_INPUT_SIZE_BYTES_METADATA_KEY in ref.metadata
                )
                failed_input_name = None if managed else source_presentation_name(ref)
                pixels = _load_image_frames(path, loaded, pixels, context.cancellation.check)

            failed_input_name = None

            target_mode = "RGBA" if mode == "smart" and any(has_alpha(img) for img in loaded) else "RGB"
            for img in loaded:
                context.cancellation.check()
                if target_mode == "RGBA":
                    converted.append(img.convert("RGBA") if img.mode != "RGBA" else img.copy())
                else:
                    converted.append(_to_rgb_with_white(img))

            output_path = context.workspace.create_artifact_path("primary", ".tif")
            context.cancellation.check()
            converted[0].save(
                output_path,
                format="TIFF",
                save_all=True,
                append_images=converted[1:],
                compression="tiff_lzw",
            )
        except CancellationRequested:
            raise
        except Exception as exc:
            context.logger.error("Image merge to TIFF failed")
            message = "Image merge to TIFF failed."
            if failed_input_name is not None:
                message = f"Failed to load image: {failed_input_name}."
            if isinstance(exc, ImageMergeLimitError):
                message = str(exc)
            code = "IMG2TIFF-LIMIT" if isinstance(exc, ImageMergeLimitError) else "IMG2TIFF-ERROR"
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(error_type="conversion_failed", message=message, diagnostic_code=code),
                diagnostics=[ConversionDiagnostic(level="error", message=message, code=code)],
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
            metadata={"image_count": len(input_paths), "frame_count": len(loaded), "mode": mode},
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
                    level="info",
                    message=f"Merged {len(input_paths)} images to TIFF ({len(loaded)} pages)",
                    code="IMG2TIFF-OK",
                )
            ],
            metrics=ConversionMetrics(
                input_bytes=sum(file_size(p) for p in input_paths),
                output_bytes=file_size(output_path),
                extra={"image_count": len(input_paths), "frame_count": len(loaded), "mode": mode},
            ),
        )
