"""Image format conversion converter."""

from __future__ import annotations

from typing import TYPE_CHECKING

from PIL import Image

from docwen_core.paths import input_stem
from docwen_plugin_image._common import (
    file_size,
    media_type_for,
    new_artifact_id,
    normalize_format,
    prepare_flat_export,
    save_image_with_options,
    source_format_from_context,
)

if TYPE_CHECKING:
    from docwen_core.protocols.execution_context import ConverterContext


class ImageFormatConverter:
    """Convert one image to another image format, writing only to staging."""

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
        target = normalize_format(context.request.target_format)
        if target == "image":
            target = normalize_format(str(context.request.options.get("target_format") or "png"))
        if target == "jpeg":
            target = "jpg"
        if target == "tiff":
            target = "tif"

        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting image format conversion")

        source_format = source_format_from_context(context)
        suffix = f".{target}"
        artifacts: list[ArtifactManifest] = []
        frames = 1
        diagnostics: list[ConversionDiagnostic] = []

        try:
            with Image.open(input_path) as img:
                n_frames = getattr(img, "n_frames", 1)

                if n_frames > 1 and ((source_format == "tif" and target != "tif") or target in ("jpg", "bmp")):
                    frames = n_frames
                    frame_label = "page" if source_format == "tif" else "frame"
                    if frame_label == "frame":
                        diagnostics.append(
                            ConversionDiagnostic(
                                level="warning",
                                code="IMAGEFMT-FRAMES-SPLIT",
                                message="The static target contains separate frames; animation playback timing is not retained.",
                            )
                        )
                    for idx in range(1, n_frames + 1):
                        context.cancellation.check()
                        output_path = context.workspace.create_artifact_path(
                            "primary" if idx == 1 else "auxiliary", suffix
                        )
                        img.seek(idx - 1)
                        with img.copy() as frame:
                            save_image_with_options(frame, output_path, target, context.request.options)
                        artifact = ArtifactManifest(
                            artifact_id=new_artifact_id(),
                            kind="primary" if idx == 1 else "auxiliary",
                            staging_path=output_path,
                            suggested_name=f"{input_stem(input_path)}_{frame_label}{idx}.{target}",
                            media_type=media_type_for(target),
                            metadata={
                                "source_format": source_format,
                                "target_format": target,
                                "page_index": idx - 1,
                            },
                            is_primary=(idx == 1),
                        )
                        context.workspace.add_artifact(artifact)
                        artifacts.append(artifact)
                else:
                    output_path = context.workspace.create_artifact_path("primary", suffix)
                    if n_frames > 1:
                        from docwen_plugin_image.format_conversion.sequence import save_sequence

                        if source_format != "tif" and target == "tif":
                            diagnostics.append(
                                ConversionDiagnostic(
                                    level="warning",
                                    code="IMAGEFMT-ANIMATION-TO-PAGES",
                                    message="All animation frames are retained as TIFF pages; playback timing is not retained.",
                                )
                            )
                        if img.info.get("default_image") and target != "png":
                            diagnostics.append(
                                ConversionDiagnostic(
                                    level="warning",
                                    code="IMAGEFMT-POSTER-OMITTED",
                                    message="The APNG poster is omitted; animation frames are retained.",
                                )
                            )
                        frames = save_sequence(
                            img, output_path, target, context.request.options, cancel_check=context.cancellation.check
                        )
                        width, height = img.size
                    elif target in ("jpg", "webp"):
                        img.load()
                        prepared, save_metadata = prepare_flat_export(img)
                        try:
                            save_image_with_options(
                                prepared,
                                output_path,
                                target,
                                context.request.options,
                                save_metadata=save_metadata,
                            )
                            width, height = prepared.size
                        finally:
                            prepared.close()
                    else:
                        save_image_with_options(img, output_path, target, context.request.options)
                        width, height = img.size
                    artifact = ArtifactManifest(
                        artifact_id=new_artifact_id(),
                        kind="primary",
                        staging_path=output_path,
                        suggested_name=f"{input_stem(input_path)}.{target}",
                        media_type=media_type_for(target),
                        metadata={"target_format": target, "width": width, "height": height},
                        is_primary=True,
                    )
                    context.workspace.add_artifact(artifact)
                    artifacts.append(artifact)
        except Exception as exc:
            context.logger.error(f"Image format conversion failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed",
                    message=str(exc),
                    diagnostic_code="IMAGEFMT-CONVERT-ERROR",
                ),
                diagnostics=[
                    ConversionDiagnostic(
                        level="error", message=f"Image conversion failed: {exc}", code="IMAGEFMT-CONVERT-ERROR"
                    )
                ],
            )

        context.progress.report_progress(100.0, "Image format conversion complete")
        if artifacts:
            context.progress.report_artifact_ready(artifacts[0].artifact_id, artifacts[0].suggested_name)

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=artifacts,
            diagnostics=[
                *diagnostics,
                ConversionDiagnostic(level="info", message=f"Converted image to {target.upper()}", code="IMAGEFMT-OK"),
            ],
            metrics=ConversionMetrics(
                input_bytes=file_size(input_path),
                output_bytes=sum(file_size(a.staging_path) for a in artifacts),
                extra={"artifact_count": len(artifacts), "frame_count": frames, "target_format": target},
            ),
        )
