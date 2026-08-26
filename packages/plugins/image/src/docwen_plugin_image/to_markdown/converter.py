"""Image to Markdown converter."""

from __future__ import annotations

import logging
import shutil
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.errors import CancellationRequested
from docwen_core.export_semantics import (
    format_image_link,
    get_markdown_export_modes,
    resolve_markdown_request_policy,
)
from docwen_core.links import format_image_placeholder, normalize_link_target
from docwen_core.markdown_utils import (
    format_sanitized_image_link,
    sanitize_filename,
)
from docwen_core.paths import input_stem, normalize_path
from docwen_core.text.image_markdown import build_base64_image_data_uri, build_image_ocr_sidecar
from docwen_core.text.ocr import (
    OcrOutcome,
    OcrStatus,
    format_ocr_best_effort_warning,
    run_ocr_outcome,
)
from docwen_core.yaml_tools import extract_yaml, generate_basic_yaml_frontmatter
from docwen_plugin_image._common import (
    file_size,
    media_type_for,
    new_artifact_id,
    save_image_with_options,
    source_format_from_context,
)
from docwen_plugin_image.to_markdown.ocr_text import split_ocr_heading_body

_logger = logging.getLogger(__name__)

_IMAGE_LINK_STYLES = frozenset({"wiki_embed", "wiki_link", "markdown_embed", "markdown_link"})

if TYPE_CHECKING:
    from docwen_core.models.result import ConversionResult
    from docwen_core.protocols.execution_context import ConverterContext


def _coerce_bool(value: object, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"1", "true", "yes", "on"}:
            return True
        if normalized in {"0", "false", "no", "off"}:
            return False
    return bool(value)


def _option_or_config(
    options: dict[str, object],
    option_key: str,
    context: ConverterContext,
    config_key: str,
    default: object,
) -> object:
    if option_key in options:
        return options[option_key]
    return context.config.get(config_key, default)


def _convert_tiff_physical_pages(
    context: ConverterContext,
    *,
    input_path: str,
    keep_images: bool,
    enable_ocr: bool,
    ocr_language: str,
    current_locale: str,
) -> ConversionResult:
    """Emit one typed OCR fragment and optional PNG resource per TIFF frame."""
    from PIL import Image

    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import (
        ConversionDiagnostic,
        ConversionErrorInfo,
        ConversionMetrics,
        ConversionResult,
    )

    task_id = context.request.request_id
    input_stem_value = input_stem(input_path)
    created_paths: list[Path] = []
    page_artifacts: list[ArtifactManifest] = []
    image_artifacts: list[ArtifactManifest] = []
    diagnostics: list[ConversionDiagnostic] = []
    ocr_chars = 0

    try:
        with Image.open(input_path) as source:
            physical_page_count = int(getattr(source, "n_frames", 1))
            if physical_page_count < 1:
                raise ValueError("TIFF must contain at least one physical frame")

            for zero_index in range(physical_page_count):
                context.cancellation.check()
                page_number = zero_index + 1
                source.seek(zero_index)
                frame = source.copy()
                frame_path: Path | None = None
                retain_frame = False
                try:
                    frame.load()
                    if enable_ocr or keep_images:
                        frame_path = Path(context.workspace.create_artifact_path("image", ".png"))
                        created_paths.append(frame_path)
                        save_image_with_options(frame, str(frame_path), "png", {})

                    if enable_ocr:
                        assert frame_path is not None
                        context.cancellation.check()
                        try:
                            outcome = run_ocr_outcome(
                                str(frame_path),
                                source_format="png",
                                ocr_language=ocr_language,
                                current_locale=current_locale,
                            )
                        except CancellationRequested:
                            raise
                        except Exception as exc:
                            outcome = OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))
                        context.cancellation.check()

                        page_path = Path(context.workspace.create_artifact_path("auxiliary", ".md"))
                        created_paths.append(page_path)
                        page_text = outcome.recognized_text.rstrip()
                        page_path.write_bytes(f"{page_text}\n".encode() if page_text else b"")
                        page_artifact = ArtifactManifest(
                            artifact_id=new_artifact_id(),
                            kind="auxiliary",
                            staging_path=str(page_path),
                            suggested_name=f"{input_stem_value}__page_{page_number:04d}_ocr.md",
                            media_type="text/markdown",
                            metadata={
                                "fragment_kind": "page",
                                "page_index": page_number,
                                "page_count": physical_page_count,
                                "source_page": page_number,
                                "ocr_status": outcome.status.value,
                            },
                            is_primary=False,
                        )
                        page_artifacts.append(page_artifact)
                        ocr_chars += len(outcome.recognized_text)
                        if warning := format_ocr_best_effort_warning(outcome.status):
                            location = f"{Path(input_path).name}:frame-{page_number}"
                            diagnostics.append(
                                ConversionDiagnostic(
                                    level="warning",
                                    message=warning,
                                    code="OCR-BEST-EFFORT",
                                    location=location,
                                    artifact_id=page_artifact.artifact_id,
                                )
                            )

                    if keep_images:
                        assert frame_path is not None
                        retain_frame = True
                        image_artifact = ArtifactManifest(
                            artifact_id=new_artifact_id(),
                            kind="image",
                            staging_path=str(frame_path),
                            suggested_name=f"{input_stem_value}__page_{page_number:04d}.png",
                            media_type="image/png",
                            metadata={"source_format": "tif", "source_page": page_number},
                            is_primary=False,
                        )
                        image_artifacts.append(image_artifact)
                finally:
                    frame.close()
                    if frame_path is not None and not retain_frame:
                        try:
                            frame_path.unlink(missing_ok=True)
                        except OSError:
                            _logger.warning("Unable to remove request-owned TIFF OCR frame %s", frame_path)

        yaml_frontmatter = generate_basic_yaml_frontmatter(
            input_stem_value,
            extra={"source_format": "tif"},
            yaml_key_labels=context.request.options.get("yaml_key_labels"),
        )
        primary_path = Path(context.workspace.create_artifact_path("primary", ".md"))
        created_paths.append(primary_path)
        primary_path.write_text(yaml_frontmatter, encoding="utf-8")
        primary = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind="primary",
            staging_path=str(primary_path),
            suggested_name=f"{input_stem_value}.md",
            media_type="text/markdown",
            metadata={
                "source_format": "tif",
                "physical_page_count": physical_page_count,
                "keep_images": keep_images,
                "ocr_enabled": enable_ocr,
            },
            is_primary=True,
        )
        artifacts = [primary, *page_artifacts, *image_artifacts]
        for artifact in artifacts:
            context.workspace.add_artifact(artifact)
            context.progress.report_artifact_ready(artifact.artifact_id, artifact.suggested_name)
        diagnostics.insert(
            0, ConversionDiagnostic(level="info", message="Converted TIFF to Markdown", code="IMG2MD-OK")
        )
        if enable_ocr:
            diagnostics.insert(
                1,
                ConversionDiagnostic(
                    level="info",
                    message=f"OCR processed {physical_page_count} physical frame(s), {ocr_chars} total characters",
                    code="IMG2MD-OCR-OK",
                ),
            )
        context.progress.report_progress(100.0, "TIFF to Markdown complete")
        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=artifacts,
            diagnostics=diagnostics,
            metrics=ConversionMetrics(
                input_bytes=file_size(input_path),
                output_bytes=sum(file_size(artifact.staging_path) for artifact in artifacts),
                extra={
                    "artifact_count": len(artifacts),
                    "physical_page_count": physical_page_count,
                    "ocr_enabled": enable_ocr,
                    "ocr_pages": len(page_artifacts),
                    "image_count": len(image_artifacts),
                },
            ),
        )
    except CancellationRequested:
        for path in created_paths:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                _logger.warning("Unable to remove cancelled TIFF artifact %s", path)
        raise
    except Exception as exc:
        for path in created_paths:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                _logger.warning("Unable to remove failed TIFF artifact %s", path)
        context.logger.error(f"TIFF to Markdown failed: {exc}")
        return ConversionResult(
            task_id=task_id,
            success=False,
            error=ConversionErrorInfo(
                error_type="conversion_failed",
                message=str(exc),
                diagnostic_code="IMG2MD-ERROR",
            ),
            diagnostics=[ConversionDiagnostic(level="error", message="TIFF to Markdown failed", code="IMG2MD-ERROR")],
        )


class ImageToMarkdownConverter:
    """Create a Markdown file referencing or embedding image assets.

    When OCR is enabled the recognised text is appended as a blockquote
    (``> `` lines) after the image link in the Markdown output.
    """

    def convert(self, context: ConverterContext):
        from docwen_core.models.artifact import ArtifactManifest
        from docwen_core.models.result import (
            ConversionDiagnostic,
            ConversionErrorInfo,
            ConversionMetrics,
            ConversionResult,
        )

        task_id = context.request.request_id
        input_path = normalize_path(context.workspace.input_path)
        options = context.request.options
        keep_images = _coerce_bool(
            _option_or_config(options, "to_md_keep_images", context, "image.to_md_keep_images", True),
            True,
        )
        enable_ocr = _coerce_bool(
            _option_or_config(options, "to_md_enable_ocr", context, "image.to_md_enable_ocr", True),
            True,
        )
        image_mode = str(
            _option_or_config(options, "image_mode", context, "image.to_md_image_extraction_mode", "file") or "file"
        )
        ocr_language = str(_option_or_config(options, "ocr_language", context, "image.ocr_language", "auto") or "auto")
        current_locale = str(options.get("locale") or "zh_CN")

        context.cancellation.check()
        context.progress.report_progress(0.0, "Starting image to Markdown conversion")

        artifacts: list[ArtifactManifest] = []
        source_format = source_format_from_context(context)
        image_filename = sanitize_filename(f"{input_stem(input_path)}.{source_format}")

        if source_format == "tif":
            return _convert_tiff_physical_pages(
                context,
                input_path=input_path,
                keep_images=keep_images,
                enable_ocr=enable_ocr,
                ocr_language=ocr_language,
                current_locale=current_locale,
            )

        # ── OCR ────────────────────────────────────────────────────
        ocr_text = ""
        if enable_ocr:
            context.progress.report_progress(5.0, "Running OCR on image")
            try:
                outcome = run_ocr_outcome(
                    input_path,
                    source_format=source_format,
                    ocr_language=ocr_language,
                    current_locale=current_locale,
                )
            except Exception as exc:
                outcome = OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))

            if message := format_ocr_best_effort_warning(outcome.status):
                context.logger.warning(f"{message} {outcome.message}".rstrip())
                context.progress.report_diagnostic(
                    "warning",
                    message,
                    code="OCR-BEST-EFFORT",
                    location=Path(input_path).name,
                )
            ocr_text = outcome.recognized_text
            context.progress.report_progress(15.0, "OCR complete")

        def _format_ocr_blockquote(text: str) -> str:
            if not text:
                return ""
            lines_out: list[str] = []
            # Prepend OCR blockquote title when configured (F-G1-005, F-G2-003)
            if ocr_title:
                lines_out.append(f"> **{ocr_title}**")
                lines_out.append(">")
            for line in text.splitlines():
                stripped = line.strip()
                if not stripped:
                    lines_out.append(">")
                    continue
                heading, body = split_ocr_heading_body(stripped)
                if body:
                    lines_out.append(f"> **{heading}** {body}")
                else:
                    lines_out.append(f"> {heading}")
            return "\n".join(lines_out)

        try:
            context.cancellation.check()

            # Resolve link style and export modes from export semantics so the
            # output respects user preferences (wiki_embed / markdown_embed, etc.).
            request_policy = resolve_markdown_request_policy(context)
            export_semantics = request_policy.export
            request_image_style = str(options.get("image_link_style") or "").strip().lower()
            image_style = (
                request_image_style if request_image_style in _IMAGE_LINK_STYLES else export_semantics.image_link_style
            )
            md_file_style = export_semantics.md_file_link_style
            # Allow per-request override of ocr_placement_mode via options
            # (e.g. {"ocr_placement": "main_md"}) — falls back to semantics default.
            export_modes = get_markdown_export_modes(
                "image",
                extraction_mode=image_mode,
                ocr_placement_mode=options.get("ocr_placement"),
                semantics=export_semantics,
            )
            ocr_placement_mode = export_modes["ocr_placement_mode"]  # "main_md" or "image_md"
            ocr_title = request_policy.ocr_blockquote_title

            if image_mode == "omit" or (not keep_images and image_mode != "base64" and image_mode != "embed"):
                link = f"<!-- image omitted: {image_filename} -->"
            elif image_mode == "embed":
                link = format_image_placeholder("./" + image_filename)
            elif image_mode == "base64":
                target = build_base64_image_data_uri(
                    image_path=input_path,
                    media_type=media_type_for(source_format),
                    export_semantics=export_semantics,
                )
                link = format_image_link(input_stem(input_path), target, style=image_style)
            elif keep_images:
                image_staging_path = context.workspace.create_artifact_path("image", f".{source_format}")
                shutil.copyfile(input_path, image_staging_path)
                image_artifact = ArtifactManifest(
                    artifact_id=new_artifact_id(),
                    kind="image",
                    staging_path=image_staging_path,
                    suggested_name=image_filename,
                    media_type=media_type_for(source_format),
                    metadata={"source_format": source_format},
                    is_primary=False,
                )
                context.workspace.add_artifact(image_artifact)
                artifacts.append(image_artifact)
                # Wiki styles route through the sanitised formatter
                # (F-I2b-004: production consumer of format_sanitized_image_link).
                if image_style in ("wiki_embed", "wiki_link"):
                    link = format_sanitized_image_link(image_filename, style=image_style)
                else:
                    target = normalize_link_target("./" + image_filename)
                    link = format_image_link(input_stem(input_path), target, style=image_style)
            else:
                raise ValueError(f"Unexpected state: image_mode={image_mode!r}, keep_images={keep_images!r}")

            # ── Build Markdown ─────────────────────────────────────
            yaml_frontmatter = generate_basic_yaml_frontmatter(
                input_stem(input_path),
                extra={"source_format": source_format},
                yaml_key_labels=options.get("yaml_key_labels"),
            )
            lines: list[str] = [yaml_frontmatter.rstrip(), "", link]

            # Append OCR blockquote after the image link (main_md mode),
            # or generate a separate per-image .md file (image_md mode).
            # F-G1-005, F-G2-003: restored ocr_placement_mode support.
            if enable_ocr and ocr_placement_mode == "image_md":
                # Build sidecar .md via shared core helper (image + OCR).
                sidecar_stem = f"{input_stem(input_path)}_ocr"
                sidecar_text, replacement_link = build_image_ocr_sidecar(
                    sidecar_stem=sidecar_stem,
                    source_format=source_format,
                    image_markdown=link,
                    ocr_text=ocr_text,
                    md_link_style=md_file_style,
                    ocr_blockquote_title=ocr_title,
                    yaml_key_labels=options.get("yaml_key_labels"),
                )

                # Write sidecar to staging and register as artifact.
                img_md_path = context.workspace.create_artifact_path("auxiliary", ".md")
                Path(img_md_path).write_text(sidecar_text, encoding="utf-8")

                ocr_md_filename = sanitize_filename(f"{sidecar_stem}.md")
                ocr_md_artifact = ArtifactManifest(
                    artifact_id=new_artifact_id(),
                    kind="auxiliary",
                    staging_path=img_md_path,
                    suggested_name=ocr_md_filename,
                    media_type="text/markdown",
                    metadata={"source_format": source_format, "ocr": True},
                    is_primary=False,
                )
                context.workspace.add_artifact(ocr_md_artifact)
                artifacts.append(ocr_md_artifact)

                # Replace image link with .md file link in primary output.
                lines = [yaml_frontmatter.rstrip(), "", replacement_link]
            elif ocr_text:
                # main_md (default): append OCR blockquote inline.
                ocr_block = _format_ocr_blockquote(ocr_text)
                if ocr_block:
                    lines.append("")
                    lines.append(ocr_block)
                    lines.append("")

            md_text = "\n".join(lines)

            # ── Post-generation validation ────────────────────────
            # Validate that the generated YAML front matter is well-formed
            # by extracting it back (F-I2b-001: production consumer of extract_yaml).
            try:
                yaml_extracted, _ = extract_yaml(md_text)
                if not yaml_extracted:
                    _logger.warning("Generated Markdown has no parseable YAML front matter")
                elif input_stem(input_path) not in yaml_extracted:
                    _logger.warning("Generated YAML front matter may be malformed")
            except Exception:
                _logger.warning("YAML front matter validation failed", exc_info=True)

            md_path = context.workspace.create_artifact_path("primary", ".md")
            Path(md_path).write_text(md_text, encoding="utf-8")
        except Exception as exc:
            context.logger.error(f"Image to Markdown failed: {exc}")
            return ConversionResult(
                task_id=task_id,
                success=False,
                error=ConversionErrorInfo(
                    error_type="conversion_failed", message=str(exc), diagnostic_code="IMG2MD-ERROR"
                ),
                diagnostics=[
                    ConversionDiagnostic(level="error", message=f"Image to Markdown failed: {exc}", code="IMG2MD-ERROR")
                ],
            )

        md_artifact = ArtifactManifest(
            artifact_id=new_artifact_id(),
            kind="primary",
            staging_path=md_path,
            suggested_name=f"{input_stem(input_path)}.md",
            media_type="text/markdown",
            metadata={"image_mode": image_mode, "keep_images": keep_images, "ocr_enabled": enable_ocr},
            is_primary=True,
        )
        context.workspace.add_artifact(md_artifact)
        artifacts.insert(0, md_artifact)

        context.progress.report_artifact_ready(md_artifact.artifact_id, md_artifact.suggested_name)
        context.progress.report_progress(100.0, "Image to Markdown complete")

        diagnostics = [ConversionDiagnostic(level="info", message="Converted image to Markdown", code="IMG2MD-OK")]
        if ocr_text:
            diagnostics.append(
                ConversionDiagnostic(
                    level="info", message=f"OCR extracted {len(ocr_text)} characters", code="IMG2MD-OCR-OK"
                )
            )

        return ConversionResult(
            task_id=task_id,
            success=True,
            artifacts=artifacts,
            diagnostics=diagnostics,
            metrics=ConversionMetrics(
                input_bytes=file_size(input_path),
                output_bytes=sum(file_size(a.staging_path) for a in artifacts),
                extra={"artifact_count": len(artifacts), "ocr_enabled": enable_ocr, "ocr_chars": len(ocr_text)},
            ),
        )
