"""Markdown resource writeback helpers for markup converters."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from urllib.parse import unquote, urlsplit

from PIL import Image

from docwen_core.detection import detect_content_format
from docwen_core.export_semantics import (
    MarkdownExportSemantics,
    format_image_link,
    resolve_markdown_request_policy,
)
from docwen_core.formats import CATEGORY_IMAGE, get_category, get_media_type
from docwen_core.markdown_utils import format_md_file_link, sanitize_filename
from docwen_core.models.artifact import (
    ARTIFACT_KIND_AUXILIARY,
    ARTIFACT_KIND_IMAGE,
    ArtifactManifest,
)
from docwen_core.text.image_markdown import build_base64_image_data_uri, build_image_ocr_sidecar
from docwen_core.text.ocr import OcrOutcome, OcrStatus, format_ocr_best_effort_warning, run_ocr_outcome
from docwen_plugin_markup._common import new_artifact_id


@dataclass(frozen=True)
class MarkdownResource:
    source_key: str
    suggested_name: str
    media_type: str
    data: bytes


@dataclass(frozen=True)
class WrittenMarkdownResource:
    source_key: str
    suggested_name: str
    staging_path: str
    artifact_id: str
    markdown_link: str
    artifact: ArtifactManifest | None
    artifacts: tuple[ArtifactManifest, ...]


class MarkdownResourceWriter:
    """Write binary resources and return Markdown link mappings."""

    def write_all(
        self,
        context: object,
        resources: list[MarkdownResource],
        *,
        image_link_style: str | None = None,
        image_mode: str = "file",
        enable_ocr: bool = False,
        ocr_placement: str = "image_md",
        ocr_language: str = "auto",
        current_locale: str = "zh_CN",
        source_format: str = "markup",
        keep_resource_artifacts: bool = True,
        export_semantics: MarkdownExportSemantics | None = None,
        ocr_blockquote_title: str | None = None,
    ) -> dict[str, WrittenMarkdownResource]:
        written: dict[str, WrittenMarkdownResource] = {}
        used_names: set[str] = set()
        request_policy = resolve_markdown_request_policy(context)
        export_semantics = export_semantics or request_policy.export
        resolved_ocr_title = (
            request_policy.ocr_blockquote_title if ocr_blockquote_title is None else ocr_blockquote_title
        )
        resolved_image_style = image_link_style or export_semantics.image_link_style
        resolved_md_file_style = export_semantics.md_file_link_style
        resolved_image_mode = str(image_mode or "file").strip().lower()
        if resolved_image_mode not in {"file", "base64", "embed", "omit"}:
            resolved_image_mode = "file"
        resolved_ocr_placement = str(ocr_placement or "image_md").strip().lower()
        if resolved_ocr_placement not in {"image_md", "main_md"}:
            resolved_ocr_placement = "image_md"

        for resource in resources:
            source_key = normalize_resource_key(resource.source_key)
            suggested_name = _unique_name(_safe_basename(resource.suggested_name), used_names)
            suffix = PurePosixPath(suggested_name).suffix or ".bin"
            staging_path = context.workspace.create_artifact_path(ARTIFACT_KIND_AUXILIARY, suffix)  # pyright: ignore[reportAttributeAccessIssue]

            with open(staging_path, "wb") as handle:
                handle.write(resource.data)

            detected_format, resource_type = _classify_staged_resource(staging_path)
            resolved_media_type = get_media_type(detected_format)
            is_image = resource_type == "image"
            is_markdown = resource_type == "markdown"
            kind = ARTIFACT_KIND_IMAGE if is_image else ARTIFACT_KIND_AUXILIARY

            artifact: ArtifactManifest | None = None
            artifact_id = ""
            item_artifacts: list[ArtifactManifest] = []
            register_resource_artifact = keep_resource_artifacts
            if is_image and resolved_image_mode in {"base64", "omit"}:
                register_resource_artifact = False
            if register_resource_artifact:
                artifact_id = new_artifact_id()
                artifact = ArtifactManifest(
                    artifact_id=artifact_id,
                    kind=kind,
                    staging_path=staging_path,
                    suggested_name=suggested_name,
                    media_type=resolved_media_type,
                    metadata={
                        "source_key": source_key,
                        "detected_format": detected_format,
                        "declared_media_type": resource.media_type,
                    },
                    is_primary=False,
                )
                context.workspace.add_artifact(artifact)  # pyright: ignore[reportAttributeAccessIssue]
                item_artifacts.append(artifact)

            if is_image:
                # Images: alt text uses the full filename for accessibility.
                label = PurePosixPath(suggested_name).name or suggested_name
                image_link = _format_image_resource_link(
                    staging_path=staging_path,
                    suggested_name=suggested_name,
                    media_type=resolved_media_type,
                    label=label,
                    image_mode=resolved_image_mode,
                    image_link_style=resolved_image_style,
                    keep_images=keep_resource_artifacts,
                    export_semantics=export_semantics,
                )
                link, ocr_artifacts = _apply_image_ocr(
                    context,
                    image_link=image_link,
                    image_staging_path=staging_path,
                    suggested_name=suggested_name,
                    enable_ocr=enable_ocr,
                    ocr_placement=resolved_ocr_placement,
                    ocr_language=ocr_language,
                    current_locale=current_locale,
                    md_file_link_style=resolved_md_file_style,
                    image_format=detected_format,
                    source_format=source_format,
                    ocr_blockquote_title=resolved_ocr_title,
                )
                item_artifacts.extend(ocr_artifacts)
            else:
                if artifact is None:
                    link = ""
                elif is_markdown:
                    link = format_md_file_link(suggested_name, style=resolved_md_file_style)
                else:
                    # Other attachments use a normal Markdown link.
                    label = PurePosixPath(suggested_name).stem or suggested_name
                    link = f"[{label}]({suggested_name})"
            item = WrittenMarkdownResource(
                source_key=source_key,
                suggested_name=suggested_name,
                staging_path=staging_path,
                artifact_id=artifact_id,
                markdown_link=link,
                artifact=artifact,
                artifacts=tuple(item_artifacts),
            )
            written[source_key] = item
            basename_key = normalize_resource_key(PurePosixPath(source_key).name)
            written.setdefault(basename_key, item)

        return written


def _classify_staged_resource(staging_path: str) -> tuple[str, str]:
    """Return ``(format, route)`` using only staged resource content.

    The original filename and declared MIME remain provenance/output naming;
    neither can promote bytes into an image or Markdown execution path.
    Plain text shares the Markdown/text workflow, matching top-level TXT/MD
    behavior.  Corrupt images fall back to a generic attachment.
    """
    detected_format = detect_content_format(staging_path).format
    if get_category(detected_format) == CATEGORY_IMAGE:
        try:
            with Image.open(staging_path) as image:
                image.load()
        except Exception:
            return "unknown", "other"
        return detected_format, "image"
    if detected_format in {"markdown", "txt"}:
        return detected_format, "markdown"
    return detected_format, "other"


def _format_image_resource_link(
    *,
    staging_path: str,
    suggested_name: str,
    media_type: str,
    label: str,
    image_mode: str,
    image_link_style: str,
    keep_images: bool,
    export_semantics: MarkdownExportSemantics,
) -> str:
    if not keep_images:
        return ""
    if image_mode == "omit":
        return f"<!-- image omitted: {label} -->"
    if image_mode == "base64":
        target = build_base64_image_data_uri(
            image_path=staging_path,
            media_type=media_type,
            export_semantics=export_semantics,
        )
        return format_image_link(label, target, style=image_link_style)
    target = f"./{suggested_name}" if image_mode == "embed" else suggested_name
    return format_image_link(label, target, style=image_link_style)


def normalize_resource_key(value: str) -> str:
    """Normalize container-internal resource references for lookup."""
    split = urlsplit(value.replace("\\", "/"))
    path = unquote(split.path).strip()
    parts: list[str] = []
    for part in path.split("/"):
        if part in ("", ".", ".."):
            continue
        parts.append(part.lower())
    return "/".join(parts)


def _safe_basename(value: str) -> str:
    name = PurePosixPath(value.replace("\\", "/")).name.strip()
    if not name:
        return "resource.bin"
    safe = "".join(ch if ch.isalnum() or ch in ".-_" else "-" for ch in name)
    return safe or "resource.bin"


def _unique_name(name: str, used_names: set[str]) -> str:
    path = PurePosixPath(name)
    stem = path.stem or "resource"
    suffix = path.suffix
    candidate = f"{stem}{suffix}"
    counter = 2
    while candidate.lower() in used_names:
        candidate = f"{stem}-{counter}{suffix}"
        counter += 1
    used_names.add(candidate.lower())
    return candidate


def _apply_image_ocr(
    context: object,
    *,
    image_link: str,
    image_staging_path: str,
    suggested_name: str,
    enable_ocr: bool,
    ocr_placement: str,
    ocr_language: str,
    current_locale: str,
    md_file_link_style: str,
    image_format: str,
    source_format: str,
    ocr_blockquote_title: str,
) -> tuple[str, list[ArtifactManifest]]:
    if not enable_ocr:
        return image_link, []

    try:
        outcome = run_ocr_outcome(
            image_staging_path,
            source_format=image_format,
            ocr_language=ocr_language,
            current_locale=current_locale,
        )
    except Exception as exc:
        # Keep the compatibility boundary best-effort even if a custom OCR
        # implementation violates the core typed-outcome contract.
        outcome = OcrOutcome(OcrStatus.RECOGNITION_FAILED, message=str(exc))

    _report_best_effort_ocr_warning(
        context,
        outcome=outcome,
        suggested_name=suggested_name,
        source_format=source_format,
    )

    ocr_text = outcome.recognized_text
    if not ocr_text:
        return image_link, []

    if ocr_placement == "image_md":
        return _write_image_ocr_sidecar(
            context,
            image_link=image_link,
            suggested_name=suggested_name,
            ocr_text=ocr_text,
            md_file_link_style=md_file_link_style,
            source_format=source_format,
            ocr_blockquote_title=ocr_blockquote_title,
        )

    inline = _format_inline_ocr_block(ocr_text, ocr_blockquote_title)
    if image_link:
        return f"{image_link}\n\n{inline}", []
    return inline, []


def _report_best_effort_ocr_warning(
    context: object,
    *,
    outcome: OcrOutcome,
    suggested_name: str,
    source_format: str,
) -> None:
    """Report one non-fatal OCR quality or fallback warning through the request sink."""
    message = format_ocr_best_effort_warning(
        outcome.status,
        context=f"{source_format} image {suggested_name}",
    )
    if message is None:
        return

    # ``MarkdownResourceWriter`` is also a public lower-level helper and some
    # extension callers provide a workspace-only context.  The production
    # ConverterContext always has this sink; keep workspace-only callers
    # compatible instead of introducing process-global diagnostic state.
    progress = getattr(context, "progress", None)
    report_diagnostic = getattr(progress, "report_diagnostic", None)
    if callable(report_diagnostic):
        report_diagnostic(
            "warning",
            message,
            code="OCR-BEST-EFFORT",
            location=suggested_name,
        )


def _write_image_ocr_sidecar(
    context: object,
    *,
    image_link: str,
    suggested_name: str,
    ocr_text: str,
    md_file_link_style: str,
    source_format: str,
    ocr_blockquote_title: str,
) -> tuple[str, list[ArtifactManifest]]:
    sidecar_stem = f"{PurePosixPath(suggested_name).stem or 'image'}_ocr"
    sidecar_text, replacement_link = build_image_ocr_sidecar(
        sidecar_stem=sidecar_stem,
        source_format=source_format,
        image_markdown=image_link,
        ocr_text=ocr_text,
        md_link_style=md_file_link_style,
        ocr_blockquote_title=ocr_blockquote_title,
    )
    sidecar_path = context.workspace.create_artifact_path(ARTIFACT_KIND_AUXILIARY, ".md")  # pyright: ignore[reportAttributeAccessIssue]
    Path(sidecar_path).write_text(sidecar_text, encoding="utf-8")
    sidecar_artifact = ArtifactManifest(
        artifact_id=new_artifact_id(),
        kind=ARTIFACT_KIND_AUXILIARY,
        staging_path=sidecar_path,
        suggested_name=sanitize_filename(f"{sidecar_stem}.md"),
        media_type="text/markdown",
        metadata={"source_format": source_format, "ocr": True, "image": suggested_name},
        is_primary=False,
    )
    context.workspace.add_artifact(sidecar_artifact)  # pyright: ignore[reportAttributeAccessIssue]
    return replacement_link, [sidecar_artifact]


def _format_inline_ocr_block(ocr_text: str, ocr_title: str) -> str:
    lines: list[str] = []
    if ocr_title:
        lines.append(f"> **{ocr_title}**")
        lines.append(">")
    for line in ocr_text.splitlines():
        stripped = line.strip()
        lines.append(f"> {stripped}" if stripped else ">")
    return "\n".join(lines)
