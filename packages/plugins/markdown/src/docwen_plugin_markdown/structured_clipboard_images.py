"""Linked-image helpers for structured clipboard converters."""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from docwen_core.export_semantics import (
    VALID_LINK_STYLES,
    format_image_link,
    resolve_markdown_request_policy,
)
from docwen_core.models.artifact import ARTIFACT_KIND_IMAGE, ArtifactManifest
from docwen_core.models.clipboard_document import (
    ClipboardBlock,
    ClipboardDocument,
    ClipboardImageRef,
    ClipboardParagraph,
    ClipboardResource,
    ClipboardTable,
)
from docwen_core.models.file_ref import MANAGED_RESOURCE_ID_METADATA_KEY
from docwen_core.text.image_markdown import build_base64_image_data_uri


@dataclass(frozen=True, slots=True)
class BoundClipboardImage:
    resource: ClipboardResource
    path: Path
    suggested_name: str


@dataclass(frozen=True, slots=True)
class ClipboardImageOccurrenceInfo:
    ordinal: int
    image: ClipboardImageRef
    table_id: str
    cell_anchor: str


@dataclass(frozen=True, slots=True)
class MarkdownImagePlan:
    artifacts: tuple[ArtifactManifest, ...]
    targets: dict[str, str]
    style: str
    mode: str

    def reference(self, image: ClipboardImageRef, resources: dict[str, BoundClipboardImage]) -> str:
        if image.resource_id is None:
            return f"[Image unavailable: {image.alt}]" if image.alt else "[Image unavailable]"
        target = self.targets.get(image.resource_id)
        if target is None:
            return f"[Image not rendered: {image.alt}]" if image.alt else "[Image not rendered]"
        if self.mode == "omit":
            label = image.alt or resources[image.resource_id].suggested_name
            return f"<!-- image omitted: {label} -->"
        label = image.alt or Path(resources[image.resource_id].suggested_name).stem
        return format_image_link(label, target, style=self.style)


def resolve_bound_images(context: Any, document: ClipboardDocument) -> dict[str, BoundClipboardImage]:
    """Resolve exact linked resources already materialized by Runtime."""

    refs = context.workspace.input_resources("linked_resource")
    by_id: dict[str, Any] = {}
    for ref in refs:
        resource_id = ref.metadata.get(MANAGED_RESOURCE_ID_METADATA_KEY)
        if not isinstance(resource_id, str) or not resource_id or resource_id in by_id:
            raise ValueError("structured clipboard linked resource identity is invalid")
        by_id[resource_id] = ref

    output: dict[str, BoundClipboardImage] = {}
    for resource in document.resources:
        if resource.media_type != "image/png":
            continue
        ref = by_id.get(resource.resource_id)
        if ref is None:
            raise ValueError("structured clipboard linked image is missing")
        suggested_name = f"clipboard-image-{resource.sha256[:16]}.png"
        output[resource.resource_id] = BoundClipboardImage(
            resource=resource,
            path=Path(ref.path),
            suggested_name=suggested_name,
        )
    return output


def prepare_markdown_images(
    context: Any,
    resources: dict[str, BoundClipboardImage],
) -> MarkdownImagePlan:
    """Apply existing Markdown image mode/link policy to bound clipboard images."""

    policy = resolve_markdown_request_policy(context)
    options = context.request.options
    mode = str(options.get("image_mode") or policy.export.image_extraction_mode or "file").strip().lower()
    if mode not in {"file", "base64", "embed", "omit"}:
        mode = "file"
    requested_style = str(options.get("image_link_style") or "").strip().lower()
    style = requested_style if requested_style in VALID_LINK_STYLES else policy.export.image_link_style

    artifacts: list[ArtifactManifest] = []
    targets: dict[str, str] = {}
    for resource_id, bound in sorted(resources.items()):
        context.cancellation.check()
        if mode == "omit":
            targets[resource_id] = ""
            continue
        if mode == "base64":
            targets[resource_id] = build_base64_image_data_uri(
                image_path=str(bound.path),
                media_type="image/png",
                export_semantics=policy.export,
            )
            continue

        staging = context.workspace.create_artifact_path(ARTIFACT_KIND_IMAGE, ".png")
        shutil.copyfile(bound.path, staging)
        artifact = ArtifactManifest(
            artifact_id=f"clipboard-image-{bound.resource.sha256[:20]}",
            kind=ARTIFACT_KIND_IMAGE,
            staging_path=staging,
            suggested_name=bound.suggested_name,
            media_type="image/png",
            metadata={
                "clipboard_resource_id": resource_id,
                "pixel_width": bound.resource.pixel_width,
                "pixel_height": bound.resource.pixel_height,
            },
            is_primary=False,
        )
        artifacts.append(artifact)
        targets[resource_id] = (
            f"./{bound.suggested_name}" if mode == "embed" else bound.suggested_name
        )

    return MarkdownImagePlan(tuple(artifacts), targets, style, mode)


def copy_bound_image_artifacts(
    context: Any,
    resources: dict[str, BoundClipboardImage],
) -> tuple[ArtifactManifest, ...]:
    """Copy each unique bound PNG once as a result resource."""

    artifacts: list[ArtifactManifest] = []
    for resource_id, bound in sorted(resources.items()):
        context.cancellation.check()
        staging = context.workspace.create_artifact_path(ARTIFACT_KIND_IMAGE, ".png")
        shutil.copyfile(bound.path, staging)
        artifacts.append(
            ArtifactManifest(
                artifact_id=f"clipboard-image-{bound.resource.sha256[:20]}",
                kind=ARTIFACT_KIND_IMAGE,
                staging_path=staging,
                suggested_name=bound.suggested_name,
                media_type="image/png",
                metadata={
                    "clipboard_resource_id": resource_id,
                    "pixel_width": bound.resource.pixel_width,
                    "pixel_height": bound.resource.pixel_height,
                },
                is_primary=False,
            )
        )
    return tuple(artifacts)


def collect_image_occurrences(
    document: ClipboardDocument,
    table_id_for: Callable[[ClipboardTable], str],
) -> tuple[ClipboardImageOccurrenceInfo, ...]:
    """Return every occurrence in document order without deduplicating references."""

    output: list[ClipboardImageOccurrenceInfo] = []

    def walk(blocks: tuple[ClipboardBlock, ...], *, table_id: str = "", cell_anchor: str = "") -> None:
        for block in blocks:
            if isinstance(block, ClipboardParagraph):
                for inline in block.inlines:
                    if isinstance(inline, ClipboardImageRef):
                        output.append(
                            ClipboardImageOccurrenceInfo(
                                ordinal=len(output) + 1,
                                image=inline,
                                table_id=table_id,
                                cell_anchor=cell_anchor,
                            )
                        )
                continue
            current_table_id = table_id_for(block)
            for cell in sorted(block.cells, key=lambda item: (item.row, item.column)):
                walk(
                    cell.blocks,
                    table_id=current_table_id,
                    cell_anchor=f"R{cell.row + 1}C{cell.column + 1}",
                )

    walk(document.blocks)
    return tuple(output)


__all__ = [
    "BoundClipboardImage",
    "ClipboardImageOccurrenceInfo",
    "MarkdownImagePlan",
    "collect_image_occurrences",
    "copy_bound_image_artifacts",
    "prepare_markdown_images",
    "resolve_bound_images",
]
