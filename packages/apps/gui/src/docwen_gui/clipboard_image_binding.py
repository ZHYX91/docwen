"""Bind provider-proven image bytes to structured clipboard occurrences."""

from __future__ import annotations

from dataclasses import dataclass, replace

from docwen_core.models.clipboard_document import (
    ClipboardBlock,
    ClipboardDocument,
    ClipboardHardBreak,
    ClipboardImageRef,
    ClipboardParagraph,
    ClipboardText,
    clipboard_document_to_bytes,
)
from docwen_gui.clipboard_office_provider import (
    ClipboardOfficeProviderError,
    ProviderImageOccurrence,
    ProviderImageProjection,
)


@dataclass(frozen=True, slots=True)
class _ImageCandidate:
    path: str
    before_text: str
    image: ClipboardImageRef


def _paragraph_candidates(paragraph: ClipboardParagraph, path: str) -> list[_ImageCandidate]:
    output: list[_ImageCandidate] = []
    before: list[str] = []
    for inline in paragraph.inlines:
        if isinstance(inline, ClipboardText):
            before.append(inline.value)
        elif isinstance(inline, ClipboardHardBreak):
            before.append("\n")
        elif isinstance(inline, ClipboardImageRef):
            output.append(_ImageCandidate(path, "".join(before), inline))
    return output


def _collect_candidates(document: ClipboardDocument) -> tuple[_ImageCandidate, ...]:
    output: list[_ImageCandidate] = []

    def walk_blocks(blocks: tuple[ClipboardBlock, ...], path: str) -> None:
        table_index = 0
        for block in blocks:
            if isinstance(block, ClipboardParagraph):
                output.extend(_paragraph_candidates(block, path))
                continue
            table_path = f"{path}/table:{table_index}"
            table_index += 1
            for cell in block.cells:
                walk_blocks(cell.blocks, f"{table_path}/cell:{cell.row}:{cell.column}")

    walk_blocks(document.blocks, "body")
    return tuple(output)


def _match_occurrence(
    occurrence: ProviderImageOccurrence,
    candidates: tuple[_ImageCandidate, ...],
    used: set[int],
) -> int:
    matches: list[int] = []
    for index, candidate in enumerate(candidates):
        if index in used or candidate.path != occurrence.container_path:
            continue
        if occurrence.previous_text and not candidate.before_text.endswith(occurrence.previous_text):
            continue
        matches.append(index)
    if len(matches) != 1:
        raise ClipboardOfficeProviderError(
            "clipboard.provider_binding_invalid",
            "Clipboard provider image occurrence could not be bound uniquely to document structure.",
        )
    return matches[0]


def bind_provider_images(
    document: ClipboardDocument,
    projection: ProviderImageProjection,
    existing_resource_bytes: tuple[tuple[str, str, str, bytes], ...] = (),
) -> tuple[ClipboardDocument, tuple[tuple[str, str, str, bytes], ...]]:
    """Bind provider occurrences by structural path/context, never by global ordinal."""

    candidates = _collect_candidates(document)
    identifier_to_resource = dict(projection.identifier_to_resource_id)
    replacements: dict[int, ClipboardImageRef] = {}
    used: set[int] = set()
    for occurrence in projection.occurrences:
        candidate_index = _match_occurrence(occurrence, candidates, used)
        used.add(candidate_index)
        candidate = candidates[candidate_index]
        resource_id = identifier_to_resource.get(occurrence.identifier)
        if resource_id is None:
            raise ClipboardOfficeProviderError(
                "clipboard.provider_binding_invalid",
                "Clipboard provider image occurrence has no frozen resource.",
            )
        replacements[id(candidate.image)] = ClipboardImageRef(
            resource_id=resource_id,
            alt=candidate.image.alt,
            missing_reason="",
            extent_cx_emu=occurrence.extent_cx_emu,
            extent_cy_emu=occurrence.extent_cy_emu,
        )

    def replace_blocks(blocks: tuple[ClipboardBlock, ...]) -> tuple[ClipboardBlock, ...]:
        output: list[ClipboardBlock] = []
        for block in blocks:
            if isinstance(block, ClipboardParagraph):
                inlines = tuple(
                    replacements.get(id(inline), inline) if isinstance(inline, ClipboardImageRef) else inline
                    for inline in block.inlines
                )
                output.append(replace(block, inlines=inlines))
                continue
            cells = tuple(replace(cell, blocks=replace_blocks(cell.blocks)) for cell in block.cells)
            output.append(replace(block, cells=cells))
        return tuple(output)

    resource_map = {resource.resource_id: resource for resource in document.resources}
    for resource in projection.resources:
        existing = resource_map.get(resource.resource_id)
        if existing is not None and existing != resource:
            raise ClipboardOfficeProviderError(
                "clipboard.provider_binding_invalid",
                "Clipboard provider resource conflicts with an existing resource.",
            )
        resource_map[resource.resource_id] = resource

    byte_map = {item[0]: item for item in existing_resource_bytes}
    for item in projection.resource_bytes:
        existing = byte_map.get(item[0])
        if existing is not None and existing != item:
            raise ClipboardOfficeProviderError(
                "clipboard.provider_binding_invalid",
                "Clipboard provider bytes conflict with an existing resource.",
            )
        byte_map[item[0]] = item

    merged = ClipboardDocument(
        blocks=replace_blocks(document.blocks),
        resources=tuple(resource_map[key] for key in sorted(resource_map)),
    )
    clipboard_document_to_bytes(merged)
    return merged, tuple(byte_map[key] for key in sorted(byte_map))


__all__ = ["bind_provider_images"]
