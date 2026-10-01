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
    after_text: str
    previous_text: str
    next_text: str
    image: ClipboardImageRef


def _paragraph_text(paragraph: ClipboardParagraph) -> str:
    return "".join(
        inline.value if isinstance(inline, ClipboardText) else "\n"
        for inline in paragraph.inlines
        if isinstance(inline, (ClipboardText, ClipboardHardBreak))
    )


def _context_text(value: str) -> str:
    # HTML source wrapping and Office's NBSP padding are comparison-only.
    # Authored body characters are retained from the captured plain format.
    return " ".join(value.split())


def _paragraph_candidates(
    paragraph: ClipboardParagraph, path: str, previous: str, following: str
) -> list[_ImageCandidate]:
    output: list[_ImageCandidate] = []
    segments: list[list[str]] = [[]]
    images: list[ClipboardImageRef] = []
    for inline in paragraph.inlines:
        if isinstance(inline, ClipboardText):
            segments[-1].append(inline.value)
        elif isinstance(inline, ClipboardHardBreak):
            segments[-1].append("\n")
        elif isinstance(inline, ClipboardImageRef):
            images.append(inline)
            segments.append([])
    for index, image in enumerate(images):
        output.append(
            _ImageCandidate(path, "".join(segments[index]), "".join(segments[index + 1]), previous, following, image)
        )
    return output


def _collect_candidates(document: ClipboardDocument) -> tuple[_ImageCandidate, ...]:
    output: list[_ImageCandidate] = []

    def walk_blocks(blocks: tuple[ClipboardBlock, ...], path: str) -> None:
        texts = [_paragraph_text(block) if isinstance(block, ClipboardParagraph) else "" for block in blocks]
        previous_texts: list[str] = []
        previous = ""
        for text in texts:
            previous_texts.append(previous)
            if _context_text(text):
                previous = text
        next_texts = [""] * len(blocks)
        following = ""
        for index in range(len(blocks) - 1, -1, -1):
            next_texts[index] = following
            if _context_text(texts[index]):
                following = texts[index]
        table_index = 0
        for block_index, block in enumerate(blocks):
            if isinstance(block, ClipboardParagraph):
                output.extend(_paragraph_candidates(block, path, previous_texts[block_index], next_texts[block_index]))
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
        if _context_text(candidate.previous_text) != _context_text(occurrence.previous_text):
            continue
        if _context_text(candidate.next_text) != _context_text(occurrence.next_text):
            continue
        if _context_text(candidate.before_text) != _context_text(occurrence.before_text):
            continue
        if _context_text(candidate.after_text) != _context_text(occurrence.after_text):
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
    previous_candidate = -1
    for occurrence in projection.occurrences:
        candidate_index = _match_occurrence(occurrence, candidates, used)
        if candidate_index <= previous_candidate:
            raise ClipboardOfficeProviderError(
                "clipboard.provider_binding_invalid",
                "Clipboard provider image order conflicts with document structure.",
            )
        previous_candidate = candidate_index
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
