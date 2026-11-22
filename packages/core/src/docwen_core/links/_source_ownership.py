"""Source-ordered ownership on an authoritative Markdown block projection."""

from __future__ import annotations

import re
from collections.abc import Sequence
from itertools import pairwise

from docwen_core.links._block_context import markdown_inline_regions, markdown_source_lines
from docwen_core.links._markdown_inline import (
    _ACTIVE_REFERENCE_LABELS,
    MarkdownInlineSourceOwner,
    _is_backslash_escaped,
)
from docwen_core.links._markdown_references import (
    _ACTIVE_REFERENCE_DEFINITIONS,
    _ACTIVE_REFERENCE_IMAGE_SOURCE,
    MarkdownReferenceDefinition,
    reference_image_source,
)
from docwen_core.links._non_embed import _split_fenced_code_blocks, markdown_inline_source_owners
from docwen_core.yaml_tools import yaml_front_matter_end

_COMMENT_RE = re.compile(r"<!--|%%")


def markdown_source_owners(
    source: str,
    *,
    metadata_patterns: Sequence[re.Pattern[str]] = (),
    semantic_url_suffix: bool = False,
    full_document: bool = False,
    include_front_matter: bool = True,
    structural_tables: bool = False,
    extended_headings: bool = False,
) -> list[MarkdownInlineSourceOwner]:
    """Resolve literals, metadata and comments in source order.

    Callers first mask authoritative blocks without changing source length.
    Earlier comments own their original closer; later-looking invalidated
    atoms cannot own the suffix. Wiki ownership does not hide its semantic
    record. Consumers decide which owner kinds to exclude from their scan.
    """
    body_start = None
    if full_document:
        body_start = yaml_front_matter_end(source) if include_front_matter else 0
    return _collect_source_owners(
        source,
        body_start=body_start,
        metadata_patterns=metadata_patterns,
        semantic_url_suffix=semantic_url_suffix,
        structural_tables=structural_tables,
        extended_headings=extended_headings,
    )


def _collect_source_owners(
    source: str,
    *,
    body_start: int | None,
    metadata_patterns: Sequence[re.Pattern[str]] = (),
    semantic_url_suffix: bool = False,
    structural_tables: bool = False,
    extended_headings: bool = False,
) -> list[MarkdownInlineSourceOwner]:
    """Scan a document with its known body start, or an inline fragment (None)."""
    full_document = body_start is not None
    definitions: dict[str, MarkdownReferenceDefinition] | None = {} if full_document else None

    def atoms_after(start: int) -> list[tuple[int, int, int, int, str]]:
        base = start
        remaining = source[start:]
        if full_document and start and source[start - 1] not in "\r\n":
            base = max(source.rfind("\n", 0, start), source.rfind("\r", 0, start)) + 1
            # A resumed scan is still in the original physical line. Neutral
            # punctuation prevents a comment closer from manufacturing a new
            # fence, heading or indented-code start in its same-line suffix.
            remaining = "." * (start - base) + remaining
        atoms: list[tuple[int, int, int, int, str]] = []
        projection = remaining
        if body_start is not None:
            parts: list[str] = []
            offset = 0
            source_prefixes: dict[int, int] = {}
            segments = _split_fenced_code_blocks(
                remaining,
                extended_headings=extended_headings,
                source_prefixes=source_prefixes,
                # The first block scan owns all definitions, including forward references.
                # A resumed comment scan can change its suffix's block context only.
                references=definitions if start == 0 else None,
                front_matter_end=body_start if start == 0 else 0,
            )
            for segment, protected in segments:
                end = offset + len(segment)
                if protected:
                    atoms.append((base + offset, 0, base + end, base + end, "block"))
                    masked = [char if char in "\r\n" else " " for char in segment]
                    line_offset = offset
                    for line in markdown_source_lines(segment):
                        prefix = source_prefixes.get(line_offset, 0)
                        segment_line_start = line_offset - offset
                        masked[segment_line_start : segment_line_start + prefix] = line[:prefix]
                        line_offset += len(line)
                    parts.append("".join(masked))
                else:
                    parts.append(segment)
                offset = end
            projection = "".join(parts)
        image_source = (
            reference_image_source(projection, structural_tables=structural_tables, extended_headings=extended_headings)
            if definitions or _ACTIVE_REFERENCE_DEFINITIONS.get()
            else None
        )
        regions = (
            (
                pairwise(image_source.boundaries)
                if image_source
                else markdown_inline_regions(
                    projection, structural_tables=structural_tables, extended_headings=extended_headings
                )
            )
            if full_document
            else [(0, len(projection))]
        )
        for region_start, region_end in regions:
            reference_token = _ACTIVE_REFERENCE_LABELS.set(
                frozenset(definitions) if definitions is not None else _ACTIVE_REFERENCE_LABELS.get()
            )
            definitions_token = _ACTIVE_REFERENCE_DEFINITIONS.set(
                definitions if definitions is not None else _ACTIVE_REFERENCE_DEFINITIONS.get()
            )
            region = image_source.region(region_start, region_end) if image_source else None
            source_token = _ACTIVE_REFERENCE_IMAGE_SOURCE.set(region)
            try:
                inline_owners = markdown_inline_source_owners(
                    region.source if region else projection[region_start:region_end], include_link_syntax=full_document
                )
            finally:
                _ACTIVE_REFERENCE_IMAGE_SOURCE.reset(source_token)
                _ACTIVE_REFERENCE_DEFINITIONS.reset(definitions_token)
                _ACTIVE_REFERENCE_LABELS.reset(reference_token)
            for owner in inline_owners:
                owner_start = region_start + owner.start
                owner_end = region_start + owner.end
                consume_end = owner_end
                if owner.kind == "url" and semantic_url_suffix:
                    suffix = remaining.find("@[[", owner_start, owner_end)
                    if suffix >= 0:
                        consume_end = suffix
                atoms.append((base + owner_start, 0, base + owner_end, base + consume_end, owner.kind))
        for priority, pattern in enumerate(metadata_patterns, start=1):
            atoms.extend(
                (base + match.start(), priority, base + match.end(), base + match.end(), "literal")
                for match in pattern.finditer(remaining)
            )
        return sorted(atom for atom in atoms if atom[0] >= start)

    owners: list[MarkdownInlineSourceOwner] = []
    atoms = atoms_after(0)
    atom_index = 0
    cursor = 0
    while cursor < len(source):
        match = _COMMENT_RE.search(source, cursor)
        while match is not None and match.group(0) == "<!--" and _is_backslash_escaped(source, match.start()):
            match = _COMMENT_RE.search(source, match.end())
        atom = atoms[atom_index] if atom_index < len(atoms) else None
        if atom is not None and (
            match is None or atom[0] < match.start() or (atom[0] == match.start() and atom[4] == "block")
        ):
            start, _priority, _parsed_end, end, kind = atom
        elif match is not None:
            start = match.start()
            closer = "-->" if match.group(0) == "<!--" else "%%"
            closing_start = source.find(closer, match.end())
            end = len(source) if closing_start < 0 else closing_start + len(closer)
            kind = "comment"
        else:
            break
        owners.append(MarkdownInlineSourceOwner(start, end, kind))
        cursor = end
        crosses_owner = False
        while atom_index < len(atoms) and atoms[atom_index][0] < end:
            crosses_owner |= atoms[atom_index][2] > end
            atom_index += 1
        if crosses_owner:
            atoms = atoms_after(end)
            atom_index = 0
    return owners
