"""Source-ordered ownership on an authoritative Markdown block projection."""

from __future__ import annotations

import re
from collections.abc import Sequence

from docwen_core.links._block_context import markdown_inline_regions
from docwen_core.links._markdown_inline import MarkdownInlineSourceOwner, _is_backslash_escaped
from docwen_core.links._non_embed import markdown_inline_source_owners, split_markdown_block_segments

_COMMENT_RE = re.compile(r"<!--|%%")


def markdown_source_owners(
    source: str,
    *,
    metadata_patterns: Sequence[re.Pattern[str]] = (),
    semantic_url_suffix: bool = False,
    full_document: bool = False,
    structural_tables: bool = False,
    extended_headings: bool = False,
) -> list[MarkdownInlineSourceOwner]:
    """Resolve literals, metadata and comments in source order.

    Callers first mask authoritative blocks without changing source length.
    Earlier comments own their original closer; later-looking invalidated
    atoms cannot own the suffix. Wiki ownership does not hide its semantic
    record. Consumers decide which owner kinds to exclude from their scan.
    """

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
        if full_document:
            parts: list[str] = []
            offset = 0
            for segment, protected in split_markdown_block_segments(remaining, extended_headings=extended_headings):
                end = offset + len(segment)
                if protected:
                    atoms.append((base + offset, 0, base + end, base + end, "block"))
                    parts.append("".join(char if char in "\r\n" else " " for char in segment))
                else:
                    parts.append(segment)
                offset = end
            projection = "".join(parts)
        regions = (
            markdown_inline_regions(
                projection, structural_tables=structural_tables, extended_headings=extended_headings
            )
            if full_document
            else [(0, len(projection))]
        )
        for region_start, region_end in regions:
            for owner in markdown_inline_source_owners(
                projection[region_start:region_end], include_link_syntax=full_document
            ):
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
