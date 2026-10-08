"""Source-ordered ownership on an authoritative Markdown block projection."""

from __future__ import annotations

import re
from collections.abc import Sequence

from docwen_core.links._markdown_inline import MarkdownInlineSourceOwner, _is_backslash_escaped
from docwen_core.links._non_embed import markdown_inline_source_owners

_COMMENT_RE = re.compile(r"<!--|%%")


def markdown_source_owners(
    source: str,
    *,
    metadata_patterns: Sequence[re.Pattern[str]] = (),
    semantic_url_suffix: bool = False,
) -> list[MarkdownInlineSourceOwner]:
    """Resolve literals, metadata and comments in source order.

    Callers first mask authoritative blocks without changing source length.
    Earlier comments own their original closer; later-looking invalidated
    atoms cannot own the suffix. Wiki ownership does not hide its semantic
    record. Consumers decide which owner kinds to exclude from their scan.
    """

    def atoms_after(start: int) -> list[tuple[int, int, int, int, str]]:
        remaining = source[start:]
        atoms: list[tuple[int, int, int, int, str]] = []
        for owner in markdown_inline_source_owners(remaining):
            consume_end = owner.end
            if owner.kind == "url" and semantic_url_suffix:
                suffix = remaining.find("@[[", owner.start, owner.end)
                if suffix >= 0:
                    consume_end = suffix
            atoms.append((start + owner.start, 0, start + owner.end, start + consume_end, owner.kind))
        for priority, pattern in enumerate(metadata_patterns, start=1):
            atoms.extend(
                (start + match.start(), priority, start + match.end(), start + match.end(), "literal")
                for match in pattern.finditer(remaining)
            )
        return sorted(atoms)

    owners: list[MarkdownInlineSourceOwner] = []
    atoms = atoms_after(0)
    atom_index = 0
    cursor = 0
    while cursor < len(source):
        match = _COMMENT_RE.search(source, cursor)
        while match is not None and match.group(0) == "<!--" and _is_backslash_escaped(source, match.start()):
            match = _COMMENT_RE.search(source, match.end())
        atom = atoms[atom_index] if atom_index < len(atoms) else None
        if atom is not None and (match is None or atom[0] < match.start()):
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
