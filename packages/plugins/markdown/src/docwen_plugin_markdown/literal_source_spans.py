"""Source-ordered ownership of inline literals on a protected block projection."""

from __future__ import annotations

import re
from collections.abc import Sequence

from docwen_core.links import split_markdown_inline_segments

_COMMENT_OR_URL_RE = re.compile(r"<!--|%%|https?://[^\s<]+")


def literal_source_spans(
    source: str,
    *,
    metadata_patterns: Sequence[re.Pattern[str]],
    semantic_url_suffix: bool = False,
) -> list[tuple[int, int]]:
    """Let the first opener own its closer, including inside later-looking atoms.

    Callers mask authoritative block ranges before this scan. Inline atoms,
    metadata, comments and URLs then compete in source order. An atom opened
    inside a comment has no ownership outside it, even if its apparent closer
    follows the comment closer. Rescan that suffix when such a candidate had
    hidden later atoms from the public inline splitter.
    """

    def atoms_after(start: int) -> list[tuple[int, int, int]]:
        remaining = source[start:]
        atoms: list[tuple[int, int, int]] = []
        offset = start
        for segment, protected in split_markdown_inline_segments(remaining, protect_bare_urls=False):
            if protected:
                atoms.append((offset, 0, offset + len(segment)))
            offset += len(segment)
        for priority, pattern in enumerate(metadata_patterns, start=1):
            atoms.extend(
                (start + match.start(), priority, start + match.end()) for match in pattern.finditer(remaining)
            )
        return sorted(atoms)

    spans: list[tuple[int, int]] = []
    atoms = atoms_after(0)
    atom_index = 0
    cursor = 0
    while cursor < len(source):
        match = _COMMENT_OR_URL_RE.search(source, cursor)
        atom = atoms[atom_index] if atom_index < len(atoms) else None
        if atom is not None and (match is None or atom[0] < match.start()):
            start, _priority, end = atom
        elif match is not None:
            start = match.start()
            token = match.group(0)
            if token in {"<!--", "%%"}:
                closer = "-->" if token == "<!--" else "%%"
                closing_start = source.find(closer, match.end())
                end = len(source) if closing_start < 0 else closing_start + len(closer)
            else:
                suffix = token.find("@[[") if semantic_url_suffix else -1
                end = match.end() if suffix < 0 else start + suffix
        else:
            break
        spans.append((start, end))
        cursor = end
        crosses_owner = False
        while atom_index < len(atoms) and atoms[atom_index][0] < end:
            crosses_owner |= atoms[atom_index][2] > end
            atom_index += 1
        if crosses_owner:
            atoms = atoms_after(end)
            atom_index = 0
    return spans
