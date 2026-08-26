"""Canonical heading/body merge punctuation contract.

The option schema, Markdown plugin, and GUI all consume this lightweight Core
module. Keeping the ordered editable default here prevents cross-layer drift
without making Core depend on a conversion plugin.
"""

from __future__ import annotations

DEFAULT_HEADING_MERGE_PUNCTUATION = "。：！？.:!?"
HEADING_MERGE_PUNCTUATION_SET = frozenset(DEFAULT_HEADING_MERGE_PUNCTUATION)
HALFWIDTH_HEADING_MERGE_PUNCTUATION = frozenset(".,;:!?")


def normalize_heading_merge_punctuation(value: object | None) -> frozenset[str]:
    """Return the request/config value as a de-duplicated character set.

    ``None`` means the canonical strong-ending default. An explicit empty
    string remains empty, and whitespace is ignored.
    """

    source = DEFAULT_HEADING_MERGE_PUNCTUATION if value is None else str(value)
    return frozenset(character for character in source if not character.isspace())
