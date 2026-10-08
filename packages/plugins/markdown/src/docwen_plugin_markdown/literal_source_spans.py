"""Select excluded source spans from Core's shared lexical ownership."""

from __future__ import annotations

import re
from collections.abc import Sequence

from docwen_core.links import markdown_source_owners


def literal_source_spans(
    source: str,
    *,
    metadata_patterns: Sequence[re.Pattern[str]] = (),
    semantic_url_suffix: bool = False,
    exclude_wikilinks: bool = False,
) -> list[tuple[int, int]]:
    """Let each consumer decide whether Wiki ownership hides its record."""
    return [
        (owner.start, owner.end)
        for owner in markdown_source_owners(
            source, metadata_patterns=metadata_patterns, semantic_url_suffix=semantic_url_suffix
        )
        if owner.kind != "wikilink" or exclude_wikilinks
    ]
