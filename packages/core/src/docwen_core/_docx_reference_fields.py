"""One closed REF instruction grammar for native field-format retention."""

from __future__ import annotations

import re

_REF_INSTRUCTION_RE = re.compile(
    r"[ \t]*(?i:REF)[ \t]+([^ \t\r\n]+)([ \t]+(?i:\\n))?[ \t]+(?i:\\h)"
    r"(?:[ \t]+\\\*[ \t]+(?i:MERGEFORMAT))?[ \t]*"
)


def reference_instruction(bookmark_name: str, *, heading_number_only: bool = False) -> str:
    """Retain the cached result's text formatting when Word updates the REF."""

    number_switch = " \\n" if heading_number_only else ""
    return f" REF {bookmark_name}{number_switch} \\h \\* MERGEFORMAT "


def parse_reference_instruction(instruction: str) -> tuple[str, bool] | None:
    """Accept legacy or format-retaining REF syntax, with no other switches.

    Word may normalize keyword case or horizontal whitespace on save. Neither
    changes field semantics. Bookmark spelling and heading-number selection
    remain available for exact comparison with authenticated source metadata.
    """

    match = _REF_INSTRUCTION_RE.fullmatch(instruction)
    return None if match is None else (match.group(1), match.group(2) is not None)
