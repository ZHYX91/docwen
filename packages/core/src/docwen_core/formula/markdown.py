"""Parse LaTeX formulas from Markdown text."""

from __future__ import annotations

import re
from typing import TypedDict


class FormulaSpan(TypedDict):
    """A parsed LaTeX formula span from Markdown text."""

    latex: str
    is_inline: bool
    start: int
    end: int


def parse_latex_from_markdown(md_text: str) -> list[FormulaSpan]:
    """Extract LaTeX formulas from Markdown text.

    Handles:
    - Block formulas: $$...$$ (with \\begin...\\end content)
    - Inline formulas: $...$
    - Does NOT confuse $$ with $
    - Multiple formulas in same text

    Args:
        md_text: Raw Markdown text.

    Returns:
        List of formula spans sorted by start position.
    """
    formulas: list[FormulaSpan] = []

    # Match block formulas $$...$$
    # Uses DOTALL to allow multi-line block formulas
    block_pattern = r"\$\$\s*(.+?)\s*\$\$"
    for match in re.finditer(block_pattern, md_text, re.DOTALL):
        latex = match.group(1).strip()
        # Preserve newlines for \begin...\end environments, collapse otherwise
        if r"\begin" not in latex:
            latex = latex.replace("\n", " ")

        formulas.append(
            {
                "latex": latex,
                "is_inline": False,
                "start": match.start(),
                "end": match.end(),
            }
        )

    # Match inline formulas $...$ (but not $$)
    # Uses lookbehind/lookahead to avoid matching $$
    inline_pattern = r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)"
    for match in re.finditer(inline_pattern, md_text):
        # Skip if this inline match falls inside a block formula
        is_in_block = any(f["start"] <= match.start() < f["end"] for f in formulas if not f["is_inline"])
        if not is_in_block:
            formulas.append(
                {
                    "latex": match.group(1).strip(),
                    "is_inline": True,
                    "start": match.start(),
                    "end": match.end(),
                }
            )

    formulas.sort(key=lambda x: x["start"])
    return formulas
