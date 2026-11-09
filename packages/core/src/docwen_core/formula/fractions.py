"""Fraction properties shared by the MathML output converters."""

from __future__ import annotations

import re
from typing import Literal

_ZERO_LENGTH = re.compile(r"[+-]?(?:0+(?:\.0*)?|\.0+)(?:em|ex|px|in|cm|mm|pt|pc|%)?")


def mathml_fraction_type(node) -> Literal["bar", "noBar"]:
    """Return a supported fraction layout, refusing properties we would lose.

    A zero rule is a stacked expression, such as a binomial coefficient; it
    must never acquire a division bar. Skewed fractions, custom rule widths
    and non-centered arguments have no faithful representation in our shared
    LaTeX/Office Math subset, so callers must use their explicit fallback.
    """
    for name, default in (("bevelled", "false"), ("numalign", "center"), ("denomalign", "center")):
        value = node.get(name, default).strip()
        if value != default:
            raise ValueError(f"unsupported fraction property: {name}={value!r}")

    thickness = node.get("linethickness", "medium").strip()
    if _ZERO_LENGTH.fullmatch(thickness):
        return "noBar"
    if thickness != "medium":
        raise ValueError(f"unsupported fraction property: linethickness={thickness!r}")
    return "bar"
