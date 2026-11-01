"""Small WordprocessingML lexical primitives without document-library imports."""

from __future__ import annotations

import re

_INTEGER = re.compile(r"[+-]?[0-9]+")


def parse_word_xml_int(value: str | None) -> int | None:
    """Read xsd:int, including long leading zeros, without changing old XML."""
    if value is None:
        return None
    collapsed = value.strip(" \t\r\n")
    if _INTEGER.fullmatch(collapsed) is None:
        return None
    magnitude = collapsed.lstrip("+-").lstrip("0") or "0"
    if len(magnitude) > 10:
        return None
    number = int(("-" if collapsed.startswith("-") else "") + magnitude)
    return number if -(2**31) <= number < 2**31 else None
