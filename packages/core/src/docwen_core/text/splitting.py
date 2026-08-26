"""Small pure text-splitting primitives without document semantics."""

from __future__ import annotations


def split_once(text: str, boundary: int) -> tuple[str, str]:
    """Split *text* once at an explicit character boundary.

    The caller owns all business decisions about how *boundary* was found.
    """

    if boundary < 0 or boundary > len(text):
        raise ValueError(f"split boundary {boundary} is outside text length {len(text)}")
    return text[:boundary], text[boundary:]
