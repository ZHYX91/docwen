"""Shared text decoding for delimited spreadsheet inputs."""

from __future__ import annotations

import codecs
from pathlib import Path

from docwen_core.detection import delimited_text_encodings

_SAMPLE_LIMIT = 65536


def decoded_samples(file_path: str) -> list[tuple[str, str]]:
    """Decode a bounded prefix without treating a split character as corruption."""
    with Path(file_path).open("rb") as stream:
        raw = stream.read(_SAMPLE_LIMIT + 1)
    final = len(raw) <= _SAMPLE_LIMIT
    sample = raw[:_SAMPLE_LIMIT]
    candidates = delimited_text_encodings(file_path)
    result: list[tuple[str, str]] = []
    for encoding in candidates:
        try:
            text = codecs.getincrementaldecoder(encoding)(errors="strict").decode(sample, final=final)
        except UnicodeError:
            continue
        result.append((encoding, text))
    return result
