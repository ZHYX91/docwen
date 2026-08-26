"""Image-OCR text presentation helpers for Markdown export."""

from __future__ import annotations

from docwen_core.text.splitting import split_once

# Compatibility order from the image OCR renderer. This deliberately chooses
# punctuation by type priority rather than the earliest punctuation in the line.
_OCR_HEADING_BODY_DELIMITERS = ("：", ":", "。", "．", "！", "!")


def split_ocr_heading_body(line: str) -> tuple[str, str]:
    """Split one OCR line into an emphasized prefix and its remaining body.

    The delimiter remains in the prefix. This is an image-OCR blockquote
    presentation heuristic, not a general Markdown or document-parsing rule.
    """

    for delimiter in _OCR_HEADING_BODY_DELIMITERS:
        position = line.find(delimiter)
        if position != -1:
            heading, body = split_once(line, position + len(delimiter))
            return heading.strip(), body.strip()
    return line, ""
