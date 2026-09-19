"""Request-scoped image materialization before Markdown parsing.

All functions operate on raw markdown text **before** mistune parsing. Any
rewrite that is presentation-oriented must leave literal source regions intact;
source-semantic recovery binds those regions before the generic Markdown parser
runs.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from pathlib import Path
from urllib.parse import quote, unquote

from docwen_core.links import (
    split_markdown_block_segments,
    split_markdown_inline_segments,
)

# ── Wiki link patterns ───────────────────────────────────────────────────

_IMAGE_PLACEHOLDER_RE = re.compile(r"\{\{IMAGE:([^{}\r\n]+)\}\}")


def _image_placeholder_re(image_scope: str | None) -> re.Pattern[str]:
    if image_scope is None:
        return _IMAGE_PLACEHOLDER_RE
    return re.compile(rf"\{{\{{IMAGE@{re.escape(image_scope)}:([^{{}}\r\n]+)\}}\}}")


def _rewrite_non_code_markdown(text: str, rewrite: Callable[[str], str]) -> str:
    """Apply an inline rewrite only to ordinary Markdown source.

    ``split_markdown_inline_segments`` protects renderer atoms such as code,
    inline HTML and inline math, not code alone. This helper is therefore for
    transforms (for example image-placeholder materialization) that must not
    enter any protected atom.
    """

    result: list[str] = []
    for block_text, is_literal_block in split_markdown_block_segments(text):
        if is_literal_block:
            result.append(block_text)
            continue
        for inline_text, is_protected_atom in split_markdown_inline_segments(block_text):
            result.append(inline_text if is_protected_atom else rewrite(inline_text))
    return "".join(result)


def materialize_image_placeholders(
    md_body: str,
    *,
    image_scope: str | None = None,
) -> str:
    """Turn core image placeholders into table-safe Markdown images.

    ``process_markdown_links`` emits ``{{IMAGE:path|width|height}}`` for an
    embedded image. Passing that representation directly to Mistune would
    leave a literal placeholder in paragraphs and, more importantly, split a
    table cell at each dimension pipe. This adapter uses an angle-bracketed
    Markdown destination and carries optional dimensions in the title, which
    contains no table delimiters. Literal renderer atoms remain untouched.
    """

    marker = "{{IMAGE:" if image_scope is None else f"{{{{IMAGE@{image_scope}:"
    if marker not in md_body:
        return md_body

    placeholder_re = _image_placeholder_re(image_scope)
    return _rewrite_non_code_markdown(
        md_body,
        lambda text: _replace_image_placeholders(
            text,
            placeholder_re,
            decode_path=image_scope is not None,
        ),
    )


def _replace_image_placeholders(
    text: str,
    placeholder_re: re.Pattern[str],
    *,
    decode_path: bool = False,
) -> str:
    def replace(match: re.Match[str]) -> str:
        payload = match.group(1)
        image_path, width, height = _parse_image_placeholder_payload(
            payload,
            decode_path=decode_path,
        )
        if not image_path:
            return match.group(0)

        normalized_path = image_path.replace("\\", "/")
        markdown_path = quote(normalized_path, safe="/:._-~") if decode_path else normalized_path
        alt_text = Path(normalized_path).name.replace("[", r"\[").replace("]", r"\]")
        title = ""
        if width is not None or height is not None:
            size = f"{width or ''}x{height or ''}"
            title = f' "docwen-size={size}"'
        return f"![{alt_text}](<{markdown_path}>{title})"

    return placeholder_re.sub(replace, text)


def _parse_image_placeholder_payload(
    payload: str,
    *,
    decode_path: bool = False,
) -> tuple[str, int | None, int | None]:
    parts = payload.rsplit("|", 2)
    if len(parts) != 3:
        image_path = payload
        if decode_path:
            image_path = unquote(image_path)
        return image_path, None, None
    image_path, width_text, height_text = parts
    if (width_text and not width_text.isdigit()) or (height_text and not height_text.isdigit()):
        image_path = payload
        if decode_path:
            image_path = unquote(image_path)
        return image_path, None, None
    if decode_path:
        image_path = unquote(image_path)
    width = int(width_text) if width_text else None
    height = int(height_text) if height_text else None
    return image_path, width, height
