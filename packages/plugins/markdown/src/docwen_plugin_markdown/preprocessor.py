"""Preprocessing: heading merge detection, image materialization, HTML cleanup.

All functions operate on raw markdown text **before** mistune parsing.  Any
rewrite that is presentation-oriented must leave fenced and inline code
byte-for-byte untouched; source-semantic recovery binds those literal regions
before the generic Markdown parser runs.
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
from docwen_core.text.heading_merge import HEADING_MERGE_PUNCTUATION_SET

# ── Wiki link patterns ───────────────────────────────────────────────────

_IMAGE_PLACEHOLDER_RE = re.compile(r"\{\{IMAGE:([^{}\r\n]+)\}\}")


def _image_placeholder_re(image_scope: str | None) -> re.Pattern[str]:
    if image_scope is None:
        return _IMAGE_PLACEHOLDER_RE
    return re.compile(rf"\{{\{{IMAGE@{re.escape(image_scope)}:([^{{}}\r\n]+)\}}\}}")


def _rewrite_non_code_markdown(text: str, rewrite: Callable[[str], str]) -> str:
    """Apply *rewrite* only to ordinary Markdown source.

    Several historical preprocessors predate source-semantic recovery and used
    whole-document regular expressions.  That can turn literal examples inside
    fenced or inline code into real Markdown structure.  Keep one shared
    boundary here so future textual normalizers cannot accidentally repeat that
    class of bug.
    """

    result: list[str] = []
    for block_text, is_fenced in split_markdown_block_segments(text):
        if is_fenced:
            result.append(block_text)
            continue
        for inline_text, is_inline_code in split_markdown_inline_segments(block_text):
            result.append(inline_text if is_inline_code else rewrite(inline_text))
    return "".join(result)


def materialize_image_placeholders(
    md_body: str,
    *,
    image_scope: str | None = None,
) -> str:
    """Turn core image placeholders into table-safe Markdown images.

    ``process_markdown_links`` emits ``{{IMAGE:path|width|height}}`` for an
    embedded image.  Passing that representation directly to Mistune would
    leave a literal placeholder in paragraphs and, more importantly, split a
    table cell at each dimension pipe.  This adapter uses an angle-bracketed
    Markdown destination and carries optional dimensions in the title, which
    contains no table delimiters.  Fenced and inline code remain literal.
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


_ATX_HEADING_RE = re.compile(r"^ {0,3}(#{1,9})(?!#)(?:[ \t]+|$)(.*)$")
_UNORDERED_LIST_RE = re.compile(r"^ {0,3}[*+-](?:[ \t]+|$)")
_ORDERED_LIST_RE = re.compile(r"^ {0,3}\d{1,9}[.)](?:[ \t]+|$)")
_THEMATIC_BREAK_RE = re.compile(r"^ {0,3}(?:(?:\*[ \t]*){3,}|(?:-[ \t]*){3,}|(?:_[ \t]*){3,})$")

# ═══════════════════════════════════════════════════════════════════════════
# Setext heading conversion
# ═══════════════════════════════════════════════════════════════════════════

_SETEXT_H1_RE = re.compile(
    r"^([^\r\n]*[^ \t\r\n][^\r\n]*)\r?\n={3,}[ \t]*(?=\r?$)",
    re.MULTILINE,
)
_SETEXT_H2_RE = re.compile(
    r"^([^\r\n]*[^ \t\r\n][^\r\n]*)\r?\n-{3,}[ \t]*(?=\r?$)",
    re.MULTILINE,
)


def handle_setext_headings(md_body: str) -> str:
    """Convert Setext headings to ATX without rewriting literal code.

    Mistune understands Setext headings itself, but the renderer's historical
    heading-merge path consumes the ATX projection.  Until that path is fully
    source-position driven, keep this narrow adapter while respecting Markdown
    literal boundaries.
    """

    def rewrite(text: str) -> str:
        # Process H1 first, then H2 — order matters because ``---`` also has
        # thematic-break syntax outside a Setext pair.
        text = _SETEXT_H1_RE.sub(r"# \1", text)
        return _SETEXT_H2_RE.sub(r"## \1", text)

    return _rewrite_non_code_markdown(md_body, rewrite)


# ═══════════════════════════════════════════════════════════════════════════
# Heading merge detection
# ═══════════════════════════════════════════════════════════════════════════


def detect_heading_merges(
    md_body: str,
    mode: str = "punct_required",
    punctuation: frozenset[str] | None = None,
) -> set[int]:
    """Return 0-based rendered-heading indexes that merge with following text.

    Fenced code is never a heading source and therefore does not consume a
    heading index.  This keeps indexes aligned with the AST even when examples
    contain lines beginning with ``#``.
    """
    if mode not in {"punct_required", "always", "never"}:
        mode = "punct_required"
    if mode == "never":
        return set()

    punct = punctuation if punctuation is not None else HEADING_MERGE_PUNCTUATION_SET
    if mode == "punct_required" and not punct:
        return set()

    merges: set[int] = set()
    heading_idx = 0
    for block_text, is_fenced in split_markdown_block_segments(md_body):
        if is_fenced:
            continue
        lines = block_text.split("\n")
        for index, line in enumerate(lines):
            heading_match = _ATX_HEADING_RE.match(line)
            if heading_match is None:
                continue
            content = re.sub(r"[ \t]+#+[ \t]*$", "", heading_match.group(2)).strip()
            punctuation_allows_merge = mode == "always" or bool(content and content[-1] in punct)
            if punctuation_allows_merge and index + 1 < len(lines) and _is_plain_merge_body_line(lines[index + 1]):
                merges.add(heading_idx)
            heading_idx += 1
    return merges


def _is_plain_merge_body_line(line: str) -> bool:
    """Whether *line* is the adjacent ordinary body text accepted by old DocWen."""

    stripped = line.strip()
    if not stripped:
        return False
    if line.startswith(("    ", "\t")):
        return False
    if _ATX_HEADING_RE.match(line):
        return False
    if stripped.startswith(("$$", "|", ">", "```", "~~~")):
        return False
    if _UNORDERED_LIST_RE.match(line) or _ORDERED_LIST_RE.match(line):
        return False
    return _THEMATIC_BREAK_RE.match(line) is None


# ═══════════════════════════════════════════════════════════════════════════
# HR attachment detection
# ═══════════════════════════════════════════════════════════════════════════


def detect_hr_attachments(md_body: str) -> set[int]:
    """Return source line indexes of non-code HRs attached to prior content."""

    attached: set[int] = set()
    line_offset = 0
    for block_text, is_fenced in split_markdown_block_segments(md_body):
        lines = block_text.split("\n")
        if not is_fenced:
            for index in range(1, len(lines)):
                stripped = lines[index].strip()
                prev = lines[index - 1].strip()
                if stripped in ("---", "***", "___") and prev and not prev.startswith("#"):
                    attached.add(line_offset + index)
        line_offset += block_text.count("\n")
    return attached


# ═══════════════════════════════════════════════════════════════════════════
# HTML tag normalisation
# ═══════════════════════════════════════════════════════════════════════════

_BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)


def normalize_html_tags(md_body: str) -> str:
    """Normalize supported HTML only outside fenced and inline code."""

    return _rewrite_non_code_markdown(md_body, lambda text: _BR_RE.sub("  \n", text))
