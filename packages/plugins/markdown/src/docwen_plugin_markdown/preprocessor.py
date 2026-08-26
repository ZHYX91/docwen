"""Preprocessing: heading merge detection, image materialization, HTML cleanup.

All functions operate on raw markdown text **before** mistune parsing.
"""

from __future__ import annotations

import re
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

    result: list[str] = []
    for fenced_text, is_fenced in split_markdown_block_segments(md_body):
        if is_fenced:
            result.append(fenced_text)
            continue
        for inline_text, is_inline_code in split_markdown_inline_segments(fenced_text):
            result.append(
                inline_text
                if is_inline_code
                else _replace_image_placeholders(
                    inline_text,
                    placeholder_re,
                    decode_path=image_scope is not None,
                )
            )

    return "".join(result)


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
    """Convert Setext headings (=== and ---) to ATX format (# and ##).

    Args:
        md_body: Raw markdown text.

    Returns:
        Markdown text with Setext headings converted to ATX headings.
    """
    # Process H1 (===) first, then H2 (---) — order matters since ---
    # also matches thematic breaks, but regex multiline anchoring handles that.
    md_body = _SETEXT_H1_RE.sub(r"# \1", md_body)
    md_body = _SETEXT_H2_RE.sub(r"## \1", md_body)
    return md_body


# ═══════════════════════════════════════════════════════════════════════════
# Heading merge detection
# ═══════════════════════════════════════════════════════════════════════════


def detect_heading_merges(
    md_body: str,
    mode: str = "punct_required",
    punctuation: frozenset[str] | None = None,
) -> set[int]:
    """Return 0‑based heading indexes that should merge with next body text.

    The next source line must be an immediately adjacent plain-text line.
    Blank lines and Markdown block constructs deliberately break the merge.
    ``"always"`` removes only the punctuation requirement; it does not allow
    merging a list, table, quote, formula, code block, or thematic break.

    Args:
        md_body: Raw markdown source.
        mode: ``"punct_required"``, ``"always"``, or ``"never"``.
        punctuation: Set of punctuation chars that trigger merge. Uses
            a sensible default for Chinese + English punctuation.
    Returns:
        Set of 0‑based heading indexes (in order of appearance) to merge.
    """
    if mode not in {"punct_required", "always", "never"}:
        mode = "punct_required"
    if mode == "never":
        return set()

    punct = punctuation if punctuation is not None else HEADING_MERGE_PUNCTUATION_SET
    if mode == "punct_required" and not punct:
        return set()

    lines = md_body.split("\n")
    merges: set[int] = set()
    heading_idx = 0

    i = 0
    while i < len(lines):
        line = lines[i]
        heading_match = _ATX_HEADING_RE.match(line)
        if heading_match is None:
            i += 1
            continue

        content = re.sub(r"[ \t]+#+[ \t]*$", "", heading_match.group(2)).strip()
        punctuation_allows_merge = mode == "always" or bool(content and content[-1] in punct)
        if punctuation_allows_merge and i + 1 < len(lines) and _is_plain_merge_body_line(lines[i + 1]):
            merges.add(heading_idx)

        heading_idx += 1
        i += 1

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
    """Return line indexes of HRs that should attach to preceding paragraph.

    A horizontal rule (``---``, ``***``, ``___``) is considered "attached"
    when the previous line is non-blank content (not a heading, not blank).

    Args:
        md_body: Raw markdown source.

    Returns:
        Set of 0-based line indexes where an attached HR occurs.
    """
    lines = md_body.split("\n")
    attached: set[int] = set()
    for i in range(1, len(lines)):
        stripped = lines[i].strip()
        prev = lines[i - 1].strip()
        # Check if current line is HR and previous line is non-blank content
        if stripped in ("---", "***", "___") and prev and not prev.startswith("#"):
            attached.add(i)
    return attached


# ═══════════════════════════════════════════════════════════════════════════
# HTML tag normalisation
# ═══════════════════════════════════════════════════════════════════════════

_BR_RE = re.compile(r"<br\s*/?>", re.IGNORECASE)


def normalize_html_tags(md_body: str) -> str:
    """Preprocess HTML tags before markdown parsing.

    - ``<br>``, ``<br/>`` → two trailing spaces plus newline (hard break)
    - ``<u>``, ``<sub>``, ``<sup>`` → preserved for mistune inline passthrough

    Args:
        md_body: Raw markdown text.

    Returns:
        Preprocessed markdown text.
    """
    # <br> → markdown hard line break (two spaces + newline)
    result = _BR_RE.sub("  \n", md_body)
    return result
