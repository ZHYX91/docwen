"""Markdown utility functions for filename sanitization and link formatting.

F-I2b-004: ``format_sanitized_image_link`` — style-aware image link formatting with
           filename sanitization (distinct from the simpler
           ``export_semantics.format_image_link`` which takes ``(alt, target, style)``
           without sanitization).
"""

from __future__ import annotations

import logging
import re
import urllib.parse

logger = logging.getLogger(__name__)

# ── Sanitization patterns ────────────────────────────────────────────────

# File-system illegal characters (Windows + cross-platform)
_FS_ILLEGAL_RE = re.compile(r'[\\/:*?"<>|]')
# Wiki link syntax-sensitive characters (Obsidian / Logseq common subset)
_WIKI_SENSITIVE_RE = re.compile(r"[\[\]|#^]")
# Consecutive whitespace compression
_MULTI_SPACE_RE = re.compile(r"\s+")


def sanitize_filename(name: str) -> str:
    """Replace file-system-illegal characters with ``_``, compress whitespace,
    and strip leading/trailing dots and spaces.

    Does **not** apply URL percent-encoding — that is the caller's
    responsibility when constructing a Markdown link URL.
    """
    name = _FS_ILLEGAL_RE.sub("_", name)
    name = _MULTI_SPACE_RE.sub(" ", name)
    return name.strip(". ")


def sanitize_for_wiki_link(name: str) -> str:
    """Sanitize a filename for use inside a wiki-style link (``[[ ]]`` / ``![[ ]]``).

    Applies :func:`sanitize_filename` first, then strips characters that
    have special meaning in wiki link syntax (``[ ] | # ^``).
    """
    name = sanitize_filename(name)
    return _WIKI_SENSITIVE_RE.sub("", name)


# ── Image link formatting ────────────────────────────────────────────────

_VALID_LINK_STYLES = frozenset({"wiki_embed", "wiki_link", "markdown_embed", "markdown_link"})


def format_sanitized_image_link(
    filename: str,
    style: str = "wiki_embed",
) -> str:
    """Format an image link with style-dependent filename sanitization.

    **Wiki styles** (``wiki_embed``, ``wiki_link``):
        The *filename* is sanitised via :func:`sanitize_for_wiki_link` and
        wrapped in ``![[...]]`` or ``[[...]]``.

    **Markdown styles** (``markdown_embed``, ``markdown_link``):
        The *filename* is sanitised via :func:`sanitize_filename`, then
        URL-encoded via :func:`urllib.parse.quote`, and wrapped in
        ``![...](...)`` or ``[...](...)``.

    .. note::

        This function differs from :func:`docwen_core.export_semantics.format_image_link`
        which accepts ``(alt, target, style)`` without applying filename
        sanitization.  Use this function when the *filename* itself is the
        link target (common in image-to-Markdown converters).  Use the
        export_semantics version when *alt* and *target* are independently
        controlled (e.g. data URIs or user-supplied alt text).

    Args:
        filename: Raw image filename (may contain spaces, special chars).
        style: One of ``"wiki_embed"``, ``"wiki_link"``,
               ``"markdown_embed"``, ``"markdown_link"``.

    Returns:
        The formatted link string.
    """
    if style not in _VALID_LINK_STYLES:
        logger.debug("Unknown link style %r, falling back to wiki_embed", style)
        style = "wiki_embed"

    if style in ("wiki_embed", "wiki_link"):
        safe = sanitize_for_wiki_link(filename)
        prefix = "!" if style == "wiki_embed" else ""
        result = f"{prefix}[[{safe}]]"
    else:
        safe_fs = sanitize_filename(filename)
        safe_url = urllib.parse.quote(safe_fs)
        result = f"![{safe_fs}]({safe_url})" if style == "markdown_embed" else f"[{safe_fs}]({safe_url})"

    logger.debug("Formatted image link: %r → %r (style=%s)", filename, result, style)
    return result


def format_md_file_link(
    filename: str,
    style: str = "wiki_embed",
) -> str:
    """Format a Markdown-file link in the requested *style*.

    **Wiki styles** (``wiki_embed``, ``wiki_link``):
        The *filename* is sanitised via :func:`sanitize_for_wiki_link` and
        wrapped in ``![[...]]`` or ``[[...]]``.

    **Markdown styles** (``markdown_embed``, ``markdown_link``):
        The *filename* is sanitised via :func:`sanitize_filename` and
        wrapped in ``[...](...)``.  ``markdown_embed`` falls back to a
        standard link because Markdown cannot natively embed a .md file.

    .. note::

        Unlike :func:`format_sanitized_image_link`, this function does NOT
        apply URL percent-encoding — .md file links use the filename as-is
        (the filename itself is the link text for non-wiki styles).

    Args:
        filename: Raw .md filename (may contain spaces, special chars).
        style: One of ``"wiki_embed"``, ``"wiki_link"``,
               ``"markdown_embed"``, ``"markdown_link"``.

    Returns:
        The formatted link string.
    """
    if style not in _VALID_LINK_STYLES:
        logger.debug("Unknown link style %r, falling back to wiki_embed", style)
        style = "wiki_embed"

    if style in ("wiki_embed", "wiki_link"):
        safe = sanitize_for_wiki_link(filename)
        prefix = "!" if style == "wiki_embed" else ""
        result = f"{prefix}[[{safe}]]"
    else:
        safe_fs = sanitize_filename(filename)
        result = f"[{safe_fs}]({safe_fs})"

    logger.debug("Formatted md file link: %r → %r (style=%s)", filename, result, style)
    return result
