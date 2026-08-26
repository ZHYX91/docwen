"""Link recognition regex patterns and extension constants shared across
``docwen_core.links``.

These are the canonical definitions for wiki-embed, wiki-embed-with-size,
and Markdown embed-image patterns.  Every consumer that needs to match
Obsidian / wiki-style link syntax imports them from here rather than
inlining private regex objects.

Covers:
- F-H2-023: ``WIKI_EMBED_PATTERN``
- F-H2-024: ``WIKI_EMBED_SIZE_PATTERN``
"""

from __future__ import annotations

# ── Recognised file extensions ──────────────────────────────────────────────

IMAGE_EXTENSIONS: frozenset[str] = frozenset({".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".webp", ".ico"})
"""File extensions recognised as image types."""

MD_EXTENSIONS: frozenset[str] = frozenset({".md", ".markdown"})
"""File extensions recognised as Markdown types."""

# ── Wiki embed patterns ─────────────────────────────────────────────────────

WIKI_EMBED_PATTERN = (
    r"!\[\["
    r"((?:[^|\]\\]|\\(?![|]).)*)"  # group 1: link target
    r"(?:(?:\\)?\|"
    # A display label may itself end in ``]``.  When three or more closing
    # brackets occur, consume the leading bracket(s) as label text and reserve
    # the final pair for the wiki-link delimiter.
    r"((?:[^\]\\]|\\(?![|]).|\](?!\])|\](?=\]\]))*)?"  # group 2: optional display text
    r")?\]\]"
)
r"""Regex matching Obsidian / wiki-style embed links.

Matches::

    ![[target]]
    ![[target|display]]
    ![[target#heading|display]]
    ![[target#^block-id]]

Group 1: link target (file path with optional ``#anchor`` / ``#^block-id``).
Group 2: display text (optional, after the ``|``).
"""

WIKI_EMBED_SIZE_PATTERN = r"!\[\[([^|\]]+)\|(\d+)(?:x(\d+))?\]\]"
r"""Regex matching wiki embed links with explicit pixel dimensions.

Matches::

    ![[image.png|200x150]]
    ![[image.png|200]]
    ![[image.png|200x]]

Group 1: link target (file path).
Group 2: width in pixels.
Group 3: height in pixels (optional — may be ``None``).
"""

# ── Markdown embed-image pattern ────────────────────────────────────────────

MD_EMBED_IMAGE_PATTERN = r"!\[([^\]]*)\]\(([^)\s]+)(?:\s*=\s*(\d*)x(\d*))?\s*\)"
r"""Regex matching standard Markdown image embed syntax with optional size.

Matches::

    ![alt text](image.png)
    ![alt text](image.png =200x150)
    ![alt text](image.png =x150)

Group 1: alt text.
Group 2: URL / file path.
Group 3: width in pixels (optional).
Group 4: height in pixels (optional).
"""
