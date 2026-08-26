"""Embedded image processing and placeholder formatting.

Covers F-H2-008: ``process_embedded_image``, ``format_image_placeholder``,
and ``split_alt_text_and_size``.

These utilities provide a single shared implementation for handling embedded
image references (wiki ``![[...]]`` or Markdown ``![alt](url)``) based on a
configurable processing mode.  Each plugin calls into these helpers instead
of duplicating mode-dispatch or placeholder-generation logic.
"""

from __future__ import annotations

import logging
from enum import StrEnum
from pathlib import Path
from urllib.parse import quote

logger = logging.getLogger(__name__)


class EmbeddedImageMode(StrEnum):
    """Processing mode for embedded image references.

    ``embed``
        Replace the image link with an ``{{IMAGE:...}}`` placeholder that
        downstream consumers (e.g. DOCX renderer) can recognise and expand.
    ``keep``
        Preserve the original link text unchanged.
    ``extract_text``
        Keep only the alt / display text, falling back to the image filename.
    ``remove``
        Drop the image link entirely (produce an empty string).
    """

    EMBED = "embed"
    KEEP = "keep"
    EXTRACT_TEXT = "extract_text"
    REMOVE = "remove"


def format_image_placeholder(
    image_path: str,
    width: int | None = None,
    height: int | None = None,
    *,
    image_scope: str | None = None,
) -> str:
    """Format an ``{{IMAGE:...}}`` placeholder for embedded images.

    When neither *width* nor *height* is given the placeholder contains only
    the file path::

        {{IMAGE:photo.png}}

    When dimensions are provided they appear as pipe-separated values::

        {{IMAGE:photo.png|200|150}}
        {{IMAGE:photo.png|200|}}          (width only)
    """
    marker = "IMAGE" if image_scope is None else f"IMAGE@{image_scope}"
    marker_path = image_path if image_scope is None else quote(image_path, safe="/:\\._-~")
    if width is None and height is None:
        return f"{{{{{marker}:{marker_path}}}}}"
    w = "" if width is None else str(width)
    h = "" if height is None else str(height)
    return f"{{{{{marker}:{marker_path}|{w}|{h}}}}}"


def split_alt_text_and_size(
    alt_text: str | None,
) -> tuple[str | None, int | None, int | None]:
    """Parse alt text that may carry pipe-separated dimension info.

    Obsidian / wiki-link convention allows size hints at the end of the
    display text after a ``|`` separator:

    - ``photo|200x150``  → ``("photo", 200, 150)``
    - ``photo|200``      → ``("photo", 200, None)``
    - ``photo|200x``     → ``("photo", 200, None)``  (height omitted)

    When no dimensions are detected the entire *alt_text* is returned as the
    display text.

    Returns ``(display_text, width, height)``.
    """
    if not alt_text:
        return None, None, None
    if "|" not in alt_text:
        return alt_text, None, None
    left, right = alt_text.rsplit("|", 1)
    right = right.strip()
    if not right:
        return alt_text, None, None
    if "x" in right:
        w_str, h_str = right.split("x", 1)
        if w_str.isdigit() and (h_str.isdigit() or h_str == ""):
            width = int(w_str)
            height = int(h_str) if h_str.isdigit() else None
            return left, width, height
        return alt_text, None, None
    if right.isdigit():
        return left, int(right), None
    return alt_text, None, None


def process_embedded_image(
    image_path: str,
    original_link: str,
    mode: EmbeddedImageMode | str,
    *,
    display_text: str | None = None,
    width: int | None = None,
    height: int | None = None,
    image_scope: str | None = None,
) -> str:
    """Process an embedded image link according to *mode*.

    Args:
        image_path: The resolved file path (or filename) for the image.
        original_link: The original link text (e.g. ``![[photo.png]]``).
        mode: How to handle the image — see :class:`EmbeddedImageMode`.
        display_text: Alt / display text.  Used by ``extract_text`` mode.
        width: Image width in pixels (for ``embed`` placeholder).
        height: Image height in pixels (for ``embed`` placeholder).

    Returns:
        The replacement text that should appear in the output document.
    """
    mode = EmbeddedImageMode(mode)

    logger.debug(
        "Processing embedded image: %s | mode: %s",
        Path(image_path).name,
        mode.value,
    )

    if mode == EmbeddedImageMode.EMBED:
        placeholder = format_image_placeholder(
            image_path,
            width=width,
            height=height,
            image_scope=image_scope,
        )
        logger.info("Generated image placeholder: %s", placeholder)
        return placeholder
    elif mode == EmbeddedImageMode.KEEP:
        logger.debug("Keeping original image link: %s", original_link)
        return original_link
    elif mode == EmbeddedImageMode.EXTRACT_TEXT:
        if display_text:
            logger.debug("Extracting image display text: %s", display_text)
            return display_text
        else:
            filename = Path(image_path).name
            logger.debug("Extracting image filename: %s", filename)
            return filename
    elif mode == EmbeddedImageMode.REMOVE:
        logger.debug("Removing image link")
        return ""
    else:
        logger.warning(
            "Unknown embedded image mode: %s, falling back to embed",
            mode.value,
        )
        return format_image_placeholder(
            image_path,
            width=width,
            height=height,
            image_scope=image_scope,
        )
