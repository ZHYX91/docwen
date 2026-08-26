"""Centralised link-processing error semantics.

Provides a single, well-typed classification of link-error kinds together
with a deterministic placeholder / keep-link generator that all consumers
(direct callers of ``process_embedded_md_file``, the batch resolver, and
eventually plugin converters) can use.  This removes the need for every
call site to construct ``File not found: ...`` strings by hand.

Covers:
- F-H2-010: section-not-found → differentiated placeholder
- F-H2-011: block-not-found   → differentiated placeholder
- F-H2-012: file-not-found    → mode-aware output (ignore/keep/placeholder)
- F-H2-013: circular-ref      → mode-aware output
- F-H2-014: max-depth         → mode-aware output
"""

from __future__ import annotations

from enum import IntEnum, StrEnum

# ═══════════════════════════════════════════════════════════════════════════
# Error action mode
# ═══════════════════════════════════════════════════════════════════════════


class NotFoundAction(StrEnum):
    """Action when a referenced section, block, file, or recursion limit is
    encountered during link processing.

    ``ignore``
        Return an empty string.
    ``keep``
        Preserve the original ``![[...]]`` link text.
    ``placeholder``
        Return a human-readable placeholder string.
    """

    IGNORE = "ignore"
    KEEP = "keep"
    PLACEHOLDER = "placeholder"


# ═══════════════════════════════════════════════════════════════════════════
# Error kind classification
# ═══════════════════════════════════════════════════════════════════════════


class LinkErrorKind(IntEnum):
    """Fine-grained classification of every link-processing error condition.

    Values are ordered by severity (lower = more specific / actionable).
    """

    SECTION_NOT_FOUND = 1
    """A heading (``#heading``) was referenced but does not exist in the target file."""

    BLOCK_NOT_FOUND = 2
    """A block-id (``#^block-id``) was referenced but was not found."""

    FILE_NOT_FOUND = 3
    """The linked file could not be resolved to any existing path."""

    CIRCULAR_REFERENCE = 4
    """A self-referencing or mutually-referencing embed loop was detected."""

    MAX_DEPTH = 5
    """The maximum recursive embed depth was reached."""

    @property
    def label(self) -> str:
        """Human-readable Chinese label for the error kind."""
        _labels: dict[LinkErrorKind, str] = {
            LinkErrorKind.SECTION_NOT_FOUND: "章节未找到",
            LinkErrorKind.BLOCK_NOT_FOUND: "块未找到",
            LinkErrorKind.FILE_NOT_FOUND: "文件未找到",
            LinkErrorKind.CIRCULAR_REFERENCE: "循环引用",
            LinkErrorKind.MAX_DEPTH: "达到最大深度",
        }
        return _labels[self]


# ═══════════════════════════════════════════════════════════════════════════
# Public API
# ═══════════════════════════════════════════════════════════════════════════


def make_error_placeholder(
    kind: LinkErrorKind | int,
    filename: str,
    *,
    heading: str | None = None,
    block_id: str | None = None,
) -> str:
    """Build a human-readable, deterministic placeholder string for *kind*.

    The output is always a bracket-wrapped ``[<category>: <target>]``
    string that downstream consumers can match against in tests and that
    end users can recognise in rendered documents.

    Args:
        kind: The error category (see :class:`LinkErrorKind`).
        filename: The file name portion of the target (e.g. ``readme.md``).
        heading: The heading name, if the embed targeted a heading.
        block_id: The block identifier, if the embed targeted a block.

    Returns:
        A placeholder string such as ``[Section not found: file.md#heading]``.
    """
    kind = LinkErrorKind(kind)

    _english_reasons: dict[LinkErrorKind, str] = {
        LinkErrorKind.SECTION_NOT_FOUND: "Section not found",
        LinkErrorKind.BLOCK_NOT_FOUND: "Block not found",
        LinkErrorKind.FILE_NOT_FOUND: "File not found",
        LinkErrorKind.CIRCULAR_REFERENCE: "Circular reference",
        LinkErrorKind.MAX_DEPTH: "Max depth reached",
    }
    reason = _english_reasons.get(kind, "Error")

    desc = filename
    if heading:
        desc += f"#{heading}"
    elif block_id:
        desc += f"#^{block_id}"

    return f"[{reason}: {desc}]"


def make_keep_link(
    filename: str,
    *,
    heading: str | None = None,
    block_id: str | None = None,
) -> str:
    """Reconstruct the ``![[...]]`` wiki-embed syntax for KEEP actions.

    Args:
        filename: The file name portion of the target.
        heading: Optional heading anchor.
        block_id: Optional block-id anchor.

    Returns:
        A wiki-embed link string, e.g. ``![[file.md#heading]]``.
    """
    target = filename
    if heading:
        target += f"#{heading}"
    elif block_id:
        target += f"#^{block_id}"
    return f"![[{target}]]"


def dispatch_error_output(
    kind: LinkErrorKind | int,
    mode: NotFoundAction | str,
    filename: str,
    *,
    heading: str | None = None,
    block_id: str | None = None,
    original_link: str | None = None,
) -> str:
    """Single entry point for error-mode dispatch.

    This is the function every consumer should call when a link-processing
    error is encountered.  It maps *mode* to the correct output:

    ===============  ======================================================
    ``ignore``       ``""`` (empty string)
    ``keep``         Reconstructed ``![[...]]`` link via *original_link* or fallback.
    ``placeholder``  ``[<category>: <file>#<anchor>]`` via :func:`make_error_placeholder`.
    ===============  ======================================================

    Args:
        kind: The error category.
        mode: How to handle the error (``"ignore"`` / ``"keep"`` / ``"placeholder"``).
        filename: The file name portion of the target.
        heading: Optional heading anchor.
        block_id: Optional block-id anchor.
        original_link: The original link text (used by ``keep``).  When
            ``None``, :func:`make_keep_link` synthesises a link.

    Returns:
        The replacement text for the error condition.
    """
    mode = NotFoundAction(mode)

    if mode == NotFoundAction.IGNORE:
        return ""
    if mode == NotFoundAction.KEEP:
        if original_link is not None:
            return original_link
        return make_keep_link(filename, heading=heading, block_id=block_id)
    # placeholder (default)
    return make_error_placeholder(
        kind,
        filename,
        heading=heading,
        block_id=block_id,
    )
