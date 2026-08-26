"""Markdown file embedding with recursive expansion, circular reference
detection, and precision extraction (heading / block_id).

Covers F-H2-009: ``process_embedded_md_file``.
"""

from __future__ import annotations

import hashlib
import hmac
import logging
import re
import secrets
from collections.abc import Callable
from enum import StrEnum
from pathlib import Path

from docwen_core.links._anchor import (
    extract_block_by_id,
    extract_section_by_heading,
    strip_yaml_front_matter,
)
from docwen_core.links._error_semantics import (
    LinkErrorKind,
    NotFoundAction,
    dispatch_error_output,
    make_error_placeholder,
    make_keep_link,
)

logger = logging.getLogger(__name__)

TABLE_CELL_BR_TOKEN = "{{DOCWEN_BR}}"
_TABLE_BREAK_TOKEN_RE = re.compile(r"\{\{DOCWEN_BR@([A-Za-z0-9_-]{32})\.([0-9a-f]{16})\}\}")
_TABLE_BREAK_SECRET = secrets.token_bytes(32)

# Callable for recursive link processing.
# Signature: (content: str, source_path: str, visited_files: set[str] | None,
#             depth: int, *, table_safe: bool = False) -> str
ProcessLinksFunc = Callable[..., str]


class EmbeddedMdMode(StrEnum):
    """Processing mode for embedded Markdown file references.

    ``embed``
        Read the file, strip YAML front matter, extract the requested
        section / block, and recursively expand any embedded links.
    ``keep``
        Preserve the original ``![[...]]`` link text unchanged.
    ``extract_text``
        Return the *display_text* if provided, otherwise the file stem.
    ``remove``
        Drop the embed entirely (produce an empty string).
    """

    EMBED = "embed"
    KEEP = "keep"
    EXTRACT_TEXT = "extract_text"
    REMOVE = "remove"


# ── helpers ────────────────────────────────────────────────────────────────


def _fmt_keep_link(
    filename: str,
    heading: str | None = None,
    block_id: str | None = None,
) -> str:
    """Reconstruct the ``![[...]]`` wiki-embed syntax for KEEP actions."""
    return make_keep_link(filename, heading=heading, block_id=block_id)


def _mk_missing_text(
    reason: str,
    filename: str,
    heading: str | None = None,
    block_id: str | None = None,
) -> str:
    """Build a human-readable placeholder for missing content.

    For known error categories the *reason* string is mapped to the
    appropriate :class:`LinkErrorKind` so that the output format stays
    consistent; unrecognised reasons fall back to a generic format.
    """
    _reason_to_kind: dict[str, LinkErrorKind] = {
        "Section not found": LinkErrorKind.SECTION_NOT_FOUND,
        "Block not found": LinkErrorKind.BLOCK_NOT_FOUND,
        "File not found": LinkErrorKind.FILE_NOT_FOUND,
        "Circular reference": LinkErrorKind.CIRCULAR_REFERENCE,
        "Max depth reached": LinkErrorKind.MAX_DEPTH,
    }
    kind = _reason_to_kind.get(reason)
    if kind is not None:
        return make_error_placeholder(
            kind,
            filename,
            heading=heading,
            block_id=block_id,
        )
    # Unrecognised reason (e.g. "Error reading") — generic fallback.
    desc = filename
    if heading:
        desc += f"#{heading}"
    elif block_id:
        desc += f"#^{block_id}"
    return f"[{reason}: {desc}]"


# ── main API ───────────────────────────────────────────────────────────────


def process_embedded_md_file(
    md_path: str,
    source_file_path: str,
    visited_files: set[str] | None,
    depth: int,
    process_links: ProcessLinksFunc | None = None,
    *,
    heading: str | None = None,
    block_id: str | None = None,
    display_text: str | None = None,
    original_link: str | None = None,
    mode: EmbeddedMdMode | str = EmbeddedMdMode.EMBED,
    on_not_found: NotFoundAction | str = NotFoundAction.PLACEHOLDER,
    on_circular: NotFoundAction | str = NotFoundAction.PLACEHOLDER,
    detect_circular: bool = True,
    table_safe: bool = False,
) -> str:
    """Process an embedded Markdown file reference and return replacement text.

    Supports Obsidian-style precision embedding:

    * ``![[file.md#heading]]`` — embed only the section under *heading*.
    * ``![[file.md#^block-id]]`` — embed only the paragraph marked
      ``^block-id``.

    When *mode* is ``embed`` the file is read, YAML front matter is
    stripped, the requested section / block is extracted, and
    *process_links* is called on the resulting content so that nested
    ``![[...]]`` references are expanded recursively.

    Args:
        md_path: Absolute or relative path to the ``.md`` file.
        source_file_path: Path of the file containing the embed link.
        visited_files: Set of resolved absolute paths visited during this
            expansion pass (used for circular reference detection).  May be
            ``None``; a new empty set is created internally if needed.
        depth: Current recursion depth.
        process_links: Callback that scans *content* for ``![[...]]``
            patterns and dispatches each one.  The callback receives
            ``(content, source_path, visited_files, depth, *, table_safe)``
            and returns the expanded text.  When ``None``, the raw extracted
            content is returned without further expansion.
        heading: Heading name for section extraction
            (``![[file.md#heading]]``).
        block_id: Block identifier for paragraph extraction
            (``![[file.md#^block-id]]``).
        display_text: Alternative display text used by ``extract_text`` mode.
        original_link: Exact user-authored embed syntax for keep/error output.
        mode: How to process the embed reference
            (:class:`EmbeddedMdMode`).
        on_not_found: Action when a section, block, or file is not found.
        on_circular: Action when a circular reference is detected.
        detect_circular: Set to ``False`` to skip circular reference
            detection.

    Returns:
        The replacement text for the embedded file reference.
    """
    mode = EmbeddedMdMode(mode)
    on_not_found = NotFoundAction(on_not_found)
    on_circular = NotFoundAction(on_circular)

    if visited_files is None:
        visited_files = set()

    filename = Path(md_path).name

    embed_desc = filename
    if heading:
        embed_desc += f"#{heading}"
    elif block_id:
        embed_desc += f"#^{block_id}"

    logger.debug(
        "process_embedded_md_file: %s | mode=%s | depth=%d",
        embed_desc,
        mode.value,
        depth,
    )

    normalized_path = str(Path(md_path).resolve())

    # ── circular-reference check ──────────────────────────────────────
    if detect_circular and normalized_path in visited_files:
        logger.warning(
            "Circular reference: %s (embedded from %s)",
            filename,
            Path(source_file_path).name,
        )
        return dispatch_error_output(
            LinkErrorKind.CIRCULAR_REFERENCE,
            on_circular,
            filename,
            heading=heading,
            block_id=block_id,
            original_link=original_link,
        )

    # ── mode dispatch ─────────────────────────────────────────────────
    if mode == EmbeddedMdMode.EMBED:
        try:
            logger.info("Reading MD file: %s (depth: %d)", embed_desc, depth + 1)
            content = Path(md_path).read_text(encoding="utf-8")
            logger.debug("Read %d characters", len(content))

            # Strip YAML front matter
            content = strip_yaml_front_matter(content)

            # ── section / block extraction ────────────────────────────
            if heading:
                extracted = extract_section_by_heading(content, heading)
                if extracted is None:
                    logger.warning(
                        "Section not found: %s#%s",
                        filename,
                        heading,
                    )
                    return dispatch_error_output(
                        LinkErrorKind.SECTION_NOT_FOUND,
                        on_not_found,
                        filename,
                        heading=heading,
                        original_link=original_link,
                    )
                content = extracted
                logger.info(
                    "Extracted section: %s#%s (%d chars)",
                    filename,
                    heading,
                    len(content),
                )
            elif block_id:
                extracted = extract_block_by_id(content, block_id)
                if extracted is None:
                    logger.warning(
                        "Block not found: %s#^%s",
                        filename,
                        block_id,
                    )
                    return dispatch_error_output(
                        LinkErrorKind.BLOCK_NOT_FOUND,
                        on_not_found,
                        filename,
                        block_id=block_id,
                        original_link=original_link,
                    )
                content = extracted
                logger.info(
                    "Extracted block: %s#^%s (%d chars)",
                    filename,
                    block_id,
                    len(content),
                )

            # ── recursive expansion ───────────────────────────────────
            visited_files.add(normalized_path)
            try:
                if process_links is not None:
                    expanded = process_links(
                        content,
                        md_path,
                        visited_files,
                        depth + 1,
                        table_safe=table_safe,
                    )
                else:
                    expanded = content
            finally:
                visited_files.discard(normalized_path)

            if table_safe:
                expanded = _make_table_safe(expanded)

            logger.info("MD file expansion complete: %s", embed_desc)
            return expanded

        except FileNotFoundError:
            logger.error("MD file not found: %s", md_path)
            return dispatch_error_output(
                LinkErrorKind.FILE_NOT_FOUND,
                on_not_found,
                filename,
                heading=heading,
                block_id=block_id,
                original_link=original_link,
            )
        except OSError as exc:
            logger.error("Failed to read MD file %s: %s", md_path, exc)
            if on_not_found == NotFoundAction.IGNORE:
                return ""
            if on_not_found == NotFoundAction.KEEP:
                return original_link or _fmt_keep_link(filename, heading, block_id)
            return _mk_missing_text(
                "Error reading",
                filename,
                heading,
                block_id,
            )
        except Exception:
            logger.exception("Unexpected error processing: %s", embed_desc)
            return _mk_missing_text(
                "Error reading",
                filename,
                heading,
                block_id,
            )

    elif mode == EmbeddedMdMode.KEEP:
        logger.debug("Keeping original embed link: %s", embed_desc)
        return original_link or _fmt_keep_link(filename, heading, block_id)

    elif mode == EmbeddedMdMode.EXTRACT_TEXT:
        if display_text:
            logger.debug("Extracting display text: %s", display_text)
            return display_text
        stem = Path(md_path).stem
        logger.debug("Extracting filename stem: %s", stem)
        return stem

    elif mode == EmbeddedMdMode.REMOVE:
        logger.debug("Removing embed link: %s", embed_desc)
        return ""

    else:
        logger.warning(
            "Unknown MD embed mode '%s', falling back to embed",
            mode.value,
        )
        return process_embedded_md_file(
            md_path,
            source_file_path,
            visited_files,
            depth,
            process_links,
            heading=heading,
            block_id=block_id,
            display_text=display_text,
            original_link=original_link,
            mode=EmbeddedMdMode.EMBED,
            on_not_found=on_not_found,
            on_circular=on_circular,
            detect_circular=detect_circular,
            table_safe=table_safe,
        )


def _make_table_safe(text: str) -> str:
    """Flatten embedded Markdown so it can live inside one pipe-table cell."""
    normalized = text.strip().replace("\r\n", "\n").replace("\r", "\n")
    if "\n" not in normalized:
        break_token = ""
    else:
        scope = secrets.token_urlsafe(24)
        signature = hmac.new(
            _TABLE_BREAK_SECRET,
            scope.encode("ascii"),
            hashlib.sha256,
        ).hexdigest()[:16]
        break_token = f"{{{{DOCWEN_BR@{scope}.{signature}}}}}"
        while break_token in normalized:
            scope = secrets.token_urlsafe(24)
            signature = hmac.new(
                _TABLE_BREAK_SECRET,
                scope.encode("ascii"),
                hashlib.sha256,
            ).hexdigest()[:16]
            break_token = f"{{{{DOCWEN_BR@{scope}.{signature}}}}}"

    parts: list[str] = []
    backslash_run = 0
    for char in normalized:
        if char == "\\":
            parts.append(char)
            backslash_run += 1
            continue
        if char == "|" and backslash_run % 2 == 0:
            parts.append("\\")
        parts.append(char)
        backslash_run = 0
    escaped = "".join(parts)
    return re.sub(r"\n+", break_token, escaped) if break_token else escaped


def make_table_safe(text: str) -> str:
    """Flatten Markdown and mint authenticated line-break capabilities."""

    return _make_table_safe(text)


def restore_table_safe_breaks(text: str) -> str:
    """Restore only authenticated break capabilities minted in this process."""

    def replace(match: re.Match[str]) -> str:
        scope = match.group(1)
        signature = match.group(2)
        expected = hmac.new(
            _TABLE_BREAK_SECRET,
            scope.encode("ascii"),
            hashlib.sha256,
        ).hexdigest()[:16]
        return "\n" if hmac.compare_digest(signature, expected) else match.group(0)

    return _TABLE_BREAK_TOKEN_RE.sub(replace, text)
