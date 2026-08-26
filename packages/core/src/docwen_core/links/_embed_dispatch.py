"""Single embed dispatch and batch Markdown-embed resolution.

Covers F-H2-026: ``process_single_embed`` — routes a single embedded link
(data URI image, image file, or Markdown file) to the correct processor so
that every converter can use a single shared dispatch entry point.
"""

from __future__ import annotations

import logging
import re
from collections.abc import Callable, Sequence
from pathlib import Path
from urllib.parse import unquote, urlsplit

from docwen_core.links._anchor import parse_anchor
from docwen_core.links._data_uri import (
    is_data_uri_image,
    resolve_data_uri_image_to_temp_file,
)
from docwen_core.links._embed_image import (
    EmbeddedImageMode,
    process_embedded_image,
    split_alt_text_and_size,
)
from docwen_core.links._error_semantics import (
    LinkErrorKind,
    dispatch_error_output,
)
from docwen_core.links._markdown_inline import encode_markdown_destination_escapes
from docwen_core.links._patterns import WIKI_EMBED_PATTERN, WIKI_EMBED_SIZE_PATTERN
from docwen_core.links._resolver import get_file_type, resolve_file_path

logger = logging.getLogger(__name__)

# Compiled regex for wiki embed link matching, derived from the canonical
# pattern in ``_patterns.py``.
_WIKI_EMBED_RE = re.compile(WIKI_EMBED_PATTERN)
_WIKI_EMBED_SIZE_RE = re.compile(WIKI_EMBED_SIZE_PATTERN)


def _table_row_cells(line: str) -> list[str]:
    """Split a pipe row using parity-aware backslash escaping."""
    cleaned = re.sub(r"^[> \t]+", "", line).strip()
    cells: list[str] = []
    cursor = 0
    backslash_run = 0
    saw_separator = False
    for index, char in enumerate(cleaned):
        if char == "\\":
            backslash_run += 1
            continue
        if char == "|" and backslash_run % 2 == 0:
            cells.append(cleaned[cursor:index])
            cursor = index + 1
            saw_separator = True
        backslash_run = 0
    if not saw_separator:
        return []
    cells.append(cleaned[cursor:])
    if cleaned.startswith("|"):
        cells = cells[1:]
    if cleaned.endswith("|"):
        cells = cells[:-1]
    return [cell.strip() for cell in cells]


def _is_table_separator_line(line: str) -> bool:
    cells = _table_row_cells(line)
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell) is not None for cell in cells)


def _is_probable_table_row_line(line: str) -> bool:
    """Return whether *line* can be a pipe-table data/header row."""
    cells = _table_row_cells(line)
    return bool(cells) and not _is_table_separator_line(line)


def _is_table_context(text: str, start_index: int) -> bool:
    """Return whether a source position is inside a supported pipe table."""
    lines = text.splitlines()
    line_index = len(re.findall(r"\r\n|\r|\n", text[:start_index]))
    if line_index >= len(lines):
        lines.append("")
    if not _is_probable_table_row_line(lines[line_index]):
        return False

    # Header row: the required separator immediately follows it.
    if line_index + 1 < len(lines) and _is_table_separator_line(lines[line_index + 1]):
        return True

    # Body row: walk the contiguous rows upward to the separator/header pair.
    cursor = line_index - 1
    while cursor >= 0 and _is_probable_table_row_line(lines[cursor]):
        cursor -= 1
    return cursor >= 1 and _is_table_separator_line(lines[cursor]) and _is_probable_table_row_line(lines[cursor - 1])


def _local_path_from_file_uri(target: str) -> str:
    """Decode a local or UNC file URI without dropping its authority."""
    parsed = urlsplit(target)
    decoded_path = unquote(parsed.path)
    if parsed.netloc and parsed.netloc.lower() != "localhost":
        return f"//{parsed.netloc}{decoded_path}"
    if re.match(r"^/[A-Za-z]:/", decoded_path):
        return decoded_path[1:]
    return decoded_path


def _unresolved_embed_output(
    original_link: str,
    link_target: str,
    on_not_found: str,
) -> str:
    if on_not_found == "ignore":
        return ""
    if on_not_found == "keep":
        return original_link
    parsed = urlsplit(link_target)
    description = Path(unquote(parsed.path) or link_target).name or link_target
    return f"[File not found: {description}]"


# ── single-embed dispatch ──────────────────────────────────────────────────


def process_single_embed(
    link_target: str,
    original_link: str,
    source_file_path: str,
    visited_files: set[str] | None,
    depth: int,
    *,
    display_text: str | None = None,
    width: int | None = None,
    height: int | None = None,
    image_mode: EmbeddedImageMode | str = EmbeddedImageMode.EMBED,
    md_mode: str = "embed",
    on_not_found: str = "placeholder",
    on_circular: str = "placeholder",
    on_max_depth: str = "placeholder",
    detect_circular: bool = True,
    max_depth: int | None = None,
    search_dirs: Sequence[str] | None = None,
    temp_dir: str | None = None,
    table_safe: bool = False,
    image_scope: str | None = None,
    process_links: Callable[..., str] | None = None,
) -> str | None:
    """Dispatch a single embedded link to the correct processor.

    Flow:

    1. Detect data URI images → materialise to temp file → process as image.
    2. Parse the link target for anchor fragments (``#heading`` /
       ``#^block-id``).
    3. Resolve the target file path against *source_file_path*.
    4. Determine the file type and route accordingly:
       - ``image`` → :func:`process_embedded_image`
       - ``markdown`` → :func:`process_embedded_md_file`
       - ``unknown`` → return ``None`` (caller decides fallback)

    Args:
        link_target: The raw link target, e.g. ``file.md#heading`` or
            ``data:image/png;base64,...``.
        original_link: The full original link text (e.g.
            ``![[file.md#heading|display]]``).  Used by ``keep`` modes.
        source_file_path: Path of the file containing the link.
        visited_files: Set of resolved paths for circular-reference detection.
        depth: Current recursive expansion depth.
        display_text: Display / alt text for ``extract_text`` modes.
        width: Image width hint in pixels.
        height: Image height hint in pixels.
        image_mode: Processing mode for image embeds
            (:class:`~docwen_core.links.EmbeddedImageMode`).
        md_mode: Processing mode for Markdown file embeds
            (:class:`~docwen_core.links.EmbeddedMdMode`).
        on_not_found: Action when a section / block / file is not found.
        on_circular: Action when a circular reference is detected.
        detect_circular: Whether to perform circular-reference detection.
        search_dirs: Directories to search relative to the source file.
        temp_dir: Directory for temporary files (data URI decoding).

    Returns:
        The replacement text, or ``None`` when the file type is unknown
        (caller should decide what to do with unrecognised embeds).
    """
    # Deferred import to avoid circular dependency at module level.
    from docwen_core.links._embed_md import (
        EmbeddedMdMode,
        process_embedded_md_file,
    )

    image_mode = EmbeddedImageMode(image_mode)
    md_mode_enum = EmbeddedMdMode(md_mode)

    # ── 1. Data URI image ─────────────────────────────────────────────
    if is_data_uri_image(link_target):
        temp_path = resolve_data_uri_image_to_temp_file(
            link_target,
            temp_dir=temp_dir,
        )
        if temp_path is not None:
            logger.info("Data URI embedded image → temp file: %s", temp_path)
            return process_embedded_image(
                temp_path,
                original_link,
                mode=image_mode,
                display_text=display_text,
                width=width,
                height=height,
                image_scope=image_scope,
            )
        # Data URI failed to decode — treat as not-found
        logger.warning("Data URI image decode failed; link=%s...", link_target[:60])
        if image_mode == EmbeddedImageMode.KEEP:
            return original_link
        if image_mode == EmbeddedImageMode.EXTRACT_TEXT:
            return display_text or ""
        if image_mode == EmbeddedImageMode.REMOVE:
            return ""
        return display_text or ""

    # Percent-protect wiki/CommonMark backslash escapes before structural
    # query/fragment parsing.  A literal ``\#`` remains part of the filename.
    is_raw_unc = link_target.startswith("\\\\")
    normalized_target = (
        link_target.replace("\\", "/") if is_raw_unc else encode_markdown_destination_escapes(link_target)
    )
    parsed_target = urlsplit(normalized_target)
    is_windows_drive = re.match(r"^[A-Za-z]:[/\\]", normalized_target) is not None
    is_raw_posix_absolute = normalized_target.startswith("/") and not normalized_target.startswith("//")
    if (normalized_target.startswith("//") and not is_raw_unc) or (
        parsed_target.scheme and parsed_target.scheme.lower() != "file" and not is_windows_drive
    ):
        logger.warning("Remote embed targets are not local capabilities: %s", link_target)
        # Remote resources are a distinct unsupported capability, not a
        # missing local file.  This branch intentionally does not consult the
        # local not-found policy: embed mode must leave an explicit,
        # user-visible explanation and must never attempt a network fetch.
        return f"[Remote embed fetching is unsupported: {link_target}]"
    if parsed_target.scheme.lower() == "file":
        # ``urlsplit`` separates structure while still encoded; decode the
        # path exactly once and never pass it through ``parse_anchor`` again.
        file_path_part = _local_path_from_file_uri(normalized_target)
        anchor = unquote(parsed_target.fragment).strip()
        heading = None if not anchor or anchor.startswith("^") else anchor
        block_id = anchor[1:] if anchor.startswith("^") else None
        normalized_target = file_path_part
    else:
        # ── 2. Parse anchor ───────────────────────────────────────────
        file_path_part, heading, block_id = parse_anchor(normalized_target)

    if heading:
        logger.debug("Detected heading embed: %s#%s", file_path_part, heading)
    elif block_id:
        logger.debug("Detected block-id embed: %s#^%s", file_path_part, block_id)

    # ── 3. Resolve file path ──────────────────────────────────────────
    resolved_path = resolve_file_path(
        file_path_part or normalized_target,
        source_file_path,
        search_dirs=list(search_dirs) if search_dirs is not None else None,
        decoded_path=True,
        preserve_absolute=(parsed_target.scheme.lower() == "file" or is_raw_unc or is_raw_posix_absolute),
    )

    if resolved_path is None:
        logger.warning(
            "File not found for embed: %s (source: %s)",
            link_target,
            Path(source_file_path).name,
        )
        if on_not_found in {"ignore", "keep"}:
            return _unresolved_embed_output(original_link, link_target, on_not_found)
        # placeholder
        desc = Path(file_path_part or normalized_target).name
        if heading:
            desc += f"#{heading}"
        elif block_id:
            desc += f"#^{block_id}"
        return f"[File not found: {desc}]"

    logger.debug(
        "Resolved path: %s → %s",
        file_path_part or link_target,
        resolved_path,
    )

    # ── 4. Route by file type ─────────────────────────────────────────
    file_type = get_file_type(resolved_path)

    if file_type == "image":
        logger.info("Dispatching image embed: %s", Path(resolved_path).name)
        return process_embedded_image(
            resolved_path,
            original_link,
            mode=image_mode,
            display_text=display_text,
            width=width,
            height=height,
            image_scope=image_scope,
        )

    if file_type == "markdown":
        logger.info("Dispatching Markdown embed: %s", Path(resolved_path).name)
        if md_mode_enum == EmbeddedMdMode.KEEP:
            # The dispatcher still resolves the target first so the independent
            # not-found policy applies, but an existing file in keep mode must
            # retain the exact user-authored path, fragment, and display text.
            return original_link
        if max_depth is not None and depth >= max_depth:
            return dispatch_error_output(
                LinkErrorKind.MAX_DEPTH,
                on_max_depth,
                Path(file_path_part or link_target).name,
                heading=heading,
                block_id=block_id,
                original_link=original_link,
            )
        return process_embedded_md_file(
            resolved_path,
            source_file_path,
            visited_files,
            depth,
            process_links=process_links,
            heading=heading,
            block_id=block_id,
            display_text=display_text,
            original_link=original_link,
            mode=md_mode_enum,
            on_not_found=on_not_found,
            on_circular=on_circular,
            detect_circular=detect_circular,
            table_safe=table_safe,
        )

    # Unknown content type — caller decides.  Keep the user-authored suffix in
    # the log only as diagnostic context; it was not used for dispatch.
    logger.warning(
        "Unsupported content for embed: %s (declared_ext=%s)",
        resolved_path,
        Path(resolved_path).suffix,
    )
    return None


# ── batch Markdown embed resolution ────────────────────────────────────────


def _parse_embed_display(text: str | None) -> tuple[str | None, int | None, int | None]:
    """Interpret the pipe-separated portion of ``![[target|display]]``.

    Returns ``(display_text, width, height)``.
    """
    if not text:
        return None, None, None
    stripped = text.strip()
    if not stripped:
        return None, None, None
    display_text, width, height = split_alt_text_and_size(stripped)
    if width is not None or height is not None:
        normalized_display = display_text.strip() if display_text else None
        return normalized_display or None, width, height
    # Handle dimension-only display: "200x150" or "200"
    if "x" in stripped:
        parts = stripped.split("x", 1)
        if parts[0].isdigit() and (parts[1].isdigit() or parts[1] == ""):
            return None, int(parts[0]), int(parts[1]) if parts[1].isdigit() else None
    if stripped.isdigit():
        return None, int(stripped), None
    return stripped, None, None


def resolve_embedded_links(
    content: str,
    source_file_path: str,
    *,
    visited_files: set[str] | None = None,
    depth: int = 0,
    max_depth: int = 10,
    image_mode: EmbeddedImageMode | str = EmbeddedImageMode.EMBED,
    md_mode: str = "embed",
    on_not_found: str = "placeholder",
    on_circular: str = "placeholder",
    on_max_depth: str = "placeholder",
    detect_circular: bool = True,
    search_dirs: Sequence[str] | None = None,
    temp_dir: str | None = None,
    table_safe: bool = False,
    image_scope: str | None = None,
    process_links: Callable[..., str] | None = None,
    _table_context_scoped: bool = False,
) -> str:
    """Scan Markdown *content* for ``![[...]]`` wiki-embed links and resolve
    every one via :func:`process_single_embed`.

    This is the top-level entry point that converters call to post-process
    Markdown that may contain unresolved embedded references.  Each match
    is replaced in-place and re-scanning continues until no matches remain
    or *max_depth* is exhausted.

    Args:
        content: Markdown text potentially containing ``![[...]]`` links.
        source_file_path: Path of the file that produced *content*.
        visited_files: Set of resolved absolute paths tracked across
            recursive calls for circular-reference detection.
        depth: Current recursion depth (callers should leave at 0).
        max_depth: Maximum total recursion depth for nested embeds.
        image_mode: Processing mode for embedded image references.
        md_mode: Processing mode for embedded Markdown file references.
        on_not_found: Action when a section / block / file is not found.
        on_circular: Action when a circular reference is detected.
        on_max_depth: Action when max embed depth is reached
            (``"ignore"`` → empty, ``"keep"`` → original link, ``"placeholder"`` → error text).
        detect_circular: Whether to check for circular references.
        search_dirs: Directories to search relative to each source file.
        temp_dir: Directory for temporary files (data URI decoding).

    Returns:
        The Markdown text with all ``![[...]]`` links resolved.
    """
    if visited_files is None:
        visited_files = set()

    if not content:
        return content

    source_file_path = str(Path(source_file_path).resolve())
    visited_files.add(source_file_path)

    # Build a closure that calls *this* function recursively so that nested
    # ``![[...]]`` links inside embedded Markdown files are also expanded.
    def _recurse(
        inner_content: str,
        inner_source: str,
        inner_visited: set[str] | None,
        inner_depth: int,
        **kwargs: object,
    ) -> str:
        # Deferred import avoids coupling the low-level dispatcher to the
        # public orchestrator while still applying its code-protection rule to
        # recursively embedded Markdown.
        from docwen_core.links._non_embed import _map_visible_markdown

        def _resolve_visible(segment: str) -> str:
            if "![[" not in segment:
                return segment
            return resolve_embedded_links(
                segment,
                inner_source,
                visited_files=inner_visited,
                depth=inner_depth,
                max_depth=max_depth,
                image_mode=image_mode,
                md_mode=md_mode,
                on_not_found=on_not_found,
                on_circular=on_circular,
                on_max_depth=on_max_depth,
                detect_circular=detect_circular,
                search_dirs=search_dirs,
                temp_dir=temp_dir,
                table_safe=table_safe,
                image_scope=image_scope,
            )

        return _map_visible_markdown(inner_content, _resolve_visible)

    parts: list[str] = []
    cursor = 0

    # Rebuild from source positions.  Mutating and then calling ``replace``
    # can replace an identical construct inside inline code instead of the
    # later visible match that was actually scanned.
    for match in _WIKI_EMBED_RE.finditer(content):
        original_link = match.group(0)
        link_target = match.group(1).strip()
        raw_display = match.group(2)
        in_table = table_safe if _table_context_scoped else table_safe and _is_table_context(content, match.start())

        # Try the size-specific pattern first for unambiguous width/height
        # extraction (F-H2-024).  When the display portion is purely numeric
        # dimensions (e.g. ``200x150``, ``200``), ``WIKI_EMBED_SIZE_PATTERN``
        # gives us the values directly without heuristic parsing.
        size_match = _WIKI_EMBED_SIZE_RE.match(original_link)
        if size_match is not None:
            display_text = None  # pure dimension — no display text
            width = int(size_match.group(2))
            height = int(size_match.group(3)) if size_match.group(3) else None
        else:
            display_text, width, height = _parse_embed_display(raw_display)

        logger.debug("Found wiki embed: %s", original_link)

        replacement = process_single_embed(
            link_target=link_target,
            original_link=original_link,
            source_file_path=source_file_path,
            visited_files=visited_files,
            depth=depth,
            display_text=display_text,
            width=width,
            height=height,
            image_mode=image_mode,
            md_mode=md_mode,
            on_not_found=on_not_found,
            on_circular=on_circular,
            on_max_depth=on_max_depth,
            detect_circular=detect_circular,
            max_depth=max_depth,
            search_dirs=search_dirs,
            temp_dir=temp_dir,
            table_safe=in_table,
            image_scope=image_scope,
            process_links=process_links if process_links is not None else _recurse,
        )

        parts.append(content[cursor : match.start()])
        parts.append(original_link if replacement is None else replacement)
        cursor = match.end()

    parts.append(content[cursor:])
    return "".join(parts)
