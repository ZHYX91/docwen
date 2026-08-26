"""Request-scoped, target-aware Markdown link processing."""

from __future__ import annotations

import logging
import re
from collections.abc import Sequence
from pathlib import Path
from urllib.parse import unquote

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links._embed_dispatch import (
    _is_table_context,
    process_single_embed,
    resolve_embedded_links,
)
from docwen_core.links._embed_image import EmbeddedImageMode
from docwen_core.links._markdown_inline import (
    _contains_active_protected_token,
    encode_markdown_destination_escapes,
    escape_markdown_source_literal,
    escape_unescaped_pipes,
    parse_inline_link,
    parse_markdown_destination,
)
from docwen_core.links._non_embed import (
    _bare_url_end,
    _map_visible_markdown,
    _process_non_embed_links,
    _split_fenced_code_blocks,
    _split_inline_code_spans,
)
from docwen_core.links._patterns import WIKI_EMBED_PATTERN

logger = logging.getLogger(__name__)

_MARKDOWN_IMAGE_SIZE_RE = re.compile(r"^=\s*(\d*)x(\d*)$")
_IMAGE_PLACEHOLDER_RE = re.compile(r"\{\{IMAGE:([^{}\r\n]+)\}\}")
_WIKI_EMBED_RE = re.compile(WIKI_EMBED_PATTERN)
_WIKI_LINK_RE = re.compile(r"(?<!!)" + WIKI_EMBED_PATTERN[1:])


def _image_placeholder_re(image_scope: str | None) -> re.Pattern[str]:
    if image_scope is None:
        return _IMAGE_PLACEHOLDER_RE
    return re.compile(rf"\{{\{{IMAGE@{re.escape(image_scope)}:([^{{}}\r\n]+)\}}\}}")


def _parse_markdown_image_target(target: str) -> tuple[str, int | None, int | None]:
    """Split a Markdown image destination from its optional ``=WxH`` hint."""
    parsed = parse_markdown_destination(target, allow_image_size=True)
    if parsed is None:
        return encode_markdown_destination_escapes(target.strip(" \t\r\n")), None, None
    match = _MARKDOWN_IMAGE_SIZE_RE.fullmatch(parsed.suffix.strip())
    if match is None:
        image_target = parsed.destination
        width = None
        height = None
    else:
        image_target = parsed.destination
        width_text = match.group(1)
        height_text = match.group(2)
        width = int(width_text) if width_text else None
        height = int(height_text) if height_text else None
    return encode_markdown_destination_escapes(image_target), width, height


def _escape_table_image_placeholder_pipes(
    text: str,
    image_scope: str | None,
) -> str:
    """Keep sized image placeholders inside one Markdown table cell."""

    def replace(match: re.Match[str]) -> str:
        if not _is_table_context(text, match.start()):
            return match.group(0)
        payload = re.sub(r"(?<!\\)\|", r"\\|", match.group(1))
        marker = "IMAGE" if image_scope is None else f"IMAGE@{image_scope}"
        return f"{{{{{marker}:{payload}}}}}"

    return _image_placeholder_re(image_scope).sub(replace, text)


def _replace_markdown_images(
    segment: str,
    *,
    mode: str,
    target_format: str,
    source_file_path: str,
    search_dirs: Sequence[str] | None,
    on_not_found: str,
    temp_dir: str | None,
    table_safe: bool,
    image_scope: str | None,
) -> str:
    """Apply a mode to standard ``![alt](target)`` image syntax."""
    normalized_mode = EmbeddedImageMode(mode)
    parts: list[str] = []
    cursor = 0
    index = 0

    while index < len(segment):
        construct = parse_inline_link(segment, index, image=True)
        if construct is None:
            index += 1
            continue
        original = segment[index : construct.end]
        in_table = table_safe and _is_table_context(segment, index)
        alt = construct.label.strip()
        target = construct.target.strip(" \t\r\n")
        if normalized_mode is EmbeddedImageMode.EXTRACT_TEXT:
            image_target, _, _ = _parse_markdown_image_target(target)
            replacement = alt or Path(unquote(image_target)).name
        elif normalized_mode is EmbeddedImageMode.REMOVE:
            replacement = ""
        elif normalized_mode is EmbeddedImageMode.KEEP and target_format == "docx":
            replacement = escape_markdown_source_literal(original)
        elif normalized_mode is EmbeddedImageMode.EMBED:
            image_target, width, height = _parse_markdown_image_target(target)
            embedded = process_single_embed(
                link_target=image_target,
                original_link=original,
                source_file_path=source_file_path,
                visited_files=None,
                depth=0,
                display_text=alt,
                width=width,
                height=height,
                image_mode=EmbeddedImageMode.EMBED,
                md_mode="keep",
                on_not_found=on_not_found,
                search_dirs=search_dirs,
                temp_dir=temp_dir,
                table_safe=table_safe,
                image_scope=image_scope,
            )
            if embedded is None:
                replacement = original
            elif target_format == "docx" and embedded == original:
                replacement = escape_markdown_source_literal(original)
            else:
                replacement = embedded
        else:
            replacement = original
        if in_table:
            replacement = escape_unescaped_pipes(replacement)
        parts.append(segment[cursor:index])
        parts.append(replacement)
        index = construct.end
        cursor = index

    parts.append(segment[cursor:])
    return "".join(parts)


def _markdown_construct_end(text: str, start: int) -> int | None:
    """Return the end of a Markdown link/image beginning at *start*."""
    construct = parse_inline_link(text, start, image=text.startswith("![", start))
    return None if construct is None else construct.end


def _escaped_markdown_construct_end(text: str, start: int) -> int | None:
    """Return the end of the canonical escaped literal emitted for DOCX."""
    marker = "!\\[" if text.startswith("!\\[", start) else "\\["
    if not text.startswith(marker, start):
        return None
    close_bracket = text.find(r"\]\(", start + len(marker))
    if close_bracket == -1:
        return None
    close_target = text.find(r"\)", close_bracket + 4)
    return None if close_target == -1 else close_target + 2


def _wiki_construct_end(text: str, start: int) -> int | None:
    """Return the end of a wiki link/embed beginning at *start*."""
    if text.startswith("![[", start):
        close = text.find("]]", start + 3)
    elif text.startswith("[[", start):
        close = text.find("]]", start + 2)
    else:
        return None
    return None if close == -1 else close + 2


def _auto_link_bare_urls_in_segment(segment: str) -> str:
    """Turn safe bare HTTP(S) URLs into explicit Markdown links."""
    parts: list[str] = []
    cursor = 0
    index = 0

    while index < len(segment):
        construct_end = _markdown_construct_end(segment, index)
        if construct_end is None:
            construct_end = _escaped_markdown_construct_end(segment, index)
        if construct_end is None:
            construct_end = _wiki_construct_end(segment, index)
        if construct_end is not None:
            index = construct_end
            continue

        if segment[index] == "<":
            tag_end = segment.find(">", index + 1)
            if tag_end != -1:
                index = tag_end + 1
                continue

        url_end = _bare_url_end(segment, index)
        if url_end is None:
            index += 1
            continue
        end = url_end
        while end < len(segment) and not segment[end].isspace() and segment[end] not in "<>":
            end += 1
        url = segment[index:url_end]
        trailing = segment[url_end:end]
        parts.append(segment[cursor:index])
        parts.append(f"<{url}>{trailing}")
        cursor = end
        index = end

    parts.append(segment[cursor:])
    return "".join(parts)


def _restore_expansions_with_ranges(
    text: str,
    expansions: dict[str, str],
) -> tuple[str, tuple[tuple[int, int], ...]]:
    """Restore embed tokens and retain their exact output provenance ranges."""
    if not expansions:
        return text, ()
    token_re = re.compile("|".join(re.escape(token) for token in sorted(expansions, key=len, reverse=True)))
    parts: list[str] = []
    ranges: list[tuple[int, int]] = []
    cursor = 0
    output_length = 0
    for match in token_re.finditer(text):
        prefix = text[cursor : match.start()]
        parts.append(prefix)
        output_length += len(prefix)
        expansion = expansions[match.group(0)]
        start = output_length
        parts.append(expansion)
        output_length += len(expansion)
        ranges.append((start, output_length))
        cursor = match.end()
    parts.append(text[cursor:])
    return "".join(parts), tuple(ranges)


def _visible_markdown_ranges(
    text: str,
    *,
    protect_bare_urls: bool,
) -> list[tuple[int, int]]:
    """Return source-position ranges outside renderer-protected atoms."""
    ranges: list[tuple[int, int]] = []
    block_offset = 0
    for block, block_protected in _split_fenced_code_blocks(text):
        if not block_protected:
            inline_offset = block_offset
            for segment, inline_protected in _split_inline_code_spans(
                block,
                protect_bare_urls=protect_bare_urls,
            ):
                if segment and not inline_protected:
                    ranges.append((inline_offset, inline_offset + len(segment)))
                inline_offset += len(segment)
        block_offset += len(block)
    return ranges


def _crosses_expansion_boundary(
    start: int,
    end: int,
    expansion_ranges: tuple[tuple[int, int], ...],
) -> bool:
    return any(
        start < expansion_start < end or start < expansion_end < end
        for expansion_start, expansion_end in expansion_ranges
    )


def _cross_boundary_construct_spans(
    text: str,
    expansion_ranges: tuple[tuple[int, int], ...],
    *,
    auto_link_bare_url: bool,
) -> list[tuple[int, int]]:
    """Find only newly composed constructs that cross an embed boundary."""
    spans: list[tuple[int, int]] = []
    for range_start, range_end in _visible_markdown_ranges(
        text,
        protect_bare_urls=not auto_link_bare_url,
    ):
        index = range_start
        while index < range_end:
            construct = parse_inline_link(text, index, image=True)
            if construct is None:
                construct = parse_inline_link(text, index, image=False)
            end = construct.end if construct is not None else None
            if end is None:
                wiki_match = _WIKI_EMBED_RE.match(text, index)
                if wiki_match is None:
                    wiki_match = _WIKI_LINK_RE.match(text, index)
                end = wiki_match.end() if wiki_match is not None else None
            if end is None and auto_link_bare_url:
                end = _bare_url_end(text, index)
            if end is None or end > range_end:
                index += 1
                continue
            if _crosses_expansion_boundary(index, end, expansion_ranges):
                spans.append((index, end))
            index = end
    return spans


def process_markdown_links(
    text: str,
    source_file_path: str,
    *,
    link_config: LinkRuntimeConfig,
    target_format: str = "md",
    temp_dir: str | None = None,
    table_safe: bool = False,
    image_scope: str | None = None,
    _visited_files: set[str] | None = None,
    _depth: int = 0,
    _canonicalize_local_docx_targets: bool = False,
    _boundary_rescan_remaining: int = 1,
) -> str:
    """Process Markdown links using one immutable request policy."""
    if not text:
        return text

    resolved_max_depth = link_config.max_depth
    resolved_image_mode = link_config.embed_wiki_image_mode
    resolved_markdown_image_mode = link_config.embed_markdown_image_mode
    resolved_md_mode = link_config.embed_md_file_mode
    resolved_wiki_mode = link_config.non_embed_wiki_mode
    resolved_markdown_mode = link_config.non_embed_markdown_mode
    resolved_auto_link = link_config.auto_link_bare_url
    resolved_search_dirs = link_config.search_dirs
    resolved_not_found = link_config.file_not_found_mode
    resolved_circular = link_config.circular_reference_mode
    resolved_max_depth_mode = link_config.max_depth_reached_mode
    resolved_detect_circular = link_config.detect_circular
    normalized_target = target_format.strip().lower()
    resolved_source = str(Path(source_file_path).resolve())

    logger.info(
        "process_markdown_links | source=%s | target=%s | depth<=%d",
        Path(resolved_source).name,
        normalized_target,
        resolved_max_depth,
    )

    protected_expansions: dict[str, str] = {}

    def _protect_expansion(expanded: str) -> str:
        index = len(protected_expansions)
        token = f"\x00DOCWEN_EMBED_{index}\x00"
        while (
            token in text
            or token in expanded
            or token in protected_expansions
            or any(token in value for value in protected_expansions.values())
        ):
            index += 1
            token = f"\x00DOCWEN_EMBED_{index}\x00"
        protected_expansions[token] = expanded
        return token

    def _process_child(
        inner_content: str,
        inner_source: str,
        inner_visited: set[str] | None,
        inner_depth: int,
        **kwargs: object,
    ) -> str:
        expanded = process_markdown_links(
            inner_content,
            inner_source,
            link_config=link_config,
            target_format=normalized_target,
            temp_dir=temp_dir,
            table_safe=table_safe,
            image_scope=image_scope,
            _visited_files=inner_visited,
            _depth=inner_depth,
            _canonicalize_local_docx_targets=True,
        )
        if kwargs.get("table_safe"):
            from docwen_core.links._embed_md import _make_table_safe

            expanded = _make_table_safe(expanded)
        return _protect_expansion(expanded)

    def _resolve_embeds(segment: str) -> str:
        if "![[" not in segment:
            return segment
        parts: list[str] = []
        cursor = 0
        index = 0
        while index < len(segment):
            construct = parse_inline_link(segment, index, image=True)
            if construct is None:
                construct = parse_inline_link(segment, index, image=False)
            if construct is not None:
                index = construct.end
                continue
            match = _WIKI_EMBED_RE.match(segment, index)
            if match is None:
                index += 1
                continue
            if _contains_active_protected_token(match.group(0)):
                index = match.end()
                continue
            parts.append(segment[cursor:index])
            original = match.group(0)
            in_table = table_safe and _is_table_context(segment, index)
            replacement = resolve_embedded_links(
                original,
                resolved_source,
                visited_files=_visited_files,
                depth=_depth,
                max_depth=resolved_max_depth,
                image_mode=resolved_image_mode,
                md_mode=resolved_md_mode,
                on_not_found=resolved_not_found,
                on_circular=resolved_circular,
                on_max_depth=resolved_max_depth_mode,
                detect_circular=resolved_detect_circular,
                search_dirs=resolved_search_dirs,
                temp_dir=temp_dir,
                table_safe=in_table,
                image_scope=image_scope,
                process_links=_process_child,
                _table_context_scoped=True,
            )
            if normalized_target == "docx" and replacement == original:
                replacement = escape_markdown_source_literal(original)
            if in_table:
                replacement = escape_unescaped_pipes(replacement)
            parts.append(replacement)
            index = match.end()
            cursor = index
        parts.append(segment[cursor:])
        return "".join(parts)

    initial_text = text
    if resolved_auto_link and normalized_target == "docx":
        initial_text = _map_visible_markdown(
            initial_text,
            _auto_link_bare_urls_in_segment,
            protect_bare_urls=False,
        )

    result = _map_visible_markdown(initial_text, _resolve_embeds)
    result = _map_visible_markdown(
        result,
        lambda segment: _replace_markdown_images(
            segment,
            mode=str(resolved_markdown_image_mode),
            target_format=normalized_target,
            source_file_path=resolved_source,
            search_dirs=resolved_search_dirs,
            on_not_found=resolved_not_found,
            temp_dir=temp_dir,
            table_safe=table_safe,
            image_scope=image_scope,
        ),
    )
    result = _process_non_embed_links(
        result,
        source_file_path=resolved_source,
        wiki_mode=resolved_wiki_mode,
        markdown_mode=resolved_markdown_mode,
        search_dirs=resolved_search_dirs,
        target_format=normalized_target,
        on_not_found=resolved_not_found,
        canonicalize_local_docx_targets=_canonicalize_local_docx_targets,
        table_safe=table_safe,
    )
    if table_safe:
        result = _escape_table_image_placeholder_pipes(result, image_scope)
    result, expansion_ranges = _restore_expansions_with_ranges(
        result,
        protected_expansions,
    )

    if expansion_ranges and _boundary_rescan_remaining > 0:
        boundary_points = {point for expansion_range in expansion_ranges for point in expansion_range}
        seen_states: set[tuple[str, tuple[int, ...]]] = set()
        while True:
            boundary_ranges = tuple((point, point) for point in sorted(boundary_points))
            spans = _cross_boundary_construct_spans(
                result,
                boundary_ranges,
                auto_link_bare_url=(resolved_auto_link and normalized_target == "docx"),
            )
            if not spans:
                break
            state = (result, tuple(sorted(boundary_points)))
            if state in seen_states:
                raise RuntimeError("Cross-boundary link processing did not converge")
            seen_states.add(state)

            start, end = spans[0]
            span_in_table = table_safe and _is_table_context(result, start)
            replacement = process_markdown_links(
                result[start:end],
                resolved_source,
                link_config=link_config,
                target_format=normalized_target,
                temp_dir=temp_dir,
                table_safe=False,
                image_scope=image_scope,
                _visited_files=_visited_files,
                _depth=_depth,
                _canonicalize_local_docx_targets=(_canonicalize_local_docx_targets),
                _boundary_rescan_remaining=0,
            )
            if span_in_table:
                replacement = escape_unescaped_pipes(replacement)

            delta = len(replacement) - (end - start)
            shifted_points = {
                point if point < start else point + delta for point in boundary_points if point < start or point > end
            }
            shifted_points.update((start, start + len(replacement)))
            boundary_points = shifted_points
            result = result[:start] + replacement + result[end:]

    logger.info("process_markdown_links complete")
    return result
