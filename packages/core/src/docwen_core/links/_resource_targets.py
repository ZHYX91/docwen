"""Source-preserving relocation of producer-declared Markdown resources."""

from __future__ import annotations

import re
from collections.abc import Callable

from docwen_core.links._block_context import html_block_mode
from docwen_core.links._markdown_inline import _INLINE_HTML_RE, parse_inline_link, parse_markdown_destination
from docwen_core.links._patterns import WIKI_EMBED_PATTERN
from docwen_core.links._source_document import MarkdownBlockLine, markdown_source_context_lines
from docwen_core.links._source_ownership import markdown_source_owners

_WIKI_LINK = re.compile(r"!?" + WIKI_EMBED_PATTERN[1:])
_HTML_TAG_NAME = re.compile(r"</?([A-Za-z][A-Za-z0-9-]*)")
_HTML_ATTRIBUTE = re.compile(
    r"\s+(?P<name>[A-Za-z_:][A-Za-z0-9_.:-]*)"
    r"(?:\s*=\s*(?:\"(?P<double>[^\"]*)\"|'(?P<single>[^']*)'|(?P<bare>[^\s\"'=<>`]+)))?"
)
_RAW_HTML_TAGS = frozenset({"pre", "script", "style", "textarea"})
_ASCII_WHITESPACE = " \t\r\n"


def rewrite_markdown_resource_targets(source: str, replace_target: Callable[[str, str], str]) -> str:
    """Rewrite recognized destination spans while retaining all other source.

    ``replace_target`` receives an authored destination plus its syntax kind
    (``markdown``, ``wiki`` or ``html``). Full-source ownership is authoritative
    for YAML, container exits, fences, code spans, comments and link metadata.
    HTML attributes are eligible only inside actual tags, never inside samples
    or another attribute's quoted value. HTML raw-text blocks stay literal.
    """
    owners = markdown_source_owners(source, full_document=True)
    edits: list[tuple[int, int, str]] = []
    context_by_start: dict[int, MarkdownBlockLine] | None = None

    def add(start: int, end: int, syntax: str) -> None:
        original = source[start:end]
        replacement = replace_target(original, syntax)
        if replacement != original:
            edits.append((start, end, replacement))

    def html_tags(start: int, end: int) -> None:
        raw_tag: str | None = None
        for match in _INLINE_HTML_RE.finditer(source, start, end):
            tag = _HTML_TAG_NAME.match(match.group())
            if tag is None:
                continue
            name = tag.group(1).lower()
            closing = match.group().startswith("</")
            if raw_tag is not None:
                if closing and name == raw_tag:
                    raw_tag = None
                continue
            if name in _RAW_HTML_TAGS and not closing:
                raw_tag = name
                continue
            if closing:
                continue
            for attribute in _HTML_ATTRIBUTE.finditer(source, match.start() + tag.end(), match.end()):
                if attribute.group("name").lower() not in {"src", "href"}:
                    continue
                for value_kind in ("double", "single", "bare"):
                    if attribute.group(value_kind) is not None:
                        add(*attribute.span(value_kind), "html")
                        break

    for owner in owners:
        if owner.kind == "block":
            # Resolve block type through the same context inventory. In an
            # HTML block only HTML tags are interpreted; apparent Markdown
            # inside the block remains exactly authored.
            if "<" not in source[owner.start : owner.end]:
                continue
            if context_by_start is None:
                context_by_start = {}
                offset = 0
                for line, context in zip(
                    source.splitlines(keepends=True), markdown_source_context_lines(source), strict=True
                ):
                    context_by_start[offset] = context
                    offset += len(line)
            context = context_by_start.get(owner.start)
            if context is None or context.block_kind != "html_block":
                continue
            mode = html_block_mode(context.content, allow_type7=True)
            if mode is not None and mode[0] is None:
                html_tags(owner.start, owner.end)
        elif owner.kind == "literal":
            literal = source[owner.start : owner.end]
            if _INLINE_HTML_RE.fullmatch(literal) is not None:
                html_tags(owner.start, owner.end)
        elif owner.kind == "wikilink":
            match = _WIKI_LINK.match(source, owner.start, owner.end)
            if match is not None:
                start, end = match.span(1)
                target = source[start:end]
                leading = len(target) - len(target.lstrip(_ASCII_WHITESPACE))
                trailing = len(target) - len(target.rstrip(_ASCII_WHITESPACE))
                add(start + leading, end - trailing, "wiki")
        elif owner.kind == "link_metadata":
            construct = parse_inline_link(source, owner.start, image=True)
            if construct is None:
                construct = parse_inline_link(source, owner.start, image=False)
            if construct is None:
                continue
            destination = parse_markdown_destination(construct.target, allow_image_size=construct.is_image)
            if destination is None:
                continue
            leading = len(construct.target) - len(construct.target.lstrip(_ASCII_WHITESPACE))
            angle = construct.target.lstrip(_ASCII_WHITESPACE).startswith("<")
            start = owner.start + (2 if construct.is_image else 1) + len(construct.label) + 2 + leading + int(angle)
            add(start, start + len(destination.destination), "markdown")

    output: list[str] = []
    previous_end = 0
    for start, end, replacement in sorted(edits):
        if start < previous_end:
            raise ValueError("overlapping Markdown resource destination ownership")
        output.extend((source[previous_end:start], replacement))
        previous_end = end
    output.append(source[previous_end:])
    return "".join(output)
