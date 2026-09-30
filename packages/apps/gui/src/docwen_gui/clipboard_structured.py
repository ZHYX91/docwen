"""Offline HTML/CF_HTML projection into the recursive clipboard-document model."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from html.parser import HTMLParser
from typing import Iterable

from docwen_core.models.clipboard_document import (
    MAX_CLIPBOARD_BLOCKS,
    MAX_CLIPBOARD_CELLS,
    MAX_CLIPBOARD_DEPTH,
    MAX_CLIPBOARD_TABLES,
    ClipboardBlock,
    ClipboardDocument,
    ClipboardDocumentError,
    ClipboardHardBreak,
    ClipboardImageRef,
    ClipboardParagraph,
    ClipboardTable,
    ClipboardTableCell,
    ClipboardText,
    clipboard_document_to_bytes,
)

_CF_HTML_FIELDS = ("StartHTML", "EndHTML", "StartFragment", "EndFragment")
_SKIP_TAGS = frozenset({"head", "script", "style", "noscript", "template"})
_PARAGRAPH_TAGS = frozenset(
    {
        "address",
        "article",
        "aside",
        "blockquote",
        "div",
        "footer",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "header",
        "li",
        "main",
        "nav",
        "p",
        "pre",
        "section",
    }
)
_ROW_GROUP_TAGS = frozenset({"thead", "tbody", "tfoot"})
_MAX_HTML_BYTES = 8 * 1024 * 1024
_MAX_HTML_NODES = 300_000
_MAX_HTML_DEPTH = 32


@dataclass(slots=True)
class _Node:
    tag: str
    attrs: dict[str, str]
    children: list["_Node | str"] = field(default_factory=list)


class _TreeBuilder(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = _Node("document", {})
        self._stack: list[_Node] = [self.root]
        self._count = 1

    def _append_node(self, tag: str, attrs: list[tuple[str, str | None]], *, push: bool) -> None:
        self._count += 1
        if self._count > _MAX_HTML_NODES:
            raise ClipboardDocumentError("clipboard.html_budget_exceeded", "Clipboard HTML node budget exceeded.")
        if len(self._stack) >= _MAX_HTML_DEPTH:
            raise ClipboardDocumentError("clipboard.html_depth_exceeded", "Clipboard HTML nesting depth exceeded.")
        node = _Node(tag.lower(), {str(key).lower(): str(value or "") for key, value in attrs})
        self._stack[-1].children.append(node)
        if push:
            self._stack.append(node)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self._append_node(tag, attrs, push=tag.lower() not in {"br", "img", "meta", "link", "hr", "input"})

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self._append_node(tag, attrs, push=False)

    def handle_endtag(self, tag: str) -> None:
        normalized = tag.lower()
        for index in range(len(self._stack) - 1, 0, -1):
            if self._stack[index].tag == normalized:
                del self._stack[index:]
                return

    def handle_data(self, data: str) -> None:
        if not data:
            return
        self._count += 1
        if self._count > _MAX_HTML_NODES:
            raise ClipboardDocumentError("clipboard.html_budget_exceeded", "Clipboard HTML node budget exceeded.")
        self._stack[-1].children.append(data)


def extract_cf_html_fragment(payload: bytes) -> bytes:
    """Return a strict CF_HTML fragment or the original UTF-8 HTML bytes.

    SourceURL and base declarations are deliberately not interpreted as read
    authority. Offsets are byte offsets into the exact captured payload.
    """

    if not isinstance(payload, bytes) or not payload or len(payload) > _MAX_HTML_BYTES:
        raise ClipboardDocumentError("clipboard.html_size_invalid", "Clipboard HTML payload size is invalid.")
    header_probe = payload[: min(len(payload), 8192)]
    if b"StartHTML:" not in header_probe:
        payload.decode("utf-8-sig", errors="strict")
        return payload

    fields: dict[str, int] = {}
    for name in _CF_HTML_FIELDS:
        match = re.search(rb"(?m)^" + name.encode("ascii") + rb":([0-9]{1,10})[ \t]*\r?$", header_probe)
        if match is None:
            raise ClipboardDocumentError("clipboard.cf_html_header_invalid", f"CF_HTML is missing {name}.")
        fields[name] = int(match.group(1))
    start_html = fields["StartHTML"]
    end_html = fields["EndHTML"]
    start_fragment = fields["StartFragment"]
    end_fragment = fields["EndFragment"]
    if not (0 <= start_html <= start_fragment <= end_fragment <= end_html <= len(payload)):
        raise ClipboardDocumentError("clipboard.cf_html_offsets_invalid", "CF_HTML byte offsets are inconsistent.")
    html_bytes = payload[start_html:end_html]
    fragment = payload[start_fragment:end_fragment]
    html_bytes.decode("utf-8-sig", errors="strict")
    fragment.decode("utf-8-sig", errors="strict")
    return fragment


def _decode_html(payload: bytes) -> str:
    fragment = extract_cf_html_fragment(payload)
    try:
        return fragment.decode("utf-8-sig", errors="strict")
    except UnicodeDecodeError as exc:
        raise ClipboardDocumentError("clipboard.html_encoding_invalid", "Clipboard HTML must be UTF-8.") from exc


def _contains_table(node: _Node) -> bool:
    for child in node.children:
        if isinstance(child, _Node):
            if child.tag == "table" or _contains_table(child):
                return True
    return False


def html_contains_table(payload: bytes) -> bool:
    parser = _TreeBuilder()
    parser.feed(_decode_html(payload))
    parser.close()
    return _contains_table(parser.root)


def _inline_nodes(items: Iterable[_Node | str]) -> tuple:
    result: list = []

    def append_text(value: str) -> None:
        if value:
            if result and isinstance(result[-1], ClipboardText):
                previous = result.pop()
                result.append(ClipboardText(previous.value + value))
            else:
                result.append(ClipboardText(value))

    def walk(item: _Node | str) -> None:
        if isinstance(item, str):
            append_text(item)
            return
        if item.tag in _SKIP_TAGS:
            return
        if item.tag == "br":
            result.append(ClipboardHardBreak())
            return
        if item.tag == "img":
            result.append(
                ClipboardImageRef(
                    None,
                    item.attrs.get("alt", ""),
                    "clipboard_resource_unavailable",
                )
            )
            return
        if item.tag == "table":
            raise ClipboardDocumentError(
                "clipboard.table_inline_invalid",
                "A table cannot be represented as an inline paragraph node.",
            )
        for child in item.children:
            walk(child)

    for item in items:
        walk(item)
    return tuple(result)


def _has_meaningful_inline(items: list[_Node | str]) -> bool:
    text_parts: list[str] = []

    def walk(item: _Node | str) -> bool:
        if isinstance(item, str):
            text_parts.append(item)
            return False
        if item.tag in _SKIP_TAGS:
            return False
        if item.tag in {"img", "br"}:
            return True
        if item.tag == "table":
            return True
        return any(walk(child) for child in item.children)

    structural = any(walk(item) for item in items)
    text = "".join(text_parts)
    if structural:
        return True
    if not text:
        return False
    if "\n" in text or "\r" in text or "\t" in text:
        return bool(text.strip(" \t\r\n"))
    return True


def _blocks_from_children(
    children: Iterable[_Node | str],
    *,
    counters: dict[str, int],
    depth: int,
) -> tuple[ClipboardBlock, ...]:
    if depth > MAX_CLIPBOARD_DEPTH:
        raise ClipboardDocumentError("clipboard.depth_exceeded", "Clipboard nesting depth exceeded.")
    blocks: list[ClipboardBlock] = []
    inline_buffer: list[_Node | str] = []

    def append_block(block: ClipboardBlock) -> None:
        counters["blocks"] += 1
        if counters["blocks"] > MAX_CLIPBOARD_BLOCKS:
            raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard block budget exceeded.")
        blocks.append(block)

    def flush_inline(*, force: bool = False) -> None:
        if not inline_buffer:
            if force:
                append_block(ClipboardParagraph(()))
            return
        if force or _has_meaningful_inline(inline_buffer):
            append_block(ClipboardParagraph(_inline_nodes(inline_buffer)))
        inline_buffer.clear()

    for child in children:
        if isinstance(child, str):
            inline_buffer.append(child)
            continue
        if child.tag in _SKIP_TAGS:
            flush_inline()
            continue
        if child.tag == "table":
            flush_inline()
            append_block(_parse_table(child, counters=counters, depth=depth))
            continue
        if child.tag in _PARAGRAPH_TAGS:
            flush_inline()
            nested_table = any(isinstance(item, _Node) and item.tag == "table" for item in child.children)
            if nested_table:
                for block in _blocks_from_children(child.children, counters=counters, depth=depth + 1):
                    blocks.append(block)
            else:
                append_block(ClipboardParagraph(_inline_nodes(child.children)))
            continue
        inline_buffer.append(child)
    flush_inline()
    return tuple(blocks)


@dataclass(frozen=True, slots=True)
class _Row:
    node: _Node
    group_end: int


def _table_rows(table: _Node) -> list[_Row]:
    grouped: list[list[_Node]] = []
    direct: list[_Node] = []

    def flush_direct() -> None:
        nonlocal direct
        if direct:
            grouped.append(direct)
            direct = []

    for child in table.children:
        if not isinstance(child, _Node):
            continue
        if child.tag == "tr":
            direct.append(child)
            continue
        if child.tag in _ROW_GROUP_TAGS:
            flush_direct()
            rows = [item for item in child.children if isinstance(item, _Node) and item.tag == "tr"]
            if rows:
                grouped.append(rows)
    flush_direct()

    rows: list[_Row] = []
    offset = 0
    for group in grouped:
        group_end = offset + len(group)
        rows.extend(_Row(row, group_end) for row in group)
        offset = group_end
    return rows


def _span(raw: str | None, *, allow_zero: bool, where: str) -> int:
    if raw is None or raw == "":
        return 1
    if not re.fullmatch(r"[0-9]+", raw):
        raise ClipboardDocumentError("clipboard.table_span_invalid", f"{where} is not a decimal span.")
    value = int(raw)
    if value == 0 and allow_zero:
        return 0
    if value <= 0 or value > MAX_CLIPBOARD_CELLS:
        raise ClipboardDocumentError("clipboard.table_span_invalid", f"{where} is out of range.")
    return value


def _cell_nodes(row: _Node) -> list[_Node]:
    return [child for child in row.children if isinstance(child, _Node) and child.tag in {"td", "th"}]


def _parse_table(table: _Node, *, counters: dict[str, int], depth: int) -> ClipboardTable:
    counters["tables"] += 1
    if counters["tables"] > MAX_CLIPBOARD_TABLES:
        raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard table budget exceeded.")
    rows = _table_rows(table)
    if not rows:
        raise ClipboardDocumentError("clipboard.table_rows_invalid", "Clipboard table has no rows.")

    coverage: dict[tuple[int, int], ClipboardTableCell] = {}
    anchors: list[ClipboardTableCell] = []
    max_column = 0
    for row_index, row_info in enumerate(rows):
        column = 0
        for cell_index, node in enumerate(_cell_nodes(row_info.node)):
            while (row_index, column) in coverage:
                column += 1
            column_span = _span(node.attrs.get("colspan"), allow_zero=False, where="colspan")
            row_span = _span(node.attrs.get("rowspan"), allow_zero=True, where="rowspan")
            if row_span == 0:
                row_span = row_info.group_end - row_index
            if row_index + row_span > row_info.group_end:
                raise ClipboardDocumentError(
                    "clipboard.table_span_out_of_group",
                    "A row span exceeds its HTML row group.",
                )
            if column + column_span > MAX_CLIPBOARD_CELLS:
                raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard table width budget exceeded.")
            blocks = _blocks_from_children(node.children, counters=counters, depth=depth + 1)
            headers = tuple(part for part in node.attrs.get("headers", "").split() if part)
            cell = ClipboardTableCell(
                row=row_index,
                column=column,
                row_span=row_span,
                column_span=column_span,
                blocks=blocks,
                header=node.tag == "th",
                scope=node.attrs.get("scope", "").lower(),
                headers=headers,
                html_id=node.attrs.get("id", ""),
            )
            for covered_row in range(row_index, row_index + row_span):
                for covered_column in range(column, column + column_span):
                    key = (covered_row, covered_column)
                    if key in coverage:
                        raise ClipboardDocumentError(
                            "clipboard.table_overlap",
                            f"Table cell overlap at row {covered_row}, column {covered_column}.",
                        )
                    coverage[key] = cell
            anchors.append(cell)
            counters["cells"] += 1
            if counters["cells"] > MAX_CLIPBOARD_CELLS:
                raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard cell budget exceeded.")
            column += column_span
            max_column = max(max_column, column)

    if not max_column or len(rows) * max_column > MAX_CLIPBOARD_CELLS:
        raise ClipboardDocumentError("clipboard.budget_exceeded", "Clipboard table grid budget exceeded.")
    missing = [
        (row, column)
        for row in range(len(rows))
        for column in range(max_column)
        if (row, column) not in coverage
    ]
    if missing:
        row, column = missing[0]
        raise ClipboardDocumentError(
            "clipboard.table_grid_incomplete",
            f"Clipboard table grid is incomplete at row {row}, column {column}.",
        )
    return ClipboardTable(len(rows), max_column, tuple(anchors))


def project_structured_clipboard_html(payload: bytes) -> ClipboardDocument:
    """Parse one captured HTML/CF_HTML byte snapshot without external reads."""

    parser = _TreeBuilder()
    parser.feed(_decode_html(payload))
    parser.close()
    counters = {"blocks": 0, "tables": 0, "cells": 0}
    document = ClipboardDocument(blocks=_blocks_from_children(parser.root.children, counters=counters, depth=1))
    # Core strict validation and deterministic serialization are the final
    # producer check before any managed snapshot can be published.
    clipboard_document_to_bytes(document)
    return document


__all__ = ["extract_cf_html_fragment", "html_contains_table", "project_structured_clipboard_html"]
