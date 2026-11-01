"""Review comments anchored only to request-owned, newly inserted numbering."""

from __future__ import annotations

import re
import secrets
from typing import Any

from docwen_core.docx_comments import append_docx_comment
from docwen_core.links import markdown_heading_ranges

NUMBERING_REVIEW_TEXT = "此序号由软件自动添加，请核对标题级别及编号是否正确。"


class NumberingProvenance:
    """Carry insertion identities through links/embeds, before the final parse.

    Markers are removed before the final Markdown parse, preserving custom
    prefix formatting and source/carrier identity. Final source offsets identify
    the original numbered headings even when embeds add identical headings.
    """

    def __init__(self, source: str) -> None:
        self._scope = f"DWNUM{secrets.token_hex(24)}X"
        while self._scope in source:
            self._scope = f"DWNUM{secrets.token_hex(24)}X"
        self._prefixes: dict[str, str] = {}
        self._pattern = re.compile(re.escape(self._scope) + r"[0-9]+Z ")

    def mark(self, prefix: str) -> str:
        if not prefix.rstrip():
            return ""
        # Keep the original whitespace boundary before the scheme prefix so
        # link processing sees the same bare-URL and inline-delimiter context.
        marker = f"{self._scope}{len(self._prefixes)}Z "
        self._prefixes[marker] = prefix
        return marker

    def strip_markers(self, text: str) -> str:
        return self._pattern.sub(lambda match: "" if match.group() in self._prefixes else match.group(), text)

    def prepare_source(
        self, source: str, *, extended_headings: bool = False, structural_tables: bool = False
    ) -> tuple[str, list[str], int]:
        if not self._prefixes:
            return source, [], 0
        output: list[str] = []
        positions: dict[int, str] = {}
        cursor = 0
        removed = 0
        for match in self._pattern.finditer(source):
            prefix = self._prefixes.get(match.group())
            if prefix is None:
                continue
            output.append(source[cursor : match.start()])
            leading = len(prefix) - len(prefix.lstrip(" \t"))
            # ATX/Setext heading parsing consumes prefix-leading ASCII space
            # and tab characters. Move only across those exact owned bytes.
            if source[match.end() : match.end() + leading] != prefix[:leading]:
                leading = 0
            positions[match.start() - removed + leading] = prefix[leading:]
            removed += match.end() - match.start()
            cursor = match.end()
        output.append(source[cursor:])
        clean = "".join(output)
        prefixes = []
        for heading in markdown_heading_ranges(
            clean, extended_headings=extended_headings, structural_tables=structural_tables
        ):
            prefixes.append(positions.pop(heading.text_start, ""))
        return clean, prefixes, len(positions)


def add_numbering_review(document: Any, run: Any) -> None:
    """Attach the numbering review using the shared template-safe writer."""
    append_docx_comment(document, [run], text=NUMBERING_REVIEW_TEXT)


def annotate_numbering_prefixes(
    nodes: list[dict[str, Any]], prefixes: list[str], *, comments: bool = True, text_numbering: bool = False
) -> int:
    """Bind the numbering transform's exact inventory to parsed headings.

    Do not infer ownership from visible numbers: authored handwritten numbers
    remain part of the title. Blank scheme levels still consume inventory slots.
    """
    headings: list[dict[str, Any]] = []

    def visit(items: list[dict[str, Any]]) -> None:
        for item in items:
            if item.get("type") == "footnotes":
                # Notes are rendered separately and their headings never
                # received incoming-text numbering in the source transform.
                continue
            if item.get("type") == "heading":
                headings.append(item)
            else:
                visit(item.get("children", []))

    visit(nodes)
    if len(headings) != len(prefixes):
        return sum(bool(prefix.rstrip()) for prefix in prefixes) if comments else 0
    skipped = 0
    for node, prefix in zip(headings, prefixes, strict=True):
        # Whitespace separating a number from its title is not the number.
        number = prefix.rstrip()
        if not number:
            continue
        if text_numbering:
            node["_docwen_text_numbering"] = True
        if not comments:
            continue
        children = node.get("children", [])
        if not children or children[0].get("type") != "text" or not children[0].get("raw", "").startswith(number):
            skipped += 1
            continue
        node["_docwen_added_number"] = number
    return skipped
