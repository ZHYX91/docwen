"""Callout presentation: ordinary quote content with a source-type comment."""

from __future__ import annotations

import re
from typing import Any, cast

from docx.oxml.text.run import CT_R
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docwen_core.docx_comments import append_docx_comment

_HEADER = re.compile(r"^\[!([^\s\[\]]+)\]([+-]?)(?:[ \t]+([^\n]*))?(?:\n|$)")


def plugin_callout_headers(markdown: Any) -> None:
    """Recognize source headers before inline parsing loses escape identity."""

    def parse_quote(block: Any, match: Any, state: Any) -> int:
        previous = state.tokens[-1] if state.tokens else None
        end = block.parse_block_quote(match, state)
        # The native rule may prepend its quote before an interrupted sibling.
        # Identify only the new quote; nested quotes already passed this rule.
        for node in state.tokens[-2:]:
            if node is previous or node.get("type") != "block_quote":
                continue
            children = node.get("children", [])
            if not children:
                continue
            first = children[0]
            source = first.get("text", "")
            header = _HEADER.match(source) if first.get("type") == "paragraph" else None
            if header is None:
                continue
            node["_docwen_callout"] = (header[1], header[2])
            title = (header[3] or "").strip()
            body = source[header.end() :]
            replacement = []
            if title:
                replacement.append({**first, "text": title + "\n"})
            if body:
                replacement.append({**first, "text": body})
            if not replacement and len(children) == 1:
                replacement.append({**first, "text": ""})
            children[:1] = replacement
        return end

    markdown.block.register("block_quote", None, parse_quote)


def add_callout_comment(document: Any, items: list[Any], metadata: tuple[str, str]) -> None:
    """Anchor source metadata to the first native paragraph, including tables."""
    for item in items:
        element = item._p if hasattr(item, "_p") else item._element
        paragraphs = (
            [element]
            if element.tag.endswith("}p")
            else list(element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"))
        )
        if not paragraphs:
            continue
        paragraph = Paragraph(paragraphs[0], document)
        runs = [
            Run(cast(CT_R, run), paragraph)
            for run in paragraph._p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        ] or [paragraph.add_run()]
        kind, fold = metadata
        text = f"Markdown callout type: {kind}"
        if fold:
            text += "\nOriginal state: " + ("expanded (+)" if fold == "+" else "collapsed (-)")
        append_docx_comment(document, runs, text=text)
        return
