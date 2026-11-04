"""Ordinary heading navigation, isolated from frozen semantic block anchors."""

from __future__ import annotations

import re
import secrets
from collections import defaultdict
from typing import Any
from urllib.parse import unquote

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import encode_markdown_angle_destination
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer_inlines import extract_text_content


class DocxNavigationLinks:
    """Resolve only unique rendered heading nodes, never ordinary ^block IDs."""

    def __init__(self, config: LinkRuntimeConfig, source: str = "") -> None:
        self.source = source
        self.config = config
        self.prefix = "docwen-navigation:" + secrets.token_hex(20)
        self.links: dict[str, tuple[str, str]] = {}
        self.hints: dict[str, str] = {}
        self.targets: dict[str, str] = {}
        self.nodes: dict[int, tuple[str, str]] = {}

    def issue(self, label: str, target: str, original: str, syntax: str) -> str:
        if not target.startswith("#"):
            return f"[{label}](<{encode_markdown_angle_destination(target)}>)"
        mode = self.config.wiki_hyperlink_fallback if syntax == "wiki" else self.config.markdown_hyperlink_fallback
        visible = "".join(
            extract_text_content(node.get("children", []))
            for node in parse_markdown_text(label, consume_initial_bom=False)
        )
        fallback = target if mode == "extract_address" else visible if mode == "extract_text" else original
        token = self.prefix + str(len(self.links))
        self.links[token] = (unquote(target[1:]), fallback)
        self.hints[token] = target
        return f"[{label}](<{token}>)"

    def prepare(self, nodes: list[dict[str, Any]], allocator: Any) -> None:
        # Embedded documents have their own source identity. Until a provenance
        # map exists, never guess a root-to-embedded heading binding.
        if "![[" in self.source:
            return
        candidates: dict[str, list[dict[str, Any]]] = defaultdict(list)

        def visit(items: list[dict[str, Any]]) -> None:
            for node in items:
                if node.get("type") == "heading":
                    title = extract_text_content(node.get("children", [])).strip()
                    # No guessed block-ID navigation, and duplicate names never pick a winner.
                    keys = {title, re.sub(r"[^\w\- ]", "", title.lower()).replace(" ", "-")}
                    for key in keys:
                        if key and not key.startswith("^"):
                            candidates[key].append(node)
                visit(node.get("children", []))

        visit(nodes)
        for fragment, _ in self.links.values():
            matching = candidates.get(fragment, [])
            if len(matching) != 1 or fragment.startswith("^"):
                continue
            node = matching[0]
            if id(node) not in self.nodes:
                name = "DWNav_" + secrets.token_hex(16)
                self.nodes[id(node)] = (name, allocator.reserve(name))
            self.targets[fragment] = self.nodes[id(node)][0]

    def bind(self, node: dict[str, Any], paragraph: Any) -> None:
        binding = self.nodes.get(id(node))
        if binding is None:
            return
        name, bookmark_id = binding
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), bookmark_id)
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), bookmark_id)
        paragraph._p.insert(1 if paragraph._p.pPr is not None else 0, start)
        paragraph._p.append(end)

    def resolve(self, target: str) -> tuple[str | None, str] | None:
        entry = self.links.get(target)
        if entry is None:
            return None
        fragment, fallback = entry
        return self.targets.get(fragment), fallback
