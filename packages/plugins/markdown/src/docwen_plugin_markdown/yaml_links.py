"""Request-local YAML links carried through string-based field processors."""

from __future__ import annotations

import secrets
from copy import deepcopy
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docwen_core.export_semantics import LinkRuntimeConfig
from docwen_core.links import project_field_links, reject_declared_input_link_lookups
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer_inlines import add_hyperlink, extract_text_content


class _RunWriter:
    def __init__(self, parent: Any, index: int, properties: Any) -> None:
        self.parent = parent
        self.index = index
        self.properties = properties

    def children(self, children: list[Any]) -> None:
        if not children:
            return
        run = OxmlElement("w:r")
        if self.properties is not None:
            run.append(deepcopy(self.properties))
        run.extend(children)
        self.parent.insert(self.index, run)
        self.index += 1

    def text(self, text: str) -> None:
        if text:
            child = OxmlElement("w:t")
            child.text = text
            child.set(qn("xml:space"), "preserve")
            self.children([child])


class YamlLinkProjection:
    """Only issued link tokens become hyperlinks; all other text stays literal."""

    def __init__(self, source: str, config: LinkRuntimeConfig, *, declared_inputs: bool = False) -> None:
        self.source = source
        self.config = config
        self.declared_inputs = declared_inputs
        self.links: dict[str, tuple[str, str]] = {}
        self.prefix = "DOCWENYAMLLINK" + secrets.token_hex(24)

    def project(self, value: Any) -> Any:
        if isinstance(value, dict):
            return {key: self.project(item) for key, item in value.items()}
        if isinstance(value, (list, tuple)):
            return type(value)(self.project(item) for item in value)
        if not isinstance(value, str):
            return value
        if self.declared_inputs:
            reject_declared_input_link_lookups(value)
        return project_field_links(
            value,
            source_file_path=self.source,
            wiki_mode=self.config.non_embed_wiki_mode,
            markdown_mode=self.config.non_embed_markdown_mode,
            search_dirs=self.config.search_dirs,
            on_not_found=self.config.file_not_found_mode,
            hyperlink_renderer=self._link,
        )

    def _link(self, label: str, target: str) -> str:
        nodes = parse_markdown_text(label)
        visible = "".join(extract_text_content(node.get("children", [])) for node in nodes)
        token = f"{self.prefix}N{len(self.links)}END"
        self.links[token] = (visible or label, target)
        return token

    def plain(self, value: Any) -> Any:
        if isinstance(value, dict):
            return {key: self.plain(item) for key, item in value.items()}
        if isinstance(value, (list, tuple)):
            return type(value)(self.plain(item) for item in value)
        if isinstance(value, str):
            for token, (label, _) in self.links.items():
                value = value.replace(token, label)
        return value

    def materialize(self, document: Any) -> None:
        """Split only token-bearing runs, preserving sibling text, breaks and styles."""
        roots = [document.element]
        roots.extend(
            part.element
            for part in document.part.related_parts.values()
            if hasattr(part, "element") and str(part.partname).startswith(("/word/header", "/word/footer"))
        )
        from docx.text.paragraph import Paragraph

        for root in roots:
            pending = [node for node in root.iter(qn("w:t")) if self.prefix in (node.text or "")]
            while pending:
                node = pending[0]
                source = node.text or ""
                run = node.getparent()
                parent = run.getparent()
                if parent.tag != qn("w:p"):
                    raise ValueError("YAML hyperlink placeholder must not be inside an existing hyperlink")
                # Header/footer paragraphs must create relationships in their own part.
                owner = document
                if root is not document.element:
                    owner = next(
                        part for part in document.part.related_parts.values() if getattr(part, "element", None) is root
                    )
                paragraph = Paragraph(parent, owner)
                index = parent.index(run)
                rpr = run.find(qn("w:rPr"))

                writer = _RunWriter(parent, index, rpr)
                children = list(run)
                node_index = children.index(node)
                writer.children([deepcopy(child) for child in children[:node_index] if child is not rpr])
                cursor = 0
                while cursor < len(source):
                    matches = [(source.find(token, cursor), token) for token in self.links if token in source[cursor:]]
                    if not matches:
                        writer.text(source[cursor:])
                        break
                    start, token = min(matches)
                    writer.text(source[cursor:start])
                    label, target = self.links[token]
                    before = len(parent)
                    add_hyperlink(paragraph, target, text=label)
                    added = list(parent)[before:]
                    if not added:
                        raise ValueError("YAML hyperlink could not be materialized")
                    for element in added:
                        for link_run in element.iter(qn("w:r")):
                            if rpr is not None:
                                properties = link_run.get_or_add_rPr()
                                for prop in rpr:
                                    previous = properties.find(prop.tag)
                                    if previous is not None:
                                        properties.remove(previous)
                                    properties.append(deepcopy(prop))
                        parent.remove(element)
                        parent.insert(writer.index, element)
                        writer.index += 1
                    cursor = start + len(token)
                writer.children([deepcopy(child) for child in children[node_index + 1 :]])
                parent.remove(run)
                pending = [node for node in root.iter(qn("w:t")) if self.prefix in (node.text or "")]
