"""Read actual ENML elements without external entities or network retrieval."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from html.entities import html5
from xml.parsers import expat

from docwen_core.text.placeholders import placeholder_namespace

_ENTITY_OR_OPAQUE = re.compile(r"<!--.*?-->|<!\[CDATA\[.*?\]\]>|<\?.*?\?>|&([A-Za-z][A-Za-z0-9]+);", re.DOTALL)
_XML_ENTITIES = frozenset({"amp", "lt", "gt", "quot", "apos"})


def _known_html_entities(xml: str) -> str:
    """ENML may use HTML named characters; comments/CDATA remain literal."""

    def replace(match: re.Match[str]) -> str:
        name = match.group(1)
        if name is None or name in _XML_ENTITIES:
            return match.group()
        value = html5.get(name + ";")
        return "".join(f"&#{ord(character)};" for character in value) if value else match.group()

    return _ENTITY_OR_OPAQUE.sub(replace, xml)


def _parse_enml(xml: str) -> ET.Element:
    builder = ET.TreeBuilder(insert_comments=True, insert_pis=True)
    # The ENEX parser has already decoded content into Unicode. Do not let a
    # stale inner encoding declaration reinterpret those characters.
    parser = expat.ParserCreate(encoding="utf-8")
    parser.StartElementHandler = builder.start
    parser.EndElementHandler = builder.end
    parser.CharacterDataHandler = builder.data
    parser.CommentHandler = builder.comment
    parser.ProcessingInstructionHandler = builder.pi

    def doctype(_name: str, _system: str | None, _public: str | None, internal: bool) -> None:
        if internal:
            raise ValueError("ENEX ENML internal entity declarations are unsupported")

    def external(*_arguments: object) -> int:
        raise ValueError("ENEX ENML external entities are unsupported")

    parser.StartDoctypeDeclHandler = doctype
    parser.ExternalEntityRefHandler = external
    parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    try:
        parser.Parse(_known_html_entities(xml.strip()).encode(), True)
        return builder.close()
    except (expat.ExpatError, ET.ParseError) as exc:
        raise ValueError("ENEX ENML is not well-formed XML") from exc


def normalize_enml(enml: str) -> tuple[str, dict[str, str]]:
    if not enml:
        return "", {}
    root = _parse_enml(enml)
    if isinstance(root.tag, str) and root.tag.lower() == "en-note":
        root.tag = "div"
    container = ET.Element("div")
    container.append(root)
    authored = enml + "".join(container.itertext()) + ET.tostring(container, encoding="unicode")
    namespace = placeholder_namespace(authored, prefix="DOCWENENEXRESOURCE")
    token_resources: dict[str, str] = {}
    for parent in list(container.iter()):
        children = []
        for child in parent:
            if not isinstance(child.tag, str) or child.tag.lower() != "en-media":
                children.append(child)
                continue
            digest = child.get("hash", "").strip()
            if re.fullmatch(r"[0-9a-fA-F]{32}", digest) is None:
                raise ValueError("ENEX ENML media has an invalid resource hash")
            if len(child) or (child.text and child.text.strip()):
                raise ValueError("ENEX ENML media cannot contain body content")
            token = f"{namespace}{len(token_resources) + 1}TOKEN"
            token_resources[token] = digest.lower()
            replacement = ET.Element("span")
            replacement.text = token
            replacement.tail = child.tail
            children.append(replacement)
        parent[:] = children
    return ET.tostring(container, encoding="unicode", method="html").strip(), token_resources
