"""Closed caption-style binding map and exact styles.xml authentication."""

from __future__ import annotations

import re
from collections.abc import Iterable
from typing import Any, Final
from zipfile import ZipFile

import lxml.etree as etree

from docwen_core._docx_semantics_v3_model import (
    CAPTION_STYLE_BINDING_MAP_NAMESPACE,
    CaptionStyleBindingV3,
    CaptionStyleKeyV3,
    DocxSemanticsV3Error,
)

CAPTION_STYLE_KEYS_V3: Final[tuple[CaptionStyleKeyV3, ...]] = (
    "figure_caption",
    "table_caption",
    "equation_caption",
    "code_block_caption",
)
_KIND_TO_KEY: Final = {
    "figure": "figure_caption",
    "table": "table_caption",
    "equation": "equation_caption",
    "code_block": "code_block_caption",
}
_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
_WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def validate_caption_style_bindings(
    bindings: Iterable[CaptionStyleBindingV3],
) -> tuple[CaptionStyleBindingV3, ...]:
    """Return the four bindings in their frozen canonical order."""

    values = tuple(bindings)
    by_key: dict[str, CaptionStyleBindingV3] = {}
    for item in values:
        if item.semantic_key not in CAPTION_STYLE_KEYS_V3 or item.semantic_key in by_key:
            raise DocxSemanticsV3Error("caption-style map semantic keys are not the exact closed set")
        if re.fullmatch(r"[A-Za-z][A-Za-z0-9]{0,252}", item.resolved_style_id) is None:
            raise DocxSemanticsV3Error("caption-style resolved style ID is not canonical")
        if not item.visible_name or item.visible_name != item.visible_name.strip() or len(item.visible_name) > 255:
            raise DocxSemanticsV3Error("caption-style visible name is not canonical")
        by_key[item.semantic_key] = item
    if set(by_key) != set(CAPTION_STYLE_KEYS_V3):
        raise DocxSemanticsV3Error("caption-style map must contain exactly four fixed bindings")
    ordered = tuple(by_key[key] for key in CAPTION_STYLE_KEYS_V3)
    if len({item.resolved_style_id.casefold() for item in ordered}) != 4:
        raise DocxSemanticsV3Error("caption-style resolved style IDs are not unique")
    if len({item.visible_name.casefold() for item in ordered}) != 4:
        raise DocxSemanticsV3Error("caption-style visible names are not unique")
    return ordered


def caption_style_binding_map_xml(bindings: Iterable[CaptionStyleBindingV3]) -> bytes:
    ordered = validate_caption_style_bindings(bindings)
    entries = "".join(
        f'<binding semantic_key="{item.semantic_key}" resolved_style_id="{item.resolved_style_id}" '
        f'visible_name="{_xml_attr(item.visible_name)}"/>'
        for item in ordered
    )
    root = (
        f'<documentCaptionStyleBindingMap xmlns="{CAPTION_STYLE_BINDING_MAP_NAMESPACE}" version="1">'
        f"{entries}</documentCaptionStyleBindingMap>"
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def parse_caption_style_binding_map(root: Any) -> tuple[CaptionStyleBindingV3, ...]:
    namespace = f"{{{CAPTION_STYLE_BINDING_MAP_NAMESPACE}}}"
    if (
        root.tag != f"{namespace}documentCaptionStyleBindingMap"
        or tuple(root.attrib.items()) != (("version", "1"),)
        or root.text is not None
        or root.tail is not None
    ):
        raise DocxSemanticsV3Error("caption-style map root is not closed and canonical")
    output: list[CaptionStyleBindingV3] = []
    for item in root:
        if (
            item.tag != f"{namespace}binding"
            or tuple(item.attrib) != ("semantic_key", "resolved_style_id", "visible_name")
            or item.text is not None
            or item.tail is not None
            or len(item) != 0
        ):
            raise DocxSemanticsV3Error("caption-style binding record is not closed and canonical")
        output.append(
            CaptionStyleBindingV3(
                semantic_key=item.get("semantic_key"),  # type: ignore[arg-type]
                resolved_style_id=item.get("resolved_style_id"),
                visible_name=item.get("visible_name"),
            )
        )
    ordered = validate_caption_style_bindings(output)
    if tuple(output) != ordered:
        raise DocxSemanticsV3Error("caption-style binding records are not canonically ordered")
    return ordered


def prove_caption_style_registry(package: ZipFile, bindings: tuple[CaptionStyleBindingV3, ...]) -> None:
    """Authenticate every resolved style against the main styles.xml registry."""

    if "word/styles.xml" not in package.namelist():
        raise DocxSemanticsV3Error("DOCX package lacks the main styles registry")
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    root = etree.fromstring(package.read("word/styles.xml"), parser)
    style_tag = f"{{{_WORD_NAMESPACE}}}style"
    name_tag = f"{{{_WORD_NAMESPACE}}}name"
    alias_tag = f"{{{_WORD_NAMESPACE}}}aliases"
    type_attr = f"{{{_WORD_NAMESPACE}}}type"
    id_attr = f"{{{_WORD_NAMESPACE}}}styleId"
    value_attr = f"{{{_WORD_NAMESPACE}}}val"
    styles = root.findall(style_tag)
    for binding in bindings:
        by_id = [item for item in styles if item.get(id_attr) == binding.resolved_style_id]
        if len(by_id) != 1 or by_id[0].get(type_attr) != "paragraph":
            raise DocxSemanticsV3Error("caption-style binding does not resolve one exact paragraph style")
        names = by_id[0].findall(name_tag)
        if (
            len(names) != 1
            or tuple(names[0].attrib) != (value_attr,)
            or names[0].get(value_attr) != binding.visible_name
            or names[0].text is not None
        ):
            raise DocxSemanticsV3Error("caption-style visible name is not exact and canonical")
        if by_id[0].find(f".//{alias_tag}") is not None:
            raise DocxSemanticsV3Error("caption-style binding must not use aliases")


def prove_caption_paragraph_style(
    paragraph: Any,
    kind: str,
    bindings: tuple[CaptionStyleBindingV3, ...],
) -> None:
    """Require the direct pStyle selected by the kind's authenticated binding."""

    key = _KIND_TO_KEY.get(kind)
    if key is None:
        raise DocxSemanticsV3Error("unsupported caption target kind")
    expected = next(item.resolved_style_id for item in bindings if item.semantic_key == key)
    p_pr = paragraph.findall(f"./{{{_WORD_NAMESPACE}}}pPr")
    styles = paragraph.findall(f"./{{{_WORD_NAMESPACE}}}pPr/{{{_WORD_NAMESPACE}}}pStyle")
    value_attr = f"{{{_WORD_NAMESPACE}}}val"
    if len(p_pr) != 1 or len(styles) != 1 or tuple(styles[0].attrib) != (value_attr,):
        raise DocxSemanticsV3Error("caption paragraph does not have one exact direct pStyle")
    if styles[0].get(value_attr) != expected:
        raise DocxSemanticsV3Error("caption paragraph pStyle does not match its semantic kind")


def caption_kind_for_paragraph_style(
    paragraph: Any,
    bindings: tuple[CaptionStyleBindingV3, ...],
) -> str | None:
    p_styles = paragraph.findall(f"./{{{_WORD_NAMESPACE}}}pPr/{{{_WORD_NAMESPACE}}}pStyle")
    if len(p_styles) != 1:
        return None
    style_id = p_styles[0].get(f"{{{_WORD_NAMESPACE}}}val")
    matches = [
        kind
        for kind, key in _KIND_TO_KEY.items()
        if next(item.resolved_style_id for item in bindings if item.semantic_key == key) == style_id
    ]
    return matches[0] if len(matches) == 1 else None


def _xml_attr(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace('"', "&quot;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\r", "&#13;")
        .replace("\n", "&#10;")
        .replace("\t", "&#9;")
    )
