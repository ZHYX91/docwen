from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import NoReturn
from xml.etree import ElementTree

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_W = f"{{{WML}}}"
_CAPTION_KINDS = {"figure", "table", "equation", "code_block"}


class V4OoxmlTargetError(ValueError):
    """The physical target inventory is not a unique typed source projection."""


@dataclass(frozen=True)
class TargetProjection:
    paragraphs: Mapping[tuple[int, int, str], ElementTree.Element]
    carrier_tags: Mapping[tuple[int, int, str], str | None]


def _fail(code: str) -> NoReturn:
    raise V4OoxmlTargetError(code)


def _key(target: Mapping[str, object]) -> tuple[int, int, str]:
    start, end, kind = target.get("source_start"), target.get("source_end"), str(target.get("kind"))
    if (
        not isinstance(start, int)
        or isinstance(start, bool)
        or not isinstance(end, int)
        or isinstance(end, bool)
        or start < 0
        or end <= start
        or kind not in {*_CAPTION_KINDS, "heading"}
    ):
        _fail("ooxml_target_source_identity_invalid")
    return start, end, kind


def _parents(root: ElementTree.Element) -> dict[ElementTree.Element, ElementTree.Element]:
    return {child: parent for parent in root.iter() for child in parent}


def _ancestor(
    element: ElementTree.Element,
    parents: Mapping[ElementTree.Element, ElementTree.Element],
    tag: str,
) -> ElementTree.Element | None:
    current = parents.get(element)
    while current is not None:
        if current.tag == tag:
            return current
        current = parents.get(current)
    return None


def _sdt_tag(sdt: ElementTree.Element) -> str | None:
    tags = sdt.findall(f"{_W}sdtPr/{_W}tag")
    return tags[0].get(f"{_W}val") if len(tags) == 1 else None


def _style(paragraph: ElementTree.Element) -> str | None:
    properties = paragraph.findall(f"{_W}pPr")
    styles = paragraph.findall(f"{_W}pPr/{_W}pStyle")
    if len(properties) != 1 or len(styles) != 1 or set(styles[0].attrib) != {f"{_W}val"}:
        return None
    return styles[0].get(f"{_W}val")


def _content(sdt: ElementTree.Element, code: str) -> ElementTree.Element:
    properties, contents = sdt.findall(f"{_W}sdtPr"), sdt.findall(f"{_W}sdtContent")
    tags = [] if len(properties) != 1 else list(properties[0])
    if (
        sdt.attrib
        or sdt.text is not None
        or sdt.tail is not None
        or len(properties) != 1
        or len(contents) != 1
        or list(sdt) != [properties[0], contents[0]]
        or properties[0].attrib
        or properties[0].text is not None
        or properties[0].tail is not None
        or len(tags) != 1
        or tags[0].tag != f"{_W}tag"
        or set(tags[0].attrib) != {f"{_W}val"}
        or tags[0].text is not None
        or list(tags[0])
        or contents[0].attrib
        or contents[0].text is not None
        or contents[0].tail is not None
    ):
        _fail(code)
    return contents[0]


def _heading_paragraph(sdt: ElementTree.Element) -> ElementTree.Element:
    children = list(_content(sdt, "ooxml_heading_target_envelope_invalid"))
    if len(children) != 1 or children[0].tag != f"{_W}p":
        _fail("ooxml_heading_target_envelope_invalid")
    return children[0]


def _logical_object(element: ElementTree.Element) -> ElementTree.Element:
    if element.tag != f"{_W}sdt" or not (_sdt_tag(element) or "").startswith("docwen-anchor-v1:"):
        return element
    children = list(_content(element, "ooxml_caption_anchor_invalid"))
    if len(children) != 1:
        _fail("ooxml_caption_anchor_invalid")
    return children[0]


def _object_kind(element: ElementTree.Element, kind: str, forbidden_styles: set[str]) -> bool:
    logical = _logical_object(element)
    if logical.tag == f"{_W}p" and _style(logical) in forbidden_styles:
        return False
    if kind == "table":
        return logical.tag == f"{_W}tbl"
    if kind == "figure":
        return logical.tag == f"{_W}p" and bool(list(logical.iter(f"{_W}drawing")) or list(logical.iter(f"{_W}pict")))
    if kind == "equation":
        return bool(list(logical.iter(f"{{{MATH}}}oMath")) or list(logical.iter(f"{{{MATH}}}oMathPara")))
    if kind == "code_block":
        return logical.tag == f"{_W}p"
    return False


def _caption_object(
    body: ElementTree.Element,
    paragraph: ElementTree.Element,
    kind: str,
    parents: Mapping[ElementTree.Element, ElementTree.Element],
    forbidden_styles: set[str],
) -> tuple[ElementTree.Element, str | None]:
    owner = parents.get(paragraph)
    if owner is None:
        _fail("ooxml_caption_detached")
    carrier: ElementTree.Element | None = None
    if owner.tag == f"{_W}sdtContent":
        carrier = parents.get(owner)
        if carrier is None or carrier.tag != f"{_W}sdt":
            _fail("ooxml_caption_carrier_invalid")
        container = owner
    elif owner is body:
        container = body
    else:
        _fail("ooxml_caption_container_invalid")
    children = list(container)
    try:
        index = children.index(paragraph)
    except ValueError as exc:
        raise V4OoxmlTargetError("ooxml_caption_detached") from exc
    object_index = index - 1 if kind == "figure" else index + 1
    if object_index < 0 or object_index >= len(children):
        _fail("ooxml_caption_object_missing")
    logical_object = children[object_index]
    if not _object_kind(logical_object, kind, forbidden_styles):
        _fail("ooxml_caption_object_kind_mismatch")
    if carrier is not None and (len(children) != 2 or {index, object_index} != {0, 1}):
        _fail("ooxml_caption_carrier_cardinality_invalid")
    return logical_object, None if carrier is None else _sdt_tag(carrier)


def _heading_style_ids(
    numbering: ElementTree.Element | None,
    targets: Sequence[Mapping[str, object]],
    sdts: Mapping[str, ElementTree.Element],
    target_tags: Mapping[tuple[int, int, str], str],
) -> dict[int, str]:
    candidates: dict[int, set[str]] = {}
    if numbering is not None:
        for level in numbering.findall(f"{_W}abstractNum/{_W}lvl"):
            raw_level = level.get(f"{_W}ilvl")
            styles = level.findall(f"{_W}pStyle")
            style_id = styles[0].get(f"{_W}val") if len(styles) == 1 else None
            if raw_level is None or not raw_level.isdecimal() or not style_id:
                _fail("ooxml_heading_style_binding_invalid")
            candidates.setdefault(int(raw_level) + 1, set()).add(style_id)
    for target in targets:
        if target.get("kind") != "heading" or target.get("target_id") is None:
            continue
        key = _key(target)
        sdt = sdts.get(target_tags[key])
        paragraph = None if sdt is None else _heading_paragraph(sdt)
        style_id = None if paragraph is None else _style(paragraph)
        level = target.get("heading_level")
        if not isinstance(level, int) or isinstance(level, bool) or not style_id:
            _fail("ooxml_heading_style_binding_invalid")
        candidates.setdefault(level, set()).add(style_id)
    result: dict[int, str] = {}
    for target in targets:
        if target.get("kind") != "heading":
            continue
        level = target.get("heading_level")
        if not isinstance(level, int) or isinstance(level, bool) or len(candidates.get(level, ())) != 1:
            _fail("ooxml_heading_style_authority_missing")
        result[level] = next(iter(candidates[level]))
    return result


def project_targets(
    document: ElementTree.Element,
    numbering: ElementTree.Element | None,
    targets: Sequence[Mapping[str, object]],
    planned_targets: Sequence[Mapping[str, object]],
    sdts: Mapping[str, ElementTree.Element],
    target_tags: Mapping[tuple[int, int, str], str],
    caption_styles: Mapping[str, tuple[str, str]],
    occurrence_tags: Mapping[tuple[int, int, str], str],
) -> TargetProjection:
    """Pair typed source targets to one exact managed paragraph and logical object."""

    bodies = document.findall(f"{_W}body")
    if len(bodies) != 1:
        _fail("ooxml_body_invalid")
    body = bodies[0]
    parents = _parents(document)
    keys = [_key(item) for item in targets]
    if len(set(keys)) != len(keys):
        _fail("ooxml_target_source_identity_duplicate")
    heading_styles = _heading_style_ids(numbering, targets, sdts, target_tags)
    heading_ids = set(heading_styles.values())
    caption_ids = {value[0] for value in caption_styles.values()}
    ordered_targets = sorted(targets, key=_key)
    managed_ids = heading_ids | caption_ids
    ordered_paragraphs = [item for item in body.iter(f"{_W}p") if _style(item) in managed_ids]
    if len(ordered_paragraphs) != len(ordered_targets):
        _fail("ooxml_managed_target_paragraph_inventory_invalid")
    paragraphs = {
        _key(target): paragraph for target, paragraph in zip(ordered_targets, ordered_paragraphs, strict=True)
    }
    target_prefix = "docwen-target-v1:"
    expected_target_tags = set(target_tags.values())
    if {tag for tag in sdts if tag.startswith(target_prefix)} != expected_target_tags:
        _fail("ooxml_target_sdt_inventory_invalid")
    occurrence_prefix = "docwen-numbering-occurrence-v1:"
    if {tag for tag in sdts if tag.startswith(occurrence_prefix)} != set(occurrence_tags.values()):
        _fail("ooxml_occurrence_sdt_inventory_invalid")
    carriers: dict[tuple[int, int, str], str | None] = {}
    used_objects: set[ElementTree.Element] = set()
    plan_by_key = {_key(item): planned for item, planned in zip(targets, planned_targets, strict=True)}
    for target in targets:
        key = _key(target)
        paragraph = paragraphs[key]
        kind = key[2]
        target_id = target.get("target_id")
        heading_level = target.get("heading_level")
        expected_style = (
            heading_styles.get(heading_level)
            if kind == "heading" and isinstance(heading_level, int) and not isinstance(heading_level, bool)
            else caption_styles.get(kind, (None, ""))[0]
        )
        if not expected_style or _style(paragraph) != expected_style:
            _fail("ooxml_target_kind_style_mismatch")
        expected_tag = target_tags.get(key) if target_id is not None else None
        if kind == "heading":
            owner = parents.get(paragraph)
            carrier = _ancestor(paragraph, parents, f"{_W}sdt")
            if expected_tag is not None:
                if carrier is None or carrier is not sdts.get(expected_tag) or parents.get(carrier) is not body:
                    _fail("ooxml_heading_target_carrier_invalid")
            elif owner is not body:
                _fail("ooxml_idless_heading_carrier_invalid")
            carriers[key] = expected_tag
            continue
        object_element, actual_tag = _caption_object(
            body,
            paragraph,
            kind,
            parents,
            managed_ids,
        )
        if object_element in used_objects:
            _fail("ooxml_caption_object_reused")
        used_objects.add(object_element)
        if expected_tag is not None:
            carrier = _ancestor(paragraph, parents, f"{_W}sdt")
            if (
                carrier is None
                or carrier is not sdts.get(expected_tag)
                or parents.get(carrier) is not body
                or actual_tag != expected_tag
            ):
                _fail("ooxml_caption_target_carrier_invalid")
        elif plan_by_key[key].get("enabled") is True:
            if parents.get(paragraph) is not body or actual_tag is not None:
                _fail("ooxml_enabled_idless_caption_carrier_invalid")
        else:
            expected_occurrence = occurrence_tags.get(key)
            carrier = _ancestor(paragraph, parents, f"{_W}sdt")
            if (
                expected_occurrence is None
                or carrier is None
                or carrier is not sdts.get(expected_occurrence)
                or parents.get(carrier) is not body
                or actual_tag != expected_occurrence
            ):
                _fail("ooxml_disabled_idless_caption_carrier_invalid")
        carriers[key] = actual_tag
    return TargetProjection(paragraphs, carriers)


__all__ = ["TargetProjection", "V4OoxmlTargetError", "project_targets"]
