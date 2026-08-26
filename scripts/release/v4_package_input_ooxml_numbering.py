from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from typing import Any, NoReturn, cast
from xml.etree import ElementTree

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W = f"{{{WML}}}"
_NUMBER_FORMATS = {
    "chinese_lower": "chineseCounting",
    "chinese_upper": "chineseCountingThousand",
    "arabic_half": "decimal",
    "arabic_full": "decimalFullWidth",
    "arabic_circled": "decimalEnclosedCircleChinese",
    "letter_upper": "upperLetter",
    "letter_lower": "lowerLetter",
    "roman_upper": "upperRoman",
    "roman_lower": "lowerRoman",
}


class V4OoxmlNumberingError(ValueError):
    """The Heading list graph differs from the resolved physical plan."""


def _fail(code: str) -> NoReturn:
    raise V4OoxmlNumberingError(code)


def _object(value: object, code: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        _fail(code)
    return cast(dict[str, Any], value)


def _array(value: object, code: str) -> list[Any]:
    if not isinstance(value, list):
        _fail(code)
    return cast(list[Any], value)


def _val(parent: ElementTree.Element, name: str) -> str | None:
    values = parent.findall(f"{_W}{name}")
    if len(values) != 1:
        _fail(f"ooxml_numbering_property_invalid:{name}")
    return values[0].get(f"{_W}val")


def _display(level: Mapping[str, object]) -> str:
    parts: list[str] = []
    for raw in _array(level.get("display"), "heading_display_invalid"):
        segment = _object(raw, "heading_display_segment_invalid")
        if set(segment) == {"literal"} and isinstance(segment["literal"], str):
            parts.append(segment["literal"])
        elif set(segment) == {"counter"}:
            counter = _object(segment["counter"], "heading_counter_segment_invalid")
            parts.append(f"%{counter.get('level')}")
        else:
            _fail("heading_display_segment_invalid")
    return "".join(parts)


def _indexes(
    numbering: ElementTree.Element,
) -> tuple[dict[int, ElementTree.Element], dict[int, ElementTree.Element]]:
    abstracts: dict[int, ElementTree.Element] = {}
    nums: dict[int, ElementTree.Element] = {}
    entries = [
        *((item, "abstractNumId", abstracts) for item in numbering.findall(f"{_W}abstractNum")),
        *((item, "numId", nums) for item in numbering.findall(f"{_W}num")),
    ]
    for item, attribute, destination in entries:
        raw = item.get(f"{_W}{attribute}")
        if raw is None or re.fullmatch(r"0|[1-9][0-9]*", raw) is None or int(raw) in destination:
            _fail("ooxml_numbering_id_invalid")
        destination[int(raw)] = item
    return abstracts, nums


def prove_heading_numbering(
    numbering: ElementTree.Element | None,
    definitions: Sequence[Mapping[str, object]],
    instances: Sequence[Mapping[str, object]],
    bindings: Sequence[tuple[Mapping[str, object], Mapping[str, object], ElementTree.Element, int, int]],
    styles_by_id: Mapping[str, tuple[str, str]],
) -> None:
    if not bindings:
        if numbering is not None or definitions or instances:
            _fail("ooxml_unexpected_heading_numbering")
        return
    if numbering is None or numbering.tag != f"{_W}numbering":
        _fail("ooxml_numbering_root_invalid")
    abstracts, nums = _indexes(numbering)
    if len(abstracts) != len(definitions) or len(nums) != len(instances):
        _fail("ooxml_numbering_inventory_not_exact")
    instance_num: dict[str, int] = {}
    definition_abstract: dict[str, int] = {}
    for _planned, materialized, paragraph, num_id, _level in bindings:
        instance_id = str(materialized.get("instance_id"))
        definition_id = str(materialized.get("definition_id"))
        if instance_id in instance_num and instance_num[instance_id] != num_id:
            _fail("ooxml_heading_instance_split")
        instance_num[instance_id] = num_id
        num = nums.get(num_id)
        if num is None:
            _fail("ooxml_heading_num_missing")
        abstract_raw = _val(num, "abstractNumId")
        if abstract_raw is None or not abstract_raw.isdecimal():
            _fail("ooxml_heading_abstract_reference_invalid")
        abstract_id = int(abstract_raw)
        if definition_id in definition_abstract and definition_abstract[definition_id] != abstract_id:
            _fail("ooxml_heading_definition_split")
        definition_abstract[definition_id] = abstract_id
        styles = paragraph.findall(f"{_W}pPr/{_W}pStyle")
        style_id = styles[0].get(f"{_W}val") if len(styles) == 1 else None
        if not style_id or styles_by_id.get(style_id, (None, ""))[0] != "paragraph":
            _fail("ooxml_heading_managed_style_missing")
    if set(instance_num) != {str(item.get("instance_id")) for item in instances} or set(definition_abstract) != {
        str(item.get("definition_id")) for item in definitions
    }:
        _fail("ooxml_heading_projection_not_exhaustive")
    definitions_by_id = {str(item.get("definition_id")): item for item in definitions}
    for definition_id, abstract_id in definition_abstract.items():
        abstract = abstracts.get(abstract_id)
        definition = definitions_by_id[definition_id]
        if abstract is None or _val(abstract, "multiLevelType") != "multilevel":
            _fail("ooxml_heading_abstract_invalid")
        actual_levels = {item.get(f"{_W}ilvl"): item for item in abstract.findall(f"{_W}lvl")}
        levels = [
            _object(item, "ooxml_heading_level_invalid")
            for item in _array(definition.get("levels"), "ooxml_heading_levels_invalid")
        ]
        if len(actual_levels) != len(levels):
            _fail("ooxml_heading_level_inventory_invalid")
        for level in levels:
            ilvl = str(int(level.get("level", 0)) - 1)
            actual = actual_levels.get(ilvl)
            if actual is None:
                _fail("ooxml_heading_level_missing")
            expected_values = {
                "start": str(level.get("start")),
                "numFmt": _NUMBER_FORMATS.get(str(level.get("number_format"))),
                "lvlRestart": "0"
                if level.get("restart_after_level") is None
                else str(level.get("restart_after_level")),
                "suff": str(level.get("suffix")),
                "lvlText": _display(level),
            }
            for name, value in expected_values.items():
                if value is None or _val(actual, name) != value:
                    _fail(f"ooxml_heading_level_property_mismatch:{name}")
            style_id = _val(actual, "pStyle")
            if not style_id or style_id not in styles_by_id:
                _fail("ooxml_heading_level_style_invalid")
    instances_by_id = {str(item.get("instance_id")): item for item in instances}
    for instance_id, num_id in instance_num.items():
        instance = instances_by_id[instance_id]
        num = nums[num_id]
        abstract_id = int(cast(str, _val(num, "abstractNumId")))
        if abstract_id != definition_abstract.get(str(instance.get("definition_id"))):
            _fail("ooxml_heading_instance_definition_mismatch")
        actual_overrides: dict[str, str] = {}
        for override in num.findall(f"{_W}lvlOverride"):
            ilvl, start = override.get(f"{_W}ilvl"), _val(override, "startOverride")
            if ilvl is None or start is None or ilvl in actual_overrides:
                _fail("ooxml_heading_override_invalid")
            actual_overrides[ilvl] = start
        expected_overrides = {
            str(int(item.get("level", 0)) - 1): str(item.get("value"))
            for raw in _array(instance.get("starts"), "ooxml_heading_starts_invalid")
            for item in (_object(raw, "ooxml_heading_start_invalid"),)
        }
        if actual_overrides != expected_overrides:
            _fail("ooxml_heading_overrides_mismatch")


__all__ = ["V4OoxmlNumberingError", "prove_heading_numbering"]
