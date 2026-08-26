"""Closed DOCX physical primitives for the v4 resolved-numbering port.

This module deliberately knows nothing about Markdown parsing, WikiLink
resolution, numbering profiles, or Machine resources.  It consumes only the
typed, already-validated numbering plan produced by
``docwen_core.models.resolved_numbering``.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import lxml.etree as etree

from docwen_core._docx_numbering_package import (
    CONTENT_TYPES_NS,
    NUMBERING_CONTENT_TYPE,
    NUMBERING_REL_TYPE,
    RELS_NS,
    WML_NS,
    ResolvedNumberingOoxmlError,
    _append_complex_field,
    _append_text_run,
    _append_zero_width_bookmark,
    _bookmark,
    _clear_paragraph_payload,
    _elements_equal,
    _existing_numbering_ids,
    _heading_abstract_num,
    _heading_num,
    _insert_num_pr,
    _lowest_unused,
    _prove_caption_bookmark_pair,
    _prove_complex_field,
    _prove_text_run,
    _require_bookmark_id,
)
from docwen_core._docx_numbering_package import (
    write_heading_numbering_projection as _write_heading_numbering_projection,
)

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_FIELD_FORMATS = {
    "arabic_half": "ARABIC",
    "letter_upper": "ALPHABETIC",
    "letter_lower": "alphabetic",
    "roman_upper": "ROMAN",
    "roman_lower": "roman",
}


@dataclass(frozen=True, slots=True)
class HeadingNumberingProjection:
    """Allocated OOXML identities and elements for one request snapshot."""

    abstract_nums: tuple[Any, ...]
    nums: tuple[Any, ...]
    definition_abstract_ids: tuple[tuple[str, int], ...]
    instance_num_ids: tuple[tuple[str, int], ...]

    def abstract_id(self, definition_id: str) -> int:
        for candidate, value in self.definition_abstract_ids:
            if candidate == definition_id:
                return value
        raise ResolvedNumberingOoxmlError(f"unknown Heading definition: {definition_id}")

    def num_id(self, instance_id: str) -> int:
        for candidate, value in self.instance_num_ids:
            if candidate == instance_id:
                return value
        raise ResolvedNumberingOoxmlError(f"unknown Heading instance: {instance_id}")


@dataclass(frozen=True, slots=True)
class _CaptionFieldPlan:
    label: str
    chapter_instruction: str | None
    chapter_result: str | None
    chapter_separator: str | None
    sequence_instruction: str
    sequence_result: str


def existing_numbering_ids(document: Any) -> tuple[set[int], set[int]]:
    """Return existing abstract/instance IDs without mutating the template."""

    return _existing_numbering_ids(document)


def create_heading_numbering_projection(
    plan: Any,
    *,
    heading_style_ids: Mapping[int, str],
    existing_abstract_ids: set[int] | None = None,
    existing_num_ids: set[int] | None = None,
) -> HeadingNumberingProjection:
    """Translate the closed Heading plan into deterministic numbering nodes.

    Definitions and instances are consumed in validator-frozen first-use
    order.  ``abstractNumId`` may use zero; ``numId`` starts at one because
    Word reserves ``numId=0`` as the explicit no-numbering sentinel.
    """

    used_abstract = set(existing_abstract_ids or ())
    used_num = set(existing_num_ids or ())
    definition_ids: list[tuple[str, int]] = []
    abstract_nums: list[Any] = []
    for definition in plan.heading_definitions:
        abstract_id = _lowest_unused(used_abstract, start=0)
        used_abstract.add(abstract_id)
        definition_ids.append((definition.definition_id, abstract_id))
        abstract_nums.append(
            _heading_abstract_num(
                abstract_id,
                definition,
                heading_style_ids=heading_style_ids,
            )
        )

    definition_id_map = dict(definition_ids)
    instance_ids: list[tuple[str, int]] = []
    nums: list[Any] = []
    for instance in plan.heading_instances:
        if instance.definition_id not in definition_id_map:
            raise ResolvedNumberingOoxmlError("Heading instance names an unknown definition")
        num_id = _lowest_unused(used_num, start=1)
        used_num.add(num_id)
        instance_ids.append((instance.instance_id, num_id))
        nums.append(_heading_num(num_id, definition_id_map[instance.definition_id], instance.starts))

    return HeadingNumberingProjection(
        abstract_nums=tuple(abstract_nums),
        nums=tuple(nums),
        definition_abstract_ids=tuple(definition_ids),
        instance_num_ids=tuple(instance_ids),
    )


def apply_heading_numbering(
    paragraph: Any,
    target: Any,
    projection: HeadingNumberingProjection,
    *,
    heading_style_ids: Mapping[int, str],
) -> None:
    """Apply direct list semantics to one enabled Heading paragraph."""

    materialization = target.materialization
    if not target.enabled or materialization is None:
        remove_paragraph_numbering(paragraph)
        return
    if getattr(materialization, "type", None) != "heading_list":
        raise ResolvedNumberingOoxmlError("Heading target lacks heading_list materialization")
    level = int(materialization.level)
    if level not in heading_style_ids:
        raise ResolvedNumberingOoxmlError(f"Heading {level} has no resolved style binding")
    expected_style_id = heading_style_ids[level]
    style_id = getattr(getattr(paragraph, "style", None), "style_id", None)
    if style_id != expected_style_id:
        raise ResolvedNumberingOoxmlError(
            f"Heading paragraph style {style_id!r} differs from resolved {expected_style_id!r}"
        )
    p_pr = paragraph._p.get_or_add_pPr()
    existing = p_pr.find(f"{{{WML_NS}}}numPr")
    if existing is not None:
        p_pr.remove(existing)
    num_pr = etree.Element(f"{{{WML_NS}}}numPr")
    ilvl = etree.SubElement(num_pr, f"{{{WML_NS}}}ilvl")
    ilvl.set(f"{{{WML_NS}}}val", str(level - 1))
    num_id = etree.SubElement(num_pr, f"{{{WML_NS}}}numId")
    num_id.set(f"{{{WML_NS}}}val", str(projection.num_id(materialization.instance_id)))
    _insert_num_pr(p_pr, num_pr)


def remove_paragraph_numbering(paragraph: Any) -> None:
    """Ensure a disabled target has no direct effective ``w:numPr``."""

    p_pr = paragraph._p.pPr
    if p_pr is None:
        return
    existing = p_pr.find(f"{{{WML_NS}}}numPr")
    if existing is not None:
        p_pr.remove(existing)


def materialize_caption_number(
    paragraph: Any,
    target: Any,
    *,
    authored_content: str,
    heading_style_names: Mapping[str, str],
    bookmark_name: str | None = None,
    bookmark_id: str | None = None,
    preserve_payload: bool = False,
) -> str:
    """Write the exact label/field/title shape for one caption target.

    The caption paragraph must already carry its resolved managed style.  The
    function never reads or edits Markdown and never derives a counter value.
    It returns the exact visible text expected from the cached field results.
    """

    materialization = target.materialization
    resolved_bookmark_id = _require_bookmark_id(bookmark_id) if bookmark_name is not None else None
    rendered_payload = tuple(item for item in paragraph._p if item.tag != f"{{{WML_NS}}}pPr")
    if not target.enabled:
        if materialization is not None or target.derived_number is not None:
            raise ResolvedNumberingOoxmlError("disabled caption contradicts its plan")
        _clear_paragraph_payload(paragraph._p)
        if bookmark_name is not None:
            assert resolved_bookmark_id is not None
            _append_zero_width_bookmark(paragraph._p, bookmark_name, resolved_bookmark_id)
        if preserve_payload:
            paragraph._p.extend(rendered_payload)
        else:
            _append_text_run(paragraph._p, authored_content)
        return authored_content
    if materialization is None or target.derived_number is None:
        raise ResolvedNumberingOoxmlError("enabled caption lacks materialization or derived number")
    if getattr(materialization, "type", None) not in {"simple_seq", "chapter_seq"}:
        raise ResolvedNumberingOoxmlError("caption materialization type is not portable")

    field_plan = _caption_field_plan(
        materialization,
        target.derived_number,
        heading_style_names=heading_style_names,
    )
    _clear_paragraph_payload(paragraph._p)

    _append_text_run(paragraph._p, field_plan.label)
    if bookmark_name is not None:
        assert resolved_bookmark_id is not None
        start = _bookmark("start", bookmark_name, resolved_bookmark_id)
        paragraph._p.append(start)

    if field_plan.chapter_instruction is not None:
        assert field_plan.chapter_result is not None
        assert field_plan.chapter_separator is not None
        _append_complex_field(
            paragraph._p,
            instruction=field_plan.chapter_instruction,
            cached_result=field_plan.chapter_result,
        )
        _append_text_run(paragraph._p, field_plan.chapter_separator)

    _append_complex_field(
        paragraph._p,
        instruction=field_plan.sequence_instruction,
        cached_result=field_plan.sequence_result,
    )

    if bookmark_name is not None:
        assert resolved_bookmark_id is not None
        paragraph._p.append(_bookmark("end", bookmark_name, resolved_bookmark_id))
    if preserve_payload:
        suffix = " " if rendered_payload else ""
        if suffix:
            _append_text_run(paragraph._p, suffix)
        paragraph._p.extend(rendered_payload)
    else:
        suffix = f" {authored_content}" if authored_content else ""
        if suffix:
            _append_text_run(paragraph._p, suffix)
    return f"{field_plan.label}{target.derived_number}{suffix}"


def prove_caption_number(
    paragraph_element: Any,
    target: Any,
    *,
    authored_content: str,
    heading_style_names: Mapping[str, str],
    bookmark_name: str | None,
    bookmark_inventory: Any,
    rendered_payload: tuple[Any, ...] | None = None,
) -> None:
    """Authenticate the exact disabled or enabled caption payload after reopen."""

    materialization = target.materialization
    payload = tuple(item for item in paragraph_element if item.tag != f"{{{WML_NS}}}pPr")
    if not target.enabled:
        if materialization is not None or target.derived_number is not None:
            raise ResolvedNumberingOoxmlError("disabled caption contradicts its plan")
        cursor = 0
        if bookmark_name is not None:
            cursor = _prove_caption_bookmark_pair(
                payload,
                cursor,
                cursor + 1,
                bookmark_name,
                bookmark_inventory,
                paragraph_element,
            )
        if rendered_payload is not None:
            _prove_payload_snapshot(payload[cursor:], rendered_payload)
            cursor = len(payload)
        elif authored_content:
            if cursor >= len(payload):
                raise ResolvedNumberingOoxmlError("disabled caption lost its authored content")
            _prove_text_run(payload[cursor], authored_content)
            cursor += 1
        if cursor != len(payload):
            raise ResolvedNumberingOoxmlError("disabled caption contains derived numbering payload")
        return

    if materialization is None or target.derived_number is None:
        raise ResolvedNumberingOoxmlError("enabled caption lacks materialization or derived number")
    field_plan = _caption_field_plan(
        materialization,
        target.derived_number,
        heading_style_names=heading_style_names,
    )
    cursor = 0
    if cursor >= len(payload):
        raise ResolvedNumberingOoxmlError("enabled caption lost its localized label")
    _prove_text_run(payload[cursor], field_plan.label)
    cursor += 1
    bookmark_start = cursor if bookmark_name is not None else None
    if bookmark_name is not None:
        cursor += 1
    if field_plan.chapter_instruction is not None:
        assert field_plan.chapter_result is not None
        assert field_plan.chapter_separator is not None
        cursor = _prove_complex_field(
            payload,
            cursor,
            instruction=field_plan.chapter_instruction,
            cached_result=field_plan.chapter_result,
        )
        if cursor >= len(payload):
            raise ResolvedNumberingOoxmlError("chapter caption lost its separator")
        _prove_text_run(payload[cursor], field_plan.chapter_separator)
        cursor += 1
    cursor = _prove_complex_field(
        payload,
        cursor,
        instruction=field_plan.sequence_instruction,
        cached_result=field_plan.sequence_result,
    )
    if bookmark_name is not None:
        assert bookmark_start is not None
        cursor = _prove_caption_bookmark_pair(
            payload,
            bookmark_start,
            cursor,
            bookmark_name,
            bookmark_inventory,
            paragraph_element,
        )
    suffix = " " if rendered_payload else f" {authored_content}" if authored_content else ""
    if suffix:
        if cursor >= len(payload):
            raise ResolvedNumberingOoxmlError("caption lost its authored content")
        _prove_text_run(payload[cursor], suffix)
        cursor += 1
    if rendered_payload is not None:
        _prove_payload_snapshot(payload[cursor:], rendered_payload)
        cursor = len(payload)
    if cursor != len(payload):
        raise ResolvedNumberingOoxmlError("caption has extra physical payload")


def _prove_payload_snapshot(actual: tuple[Any, ...], expected: tuple[Any, ...]) -> None:
    if len(actual) != len(expected) or any(
        not _elements_equal(left, right) for left, right in zip(actual, expected, strict=True)
    ):
        raise ResolvedNumberingOoxmlError("caption rendered payload differs after materialization")


def prove_heading_numbering_projection(path: str | Path, projection: HeadingNumberingProjection) -> None:
    """Prove every allocated abstract definition and instance is byte-structurally exact."""

    if not projection.abstract_nums and not projection.nums:
        return
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    with ZipFile(Path(path)) as package:
        if "word/numbering.xml" not in package.namelist():
            raise ResolvedNumberingOoxmlError("resolved Heading numbering part is missing")
        root = etree.fromstring(package.read("word/numbering.xml"), parser)
    if root.tag != f"{{{WML_NS}}}numbering":
        raise ResolvedNumberingOoxmlError("numbering.xml root is invalid")
    for expected in projection.abstract_nums:
        identifier = expected.get(f"{{{WML_NS}}}abstractNumId")
        matching = [
            item
            for item in root.findall(f"{{{WML_NS}}}abstractNum")
            if item.get(f"{{{WML_NS}}}abstractNumId") == identifier
        ]
        if len(matching) != 1 or not _elements_equal(matching[0], expected):
            raise ResolvedNumberingOoxmlError("resolved Heading abstract definition differs after reopen")
    for expected in projection.nums:
        identifier = expected.get(f"{{{WML_NS}}}numId")
        matching = [item for item in root.findall(f"{{{WML_NS}}}num") if item.get(f"{{{WML_NS}}}numId") == identifier]
        if len(matching) != 1 or not _elements_equal(matching[0], expected):
            raise ResolvedNumberingOoxmlError("resolved Heading numbering instance differs after reopen")


def _caption_field_plan(
    materialization: Any,
    derived_number: str,
    *,
    heading_style_names: Mapping[str, str],
) -> _CaptionFieldPlan:
    field_switch = _FIELD_FORMATS.get(materialization.number_format)
    if field_switch is None:
        raise ResolvedNumberingOoxmlError("caption field format is not portable")
    sequence_result = materialization.sequence_cached_number
    chapter_instruction: str | None = None
    chapter_result: str | None = None
    chapter_separator: str | None = None
    if materialization.type == "chapter_seq":
        chapter_separator = materialization.chapter_separator
        chapter_result = materialization.chapter_cached_number
        if not chapter_separator or not chapter_result or not sequence_result:
            raise ResolvedNumberingOoxmlError("chapter cached field results must be non-empty")
        if f"{chapter_result}{chapter_separator}{sequence_result}" != derived_number:
            raise ResolvedNumberingOoxmlError("chapter cached field results contradict derived number")
        style_name = heading_style_names.get(materialization.chapter_heading_style or "")
        if not style_name:
            raise ResolvedNumberingOoxmlError("chapter caption has no resolved Heading style name")
        chapter_instruction = f' STYLEREF "{style_name}" \\n '
    elif materialization.chapter_cached_number is not None or sequence_result != derived_number:
        raise ResolvedNumberingOoxmlError("simple cached field result contradicts derived number")

    counter = materialization.counter
    if materialization.sequence_action == "continue":
        sequence_instruction = f" SEQ {counter} \\* {field_switch} "
    elif materialization.sequence_action == "reset_to_start":
        if materialization.start_value is None:
            raise ResolvedNumberingOoxmlError("reset_to_start requires an explicit start")
        sequence_instruction = f" SEQ {counter} \\r {materialization.start_value} \\* {field_switch} "
    elif materialization.sequence_action == "restart_by_heading_level":
        restart_level = materialization.restart_heading_level
        restart_style = materialization.restart_heading_style
        if restart_level is None or not restart_style or not heading_style_names.get(restart_style):
            raise ResolvedNumberingOoxmlError("heading restart lacks an independently resolved style")
        sequence_instruction = f" SEQ {counter} \\s {restart_level} \\* {field_switch} "
    else:
        raise ResolvedNumberingOoxmlError("caption sequence action is not portable")
    return _CaptionFieldPlan(
        label=f"{materialization.localized_label}{materialization.label_separator}",
        chapter_instruction=chapter_instruction,
        chapter_result=chapter_result,
        chapter_separator=chapter_separator,
        sequence_instruction=sequence_instruction,
        sequence_result=sequence_result,
    )


def inline_reference_sdt(
    tag: str,
    *,
    bookmark_name: str,
    cached_number: str,
    heading_number_only: bool,
    alias: str | None,
) -> Any:
    """Create one reversible REF occurrence; Alias stays outside the field."""

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    sdt = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    tag_element = OxmlElement("w:tag")
    tag_element.set(qn("w:val"), tag)
    properties.append(tag_element)
    content = OxmlElement("w:sdtContent")
    instruction = f" REF {bookmark_name} \\n \\h " if heading_number_only else f" REF {bookmark_name} \\h "
    _append_complex_field(content, instruction=instruction, cached_result=cached_number)
    if alias is not None:
        _append_text_run(content, f" {alias}")
    sdt.extend((properties, content))
    return sdt


def write_heading_numbering_projection(path: str | Path, projection: HeadingNumberingProjection) -> None:
    """Write the allocated Heading projection into an existing DOCX package."""

    _write_heading_numbering_projection(path, projection)


__all__ = [
    "CONTENT_TYPES_NS",
    "NUMBERING_CONTENT_TYPE",
    "NUMBERING_REL_TYPE",
    "RELS_NS",
    "WML_NS",
    "HeadingNumberingProjection",
    "ResolvedNumberingOoxmlError",
    "apply_heading_numbering",
    "create_heading_numbering_projection",
    "existing_numbering_ids",
    "inline_reference_sdt",
    "materialize_caption_number",
    "prove_caption_number",
    "prove_heading_numbering_projection",
    "remove_paragraph_numbering",
    "write_heading_numbering_projection",
]
