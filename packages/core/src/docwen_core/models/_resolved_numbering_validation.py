"""Validation implementation for the v4 resolved-numbering dual input."""

from __future__ import annotations

import hashlib
import json
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Literal, NoReturn, cast

from docwen_core.models.resolved_numbering import (
    MAX_RESOLVED_NUMBERING_RESOURCE_BYTES,
    NUMBERING_EXPORT_PLAN_SCHEMA,
    NUMBERING_EXPORT_PLAN_SCHEMA_ID,
    RESOLVED_DOCUMENT_SCHEMA,
    RESOLVED_DOCUMENT_SCHEMA_ID,
    CaptionMaterialization,
    CaptionNumberFormat,
    HeadingCounterSegment,
    HeadingDefinition,
    HeadingInstance,
    HeadingLevelDefinition,
    HeadingListMaterialization,
    HeadingLiteralSegment,
    HeadingNumberFormat,
    HeadingStart,
    NumberingExportPlanEnvelope,
    NumberingTarget,
    ResolvedDocument,
    ResolvedDocumentEnvelope,
    ResolvedDocumentTarget,
    ResolvedNumberingPlan,
    ResolvedNumberingPort,
    ResolvedNumberingPortError,
    ResolvedReference,
    TargetKind,
)

_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_DEFINITION_ID_RE = re.compile(r"^[A-Za-z][A-Za-z0-9-]{0,63}$")
_SOURCE_TARGET_ID_RE = re.compile(r"^[A-Za-z0-9-]{1,128}$")
_HEADING_STYLE_RE = re.compile(r"^heading_([1-9])$")
_MAX_AUTHORED_MARKDOWN_HEADING_LEVEL = 9
_TARGET_KINDS = frozenset({"heading", "figure", "table", "equation", "code_block"})
_HEADING_FORMATS = frozenset(
    {
        "chinese_lower",
        "chinese_upper",
        "arabic_half",
        "arabic_full",
        "arabic_circled",
        "letter_upper",
        "letter_lower",
        "roman_upper",
        "roman_lower",
    }
)
_CAPTION_FORMATS = frozenset({"arabic_half", "letter_upper", "letter_lower", "roman_upper", "roman_lower"})
_COUNTER_FOR_KIND = {
    "figure": "Figure",
    "table": "Table",
    "equation": "Equation",
    "code_block": "Code",
}


class _DuplicateKey(ValueError):
    pass


def load_resolved_numbering_port(document_path: Path, plan_path: Path) -> ResolvedNumberingPort:
    try:
        document_bytes = document_path.read_bytes()
    except OSError as exc:
        raise ResolvedNumberingPortError(
            "docwen.resolved_document.invalid", "resolved-document resource cannot be read"
        ) from exc
    try:
        plan_bytes = plan_path.read_bytes()
    except OSError as exc:
        raise ResolvedNumberingPortError(
            "docwen.numbering_export_plan.invalid", "numbering-export-plan resource cannot be read"
        ) from exc
    return load_resolved_numbering_bytes(document_bytes, plan_bytes)


def load_resolved_numbering_bytes(document_bytes: bytes, plan_bytes: bytes) -> ResolvedNumberingPort:
    document_raw = _decode_resource(document_bytes, code="docwen.resolved_document.invalid")
    plan_raw = _decode_resource(plan_bytes, code="docwen.numbering_export_plan.invalid")
    document_envelope = _parse_document_envelope(document_raw)
    plan_envelope = _parse_plan_envelope(plan_raw)
    _validate_port(document_envelope, plan_envelope, plan_raw)
    return ResolvedNumberingPort(document_envelope, plan_envelope)


def canonicalize_numbering_plan(plan: object) -> bytes:
    _validate_canonical_value(plan, location="plan")
    try:
        text = json.dumps(
            plan,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        )
    except (TypeError, ValueError) as exc:
        raise ResolvedNumberingPortError(
            "docwen.numbering_export_plan.invalid", "plan cannot be RFC 8785 canonicalized"
        ) from exc
    return text.encode("utf-8")


def _decode_resource(data: bytes, *, code: str) -> dict[str, object]:
    if not data or len(data) > MAX_RESOLVED_NUMBERING_RESOURCE_BYTES or data.startswith(b"\xef\xbb\xbf"):
        raise ResolvedNumberingPortError(code, "resource must be strict UTF-8 JSON of at most 8 MiB")
    try:
        text = data.decode("utf-8", errors="strict")
        parsed = json.loads(
            text,
            object_pairs_hook=_closed_object,
            parse_float=_reject_noninteger,
            parse_constant=_reject_nonfinite,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, _DuplicateKey, ValueError) as exc:
        raise ResolvedNumberingPortError(code, "resource is not closed strict UTF-8 JSON") from exc
    if not isinstance(parsed, dict):
        raise ResolvedNumberingPortError(code, "resource envelope must be an object")
    return cast(dict[str, object], parsed)


def _closed_object(pairs: list[tuple[str, object]]) -> dict[str, object]:
    result: dict[str, object] = {}
    for key, value in pairs:
        if key in result:
            raise _DuplicateKey(key)
        result[key] = value
    return result


def _reject_noninteger(_value: str) -> NoReturn:
    raise ValueError("non-integer JSON number")


def _reject_nonfinite(_value: str) -> NoReturn:
    raise ValueError("non-finite JSON number")


def _parse_document_envelope(raw: dict[str, object]) -> ResolvedDocumentEnvelope:
    code = "docwen.resolved_document.invalid"
    _exact_keys(raw, {"$schema", "schema", "input_id", "source_sha256", "plan_sha256", "document"}, code)
    _const(raw["$schema"], RESOLVED_DOCUMENT_SCHEMA_ID, "envelope.$schema", code)
    _const(raw["schema"], RESOLVED_DOCUMENT_SCHEMA, "envelope.schema", code)
    input_id = _string(raw["input_id"], "envelope.input_id", code, minimum=1, maximum=256)
    source_sha256 = _sha256(raw["source_sha256"], "envelope.source_sha256", code)
    plan_sha256 = _sha256(raw["plan_sha256"], "envelope.plan_sha256", code)
    document_raw = _object(raw["document"], "envelope.document", code)
    _exact_keys(
        document_raw,
        {
            "authored_markdown",
            "targets",
            "references",
            "resource_occurrences",
            "citations",
            "resources",
        },
        code,
    )
    authored_markdown = _string(document_raw["authored_markdown"], "document.authored_markdown", code, minimum=0)
    _unicode_scalar_text(authored_markdown, "document.authored_markdown", code)
    targets_raw = _array(document_raw["targets"], "document.targets", code)
    references_raw = _array(document_raw["references"], "document.references", code)
    targets = tuple(_parse_document_target(item, index, code) for index, item in enumerate(targets_raw))
    references = tuple(_parse_reference(item, index, code) for index, item in enumerate(references_raw))
    from docwen_core.models._resolved_document_extras import parse_document_extras

    resource_occurrences, citations, resources = parse_document_extras(document_raw, code)
    document = ResolvedDocument(
        authored_markdown,
        targets,
        references,
        resource_occurrences,
        citations,
        resources,
    )
    _validate_document(document, source_sha256, code)
    return ResolvedDocumentEnvelope(input_id, source_sha256, plan_sha256, document)


def _parse_document_target(raw: object, index: int, code: str) -> ResolvedDocumentTarget:
    location = f"document.targets[{index}]"
    item = _object(raw, location, code)
    _exact_keys(
        item,
        {
            "source_start",
            "source_end",
            "source_slice_sha256",
            "kind",
            "target_id",
            "heading_level",
            "authored_text",
        },
        code,
    )
    kind = _target_kind(item["kind"], f"{location}.kind", code)
    heading_level = _nullable_int(
        item["heading_level"],
        f"{location}.heading_level",
        code,
        1,
        _MAX_AUTHORED_MARKDOWN_HEADING_LEVEL,
    )
    if (kind == "heading") != (heading_level is not None):
        _fail(code, f"{location}.heading_level contradicts kind")
    authored_text = _string(item["authored_text"], f"{location}.authored_text", code, minimum=0)
    _unicode_scalar_text(authored_text, f"{location}.authored_text", code)
    return ResolvedDocumentTarget(
        source_start=_integer(item["source_start"], f"{location}.source_start", code, 0),
        source_end=_integer(item["source_end"], f"{location}.source_end", code, 1),
        source_slice_sha256=_sha256(item["source_slice_sha256"], f"{location}.source_slice_sha256", code),
        kind=kind,
        target_id=_target_id(item["target_id"], f"{location}.target_id", code),
        heading_level=heading_level,
        authored_text=authored_text,
    )


def _parse_reference(raw: object, index: int, code: str) -> ResolvedReference:
    location = f"document.references[{index}]"
    item = _object(raw, location, code)
    _exact_keys(
        item,
        {
            "source_start",
            "source_end",
            "source_slice_sha256",
            "authored_token",
            "target_source_start",
            "target_source_end",
            "target_kind",
            "target_id",
            "cached_number",
            "alias",
        },
        code,
    )
    alias_raw = item["alias"]
    alias = None if alias_raw is None else _string(alias_raw, f"{location}.alias", code, minimum=1)
    return ResolvedReference(
        source_start=_integer(item["source_start"], f"{location}.source_start", code, 0),
        source_end=_integer(item["source_end"], f"{location}.source_end", code, 1),
        source_slice_sha256=_sha256(item["source_slice_sha256"], f"{location}.source_slice_sha256", code),
        authored_token=_string(item["authored_token"], f"{location}.authored_token", code, minimum=6),
        target_source_start=_integer(item["target_source_start"], f"{location}.target_source_start", code, 0),
        target_source_end=_integer(item["target_source_end"], f"{location}.target_source_end", code, 1),
        target_kind=_target_kind(item["target_kind"], f"{location}.target_kind", code),
        target_id=_target_id(item["target_id"], f"{location}.target_id", code),
        cached_number=_string(item["cached_number"], f"{location}.cached_number", code, minimum=1),
        alias=alias,
    )


def _parse_plan_envelope(raw: dict[str, object]) -> NumberingExportPlanEnvelope:
    code = "docwen.numbering_export_plan.invalid"
    _exact_keys(raw, {"$schema", "schema", "input_id", "source_sha256", "plan_sha256", "plan"}, code)
    _const(raw["$schema"], NUMBERING_EXPORT_PLAN_SCHEMA_ID, "envelope.$schema", code)
    _const(raw["schema"], NUMBERING_EXPORT_PLAN_SCHEMA, "envelope.schema", code)
    input_id = _string(raw["input_id"], "envelope.input_id", code, minimum=1, maximum=256)
    source_sha256 = _sha256(raw["source_sha256"], "envelope.source_sha256", code)
    plan_sha256 = _sha256(raw["plan_sha256"], "envelope.plan_sha256", code)
    plan_raw = _object(raw["plan"], "envelope.plan", code)
    _exact_keys(plan_raw, {"heading_definitions", "heading_instances", "targets"}, code)
    definitions = tuple(
        _parse_heading_definition(item, index, code)
        for index, item in enumerate(_array(plan_raw["heading_definitions"], "plan.heading_definitions", code))
    )
    instances = tuple(
        _parse_heading_instance(item, index, code)
        for index, item in enumerate(_array(plan_raw["heading_instances"], "plan.heading_instances", code))
    )
    targets = tuple(
        _parse_plan_target(item, index, code)
        for index, item in enumerate(_array(plan_raw["targets"], "plan.targets", code))
    )
    plan = ResolvedNumberingPlan(definitions, instances, targets)
    _validate_plan(plan)
    return NumberingExportPlanEnvelope(input_id, source_sha256, plan_sha256, plan)


def _parse_heading_definition(raw: object, index: int, code: str) -> HeadingDefinition:
    location = f"plan.heading_definitions[{index}]"
    item = _object(raw, location, code)
    _exact_keys(item, {"definition_id", "levels"}, code)
    definition_id = _definition_id(item["definition_id"], f"{location}.definition_id", code)
    levels = tuple(
        _parse_heading_level(level, level_index, code, location)
        for level_index, level in enumerate(_array(item["levels"], f"{location}.levels", code))
    )
    if not 1 <= len(levels) <= 9:
        _fail(code, f"{location}.levels must contain 1..9 definitions")
    return HeadingDefinition(definition_id, levels)


def _parse_heading_level(raw: object, index: int, code: str, parent: str) -> HeadingLevelDefinition:
    location = f"{parent}.levels[{index}]"
    item = _object(raw, location, code)
    _exact_keys(item, {"level", "start", "number_format", "display", "suffix", "restart_after_level"}, code)
    display_raw = _array(item["display"], f"{location}.display", code)
    if not 1 <= len(display_raw) <= 19:
        _fail(code, f"{location}.display must contain 1..19 segments")
    display = tuple(_parse_display_segment(segment, i, code, location) for i, segment in enumerate(display_raw))
    number_format = _enum_string(item["number_format"], _HEADING_FORMATS, f"{location}.number_format", code)
    suffix = _enum_string(item["suffix"], frozenset({"nothing", "space", "tab"}), f"{location}.suffix", code)
    return HeadingLevelDefinition(
        level=_integer(item["level"], f"{location}.level", code, 1, 9),
        start=_integer(item["start"], f"{location}.start", code, 1, 2147483647),
        number_format=cast(HeadingNumberFormat, number_format),
        display=display,
        suffix=cast(Literal["nothing", "space", "tab"], suffix),
        restart_after_level=_nullable_int(item["restart_after_level"], f"{location}.restart_after_level", code, 1, 8),
    )


def _parse_display_segment(raw: object, index: int, code: str, parent: str):
    location = f"{parent}.display[{index}]"
    item = _object(raw, location, code)
    if set(item) == {"literal"}:
        literal = _string(item["literal"], f"{location}.literal", code, minimum=1, maximum=32)
        if "%" in literal or not _is_xml_10_text(literal):
            _unsupported(f"{location}.literal cannot be translated losslessly to OOXML")
        return HeadingLiteralSegment(literal)
    if set(item) == {"counter"}:
        counter = _object(item["counter"], f"{location}.counter", code)
        _exact_keys(counter, {"level", "number_format"}, code)
        number_format = _enum_string(
            counter["number_format"], _HEADING_FORMATS, f"{location}.counter.number_format", code
        )
        return HeadingCounterSegment(
            _integer(counter["level"], f"{location}.counter.level", code, 1, 9),
            cast(HeadingNumberFormat, number_format),
        )
    _fail(code, f"{location} must be exactly one counter or literal segment")


def _parse_heading_instance(raw: object, index: int, code: str) -> HeadingInstance:
    location = f"plan.heading_instances[{index}]"
    item = _object(raw, location, code)
    _exact_keys(item, {"instance_id", "definition_id", "starts"}, code)
    starts_raw = _array(item["starts"], f"{location}.starts", code)
    if len(starts_raw) > 9:
        _fail(code, f"{location}.starts has too many entries")
    starts: list[HeadingStart] = []
    for start_index, raw_start in enumerate(starts_raw):
        start_location = f"{location}.starts[{start_index}]"
        start = _object(raw_start, start_location, code)
        _exact_keys(start, {"level", "value"}, code)
        starts.append(
            HeadingStart(
                _integer(start["level"], f"{start_location}.level", code, 1, 9),
                _integer(start["value"], f"{start_location}.value", code, 1, 2147483647),
            )
        )
    return HeadingInstance(
        _definition_id(item["instance_id"], f"{location}.instance_id", code),
        _definition_id(item["definition_id"], f"{location}.definition_id", code),
        tuple(starts),
    )


def _parse_plan_target(raw: object, index: int, code: str) -> NumberingTarget:
    location = f"plan.targets[{index}]"
    item = _object(raw, location, code)
    _exact_keys(
        item,
        {
            "source_start",
            "source_end",
            "kind",
            "enabled",
            "target_id",
            "derived_number",
            "materialization",
        },
        code,
    )
    kind = _target_kind(item["kind"], f"{location}.kind", code)
    enabled = _boolean(item["enabled"], f"{location}.enabled", code)
    derived_raw = item["derived_number"]
    materialization_raw = item["materialization"]
    if not enabled:
        if derived_raw is not None or materialization_raw is not None:
            _fail(code, f"{location} disabled target must carry null number and materialization")
        derived_number = None
        materialization = None
    else:
        derived_number = _string(derived_raw, f"{location}.derived_number", code, minimum=1)
        materialization_item = _object(materialization_raw, f"{location}.materialization", code)
        materialization_type = materialization_item.get("type")
        if kind == "heading":
            if materialization_type != "heading_list":
                _fail(code, f"{location} enabled heading requires heading_list materialization")
            materialization = _parse_heading_materialization(materialization_item, f"{location}.materialization", code)
        else:
            if materialization_type not in {"simple_seq", "chapter_seq"}:
                _fail(code, f"{location} enabled caption requires SEQ materialization")
            materialization = _parse_caption_materialization(materialization_item, f"{location}.materialization", code)
    return NumberingTarget(
        source_start=_integer(item["source_start"], f"{location}.source_start", code, 0),
        source_end=_integer(item["source_end"], f"{location}.source_end", code, 1),
        kind=kind,
        enabled=enabled,
        target_id=_target_id(item["target_id"], f"{location}.target_id", code),
        derived_number=derived_number,
        materialization=materialization,
    )


def _parse_heading_materialization(item: dict[str, object], location: str, code: str) -> HeadingListMaterialization:
    _exact_keys(item, {"type", "definition_id", "instance_id", "level"}, code)
    _const(item["type"], "heading_list", f"{location}.type", code)
    return HeadingListMaterialization(
        definition_id=_definition_id(item["definition_id"], f"{location}.definition_id", code),
        instance_id=_definition_id(item["instance_id"], f"{location}.instance_id", code),
        level=_integer(
            item["level"],
            f"{location}.level",
            code,
            1,
            _MAX_AUTHORED_MARKDOWN_HEADING_LEVEL,
        ),
    )


def _parse_caption_materialization(item: dict[str, object], location: str, code: str) -> CaptionMaterialization:
    _exact_keys(
        item,
        {
            "type",
            "counter",
            "number_format",
            "sequence_action",
            "start_value",
            "chapter_heading_level",
            "chapter_heading_style",
            "chapter_separator",
            "restart_heading_level",
            "restart_heading_style",
            "chapter_cached_number",
            "sequence_cached_number",
            "localized_label",
            "label_separator",
        },
        code,
    )
    materialization_type = _enum_string(
        item["type"], frozenset({"simple_seq", "chapter_seq"}), f"{location}.type", code
    )
    counter = _enum_string(item["counter"], frozenset(_COUNTER_FOR_KIND.values()), f"{location}.counter", code)
    number_format = _enum_string(item["number_format"], _CAPTION_FORMATS, f"{location}.number_format", code)
    action = _enum_string(
        item["sequence_action"],
        frozenset({"continue", "reset_to_start", "restart_by_heading_level"}),
        f"{location}.sequence_action",
        code,
    )
    start_value = _nullable_int(item["start_value"], f"{location}.start_value", code, 1, 2147483647)
    chapter_level = _nullable_int(
        item["chapter_heading_level"],
        f"{location}.chapter_heading_level",
        code,
        1,
        _MAX_AUTHORED_MARKDOWN_HEADING_LEVEL,
    )
    chapter_style = _nullable_string(item["chapter_heading_style"], f"{location}.chapter_heading_style", code)
    chapter_separator = _nullable_string(item["chapter_separator"], f"{location}.chapter_separator", code)
    restart_level = _nullable_int(
        item["restart_heading_level"],
        f"{location}.restart_heading_level",
        code,
        1,
        _MAX_AUTHORED_MARKDOWN_HEADING_LEVEL,
    )
    restart_style = _nullable_string(item["restart_heading_style"], f"{location}.restart_heading_style", code)
    chapter_cached_number = _nullable_string(item["chapter_cached_number"], f"{location}.chapter_cached_number", code)
    sequence_cached_number = _string(
        item["sequence_cached_number"],
        f"{location}.sequence_cached_number",
        code,
        minimum=1,
    )
    if (chapter_cached_number is not None and not _is_xml_10_text(chapter_cached_number)) or not _is_xml_10_text(
        sequence_cached_number
    ):
        _unsupported(f"{location} cached number is not XML 1.0 compatible")

    if materialization_type == "chapter_seq":
        _validate_heading_binding(chapter_level, chapter_style, f"{location}.chapter", code)
        if chapter_separator is None or not 1 <= len(chapter_separator) <= 8:
            _fail(code, f"{location}.chapter_separator is required for chapter_seq")
        if chapter_cached_number is None:
            _fail(code, f"{location}.chapter_cached_number is required for chapter_seq")
        if not _is_xml_10_text(chapter_separator):
            _unsupported(f"{location}.chapter_separator is not XML 1.0 text")
    elif (
        chapter_level is not None
        or chapter_style is not None
        or chapter_separator is not None
        or chapter_cached_number is not None
    ):
        _fail(code, f"{location} simple_seq must not carry chapter fields")

    if action == "continue":
        if start_value is not None or restart_level is not None or restart_style is not None:
            _fail(code, f"{location} continue action has contradictory reset/restart fields")
    elif action == "reset_to_start":
        if start_value is None or restart_level is not None or restart_style is not None:
            _fail(code, f"{location} reset_to_start requires only start_value")
    else:
        if start_value != 1:
            _unsupported(f"{location} restart_by_heading_level requires start_value 1")
        _validate_heading_binding(restart_level, restart_style, f"{location}.restart", code)

    localized_label = _string(item["localized_label"], f"{location}.localized_label", code, minimum=1, maximum=64)
    label_separator = _string(item["label_separator"], f"{location}.label_separator", code, minimum=0, maximum=8)
    if not _is_xml_10_text(localized_label) or not _is_xml_10_text(label_separator):
        _unsupported(f"{location} label text is not XML 1.0 compatible")
    return CaptionMaterialization(
        type=cast(Literal["simple_seq", "chapter_seq"], materialization_type),
        counter=cast(Literal["Figure", "Table", "Equation", "Code"], counter),
        number_format=cast(CaptionNumberFormat, number_format),
        sequence_action=cast(Literal["continue", "reset_to_start", "restart_by_heading_level"], action),
        start_value=start_value,
        chapter_heading_level=chapter_level,
        chapter_heading_style=chapter_style,
        chapter_separator=chapter_separator,
        restart_heading_level=restart_level,
        restart_heading_style=restart_style,
        chapter_cached_number=chapter_cached_number,
        sequence_cached_number=sequence_cached_number,
        localized_label=localized_label,
        label_separator=label_separator,
    )


def _validate_heading_binding(level: int | None, style: str | None, location: str, code: str) -> None:
    if level is None or style is None:
        _fail(code, f"{location} level and style are required together")
    match = _HEADING_STYLE_RE.fullmatch(style)
    if match is None or int(match.group(1)) != level:
        _fail(code, f"{location} style must be heading_N for the same level")


def _validate_document(document: ResolvedDocument, source_sha256: str, code: str) -> None:
    from docwen_core.models._resolved_numbering_semantics import validate_document

    validate_document(document, source_sha256, code)


def _validate_plan(plan: ResolvedNumberingPlan) -> None:
    from docwen_core.models._resolved_numbering_semantics import validate_plan

    validate_plan(plan)


def _validate_port(
    document: ResolvedDocumentEnvelope,
    plan: NumberingExportPlanEnvelope,
    plan_raw: dict[str, object],
) -> None:
    from docwen_core.models._resolved_numbering_semantics import validate_port

    raw_plan = plan_raw.get("plan")
    digest = hashlib.sha256(canonicalize_numbering_plan(raw_plan)).hexdigest()
    if digest != plan.plan_sha256:
        _fail("docwen.numbering_export_plan.invalid", "plan_sha256 does not authenticate plan")
    validate_port(document, plan)


def _validate_canonical_value(value: object, *, location: str) -> None:
    if value is None or type(value) in {bool, int}:
        return
    if isinstance(value, str):
        _unicode_scalar_text(value, location, "docwen.numbering_export_plan.invalid")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _validate_canonical_value(item, location=f"{location}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            if not isinstance(key, str):
                _fail("docwen.numbering_export_plan.invalid", f"{location} has a non-string key")
            _unicode_scalar_text(key, location, "docwen.numbering_export_plan.invalid")
            _validate_canonical_value(item, location=f"{location}.{key}")
        return
    _fail("docwen.numbering_export_plan.invalid", f"{location} contains a non-canonical JSON value")


def _exact_keys(item: Mapping[str, object], expected: set[str], code: str) -> None:
    if set(item) != expected:
        _fail(code, "closed JSON object has missing or additional properties")


def _object(value: object, location: str, code: str) -> dict[str, object]:
    if not isinstance(value, dict):
        _fail(code, f"{location} must be an object")
    return cast(dict[str, object], value)


def _array(value: object, location: str, code: str) -> list[object]:
    if not isinstance(value, list):
        _fail(code, f"{location} must be an array")
    return cast(list[object], value)


def _string(
    value: object,
    location: str,
    code: str,
    *,
    minimum: int,
    maximum: int | None = None,
) -> str:
    if not isinstance(value, str) or len(value) < minimum or (maximum is not None and len(value) > maximum):
        _fail(code, f"{location} has invalid string length")
    return value


def _nullable_string(value: object, location: str, code: str) -> str | None:
    return None if value is None else _string(value, location, code, minimum=1)


def _integer(value: object, location: str, code: str, minimum: int, maximum: int | None = None) -> int:
    if type(value) is not int or value < minimum or (maximum is not None and value > maximum):
        _fail(code, f"{location} must be an integer in range")
    return cast(int, value)


def _nullable_int(value: object, location: str, code: str, minimum: int, maximum: int) -> int | None:
    return None if value is None else _integer(value, location, code, minimum, maximum)


def _boolean(value: object, location: str, code: str) -> bool:
    if type(value) is not bool:
        _fail(code, f"{location} must be a boolean")
    return cast(bool, value)


def _enum_string(value: object, allowed: frozenset[str], location: str, code: str) -> str:
    result = _string(value, location, code, minimum=1)
    if result not in allowed:
        _fail(code, f"{location} has an unsupported enum value")
    return result


def _const(value: object, expected: str, location: str, code: str) -> None:
    if value != expected:
        _fail(code, f"{location} has the wrong identity")


def _sha256(value: object, location: str, code: str) -> str:
    result = _string(value, location, code, minimum=64, maximum=64)
    if _SHA256_RE.fullmatch(result) is None:
        _fail(code, f"{location} must be lowercase SHA-256")
    return result


def _definition_id(value: object, location: str, code: str) -> str:
    result = _string(value, location, code, minimum=1, maximum=64)
    if _DEFINITION_ID_RE.fullmatch(result) is None:
        _fail(code, f"{location} is not a portable plan identifier")
    return result


def _target_id(value: object, location: str, code: str) -> str | None:
    if value is None:
        return None
    result = _string(value, location, code, minimum=1, maximum=128)
    if _SOURCE_TARGET_ID_RE.fullmatch(result) is None:
        _fail(code, f"{location} is not a valid source target ID")
    return result


def _target_kind(value: object, location: str, code: str) -> TargetKind:
    return cast(TargetKind, _enum_string(value, _TARGET_KINDS, location, code))


def _unicode_scalar_text(value: str, location: str, code: str) -> None:
    if any(0xD800 <= ord(character) <= 0xDFFF for character in value):
        _fail(code, f"{location} contains an unpaired surrogate")


def _is_xml_10_text(value: str) -> bool:
    return all(
        character in {"\t", "\n", "\r"}
        or "\u0020" <= character <= "\ud7ff"
        or "\ue000" <= character <= "\ufffd"
        or "\U00010000" <= character <= "\U0010ffff"
        for character in value
    )


def _unsupported(message: str) -> NoReturn:
    raise ResolvedNumberingPortError("docwen.numbering_export_plan.unsupported_materialization", message)


def _fail(code: str, message: str) -> NoReturn:
    raise ResolvedNumberingPortError(code, message)
