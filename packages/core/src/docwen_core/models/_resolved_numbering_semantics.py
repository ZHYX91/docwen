"""Cross-record and deterministic-counter checks for resolved numbering."""

from __future__ import annotations

import hashlib
import re
from itertools import pairwise
from typing import NoReturn

from docwen_core.links._markdown_inline import parse_inline_link, parse_markdown_destination
from docwen_core.links._patterns import WIKI_EMBED_PATTERN
from docwen_core.models.resolved_numbering import (
    CaptionMaterialization,
    HeadingCounterSegment,
    HeadingDefinition,
    HeadingInstance,
    HeadingListMaterialization,
    HeadingLiteralSegment,
    HeadingNumberFormat,
    NumberingExportPlanEnvelope,
    NumberingTarget,
    ResolvedDocument,
    ResolvedDocumentEnvelope,
    ResolvedNumberingPlan,
    ResolvedNumberingPortError,
)
from docwen_core.text.numbering import (
    number_to_arabic_full,
    number_to_chinese,
    number_to_chinese_upper,
    number_to_circled,
    number_to_letter_lower,
    number_to_letter_upper,
    number_to_roman_lower,
    number_to_roman_upper,
)

_COUNTER_FOR_KIND = {
    "figure": "Figure",
    "table": "Table",
    "equation": "Equation",
    "code_block": "Code",
}
_WIKI_EMBED_RE = re.compile(WIKI_EMBED_PATTERN)
_CITATION_KEY_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,127}$")
_SOURCE_ID_RE = re.compile(r"^[A-Za-z0-9-]{1,128}$")
_MAX_AUTHORED_MARKDOWN_HEADING_LEVEL = 9


def validate_document(document: ResolvedDocument, source_sha256: str, code: str) -> None:
    source = document.authored_markdown
    if hashlib.sha256(source.encode("utf-8")).hexdigest() != source_sha256:
        _fail(code, "source_sha256 does not authenticate authored_markdown")

    target_keys = [item.occurrence_key for item in document.targets]
    if target_keys != sorted(target_keys) or len(target_keys) != len(set(target_keys)):
        _fail(code, "document targets must be unique and ordered by source range and kind")
    target_ids: set[str] = set()
    for index, target in enumerate(document.targets):
        location = f"document.targets[{index}]"
        source_slice = _source_slice(source, target.source_start, target.source_end, location, code)
        if hashlib.sha256(source_slice.encode("utf-8")).hexdigest() != target.source_slice_sha256:
            _fail(code, f"{location}.source_slice_sha256 does not authenticate its range")
        if target.authored_text and target.authored_text not in source_slice:
            _fail(code, f"{location}.authored_text is not present in the authenticated source slice")
        if target.kind == "heading":
            if (
                type(target.heading_level) is not int
                or not 1 <= target.heading_level <= _MAX_AUTHORED_MARKDOWN_HEADING_LEVEL
            ):
                _fail(code, f"{location}.heading_level is outside the authored Markdown boundary")
        elif target.heading_level is not None:
            _fail(code, f"{location}.heading_level is only valid for a Heading")
        if target.target_id is not None:
            if target.target_id in target_ids:
                _fail(code, f"{location}.target_id is duplicated")
            target_ids.add(target.target_id)

    reference_ranges = [(item.source_start, item.source_end) for item in document.references]
    if reference_ranges != sorted(reference_ranges) or any(
        left[1] > right[0] for left, right in pairwise(reference_ranges)
    ):
        _fail(code, "document references must be ordered and non-overlapping")
    targets = {item.occurrence_key: item for item in document.targets}
    for index, reference in enumerate(document.references):
        location = f"document.references[{index}]"
        source_slice = _source_slice(source, reference.source_start, reference.source_end, location, code)
        if source_slice != reference.authored_token:
            _fail(code, f"{location}.authored_token is not the exact authenticated source slice")
        if hashlib.sha256(source_slice.encode("utf-8")).hexdigest() != reference.source_slice_sha256:
            _fail(code, f"{location}.source_slice_sha256 does not authenticate its range")
        selector_kind, selector_id, alias = _parse_cross_reference_token(
            reference.authored_token,
            location,
            code,
        )
        if alias != reference.alias:
            _fail(code, f"{location}.alias contradicts the authored cross-reference token")
        target = targets.get(reference.target_occurrence_key)
        if target is None:
            _fail(code, f"{location} points to a missing target occurrence")
        if reference.target_id != target.target_id:
            _fail(code, f"{location}.target_id contradicts the target occurrence")
        if selector_kind == "stable":
            if target.target_id is None or selector_id != target.target_id:
                _fail(code, f"{location} stable selector contradicts the target source ID")
        elif target.kind != "heading":
            _fail(code, f"{location} soft Heading selector resolves to a non-Heading target")

    _validate_resolved_dependencies(document, source, code)


def _validate_resolved_dependencies(document: ResolvedDocument, source: str, code: str) -> None:
    resource_ids = [item.resource_id for item in document.resources]
    if resource_ids != sorted(resource_ids) or len(resource_ids) != len(set(resource_ids)):
        _fail(code, "document resources must be unique and ordered by resource_id")
    bibliography = [item for item in document.resources if item.role == "bibliography"]
    if len(bibliography) > 1:
        _fail(code, "document may contain at most one embedded bibliography")
    linked = {item.resource_id: item for item in document.resources if item.role == "linked_resource"}

    occurrence_keys = [(item.source_start, item.source_end, item.resource_id) for item in document.resource_occurrences]
    if occurrence_keys != sorted(occurrence_keys):
        _fail(code, "document resource occurrences must be ordered by range and resource_id")
    occurrence_ranges = [key[:2] for key in occurrence_keys]
    if any(left[1] > right[0] for left, right in pairwise(occurrence_ranges)):
        _fail(code, "document resource occurrences must not overlap")
    used_resources: set[str] = set()
    for index, occurrence in enumerate(document.resource_occurrences):
        location = f"document.resource_occurrences[{index}]"
        source_slice = _source_slice(source, occurrence.source_start, occurrence.source_end, location, code)
        if source_slice != occurrence.authored_token:
            _fail(code, f"{location}.authored_token is not the exact authenticated source slice")
        if hashlib.sha256(source_slice.encode("utf-8")).hexdigest() != occurrence.source_slice_sha256:
            _fail(code, f"{location}.source_slice_sha256 does not authenticate its range")
        locator = _parse_image_locator(occurrence.authored_token)
        if locator is None or locator != occurrence.authored_locator:
            _fail(code, f"{location}.authored_locator is not the token's parsed image locator")
        if occurrence.resource_id not in linked:
            _fail(code, f"{location}.resource_id does not name a linked_resource")
        used_resources.add(occurrence.resource_id)
    if used_resources != set(linked):
        _fail(code, "every linked_resource must be used by at least one range-bound occurrence")

    citation_ranges = [(item.source_start, item.source_end) for item in document.citations]
    if citation_ranges != sorted(citation_ranges) or any(
        left[1] > right[0] for left, right in pairwise(citation_ranges)
    ):
        _fail(code, "document citations must be ordered and non-overlapping")
    cluster_ids: set[str] = set()
    for index, citation in enumerate(document.citations):
        location = f"document.citations[{index}]"
        source_slice = _source_slice(source, citation.source_start, citation.source_end, location, code)
        if source_slice != citation.authored_token:
            _fail(code, f"{location}.authored_token is not the exact authenticated source slice")
        if hashlib.sha256(source_slice.encode("utf-8")).hexdigest() != citation.source_slice_sha256:
            _fail(code, f"{location}.source_slice_sha256 does not authenticate its range")
        if citation.cluster_id in cluster_ids:
            _fail(code, f"{location}.cluster_id is duplicated")
        cluster_ids.add(citation.cluster_id)
        parsed_keys = _parse_citation_keys(citation.authored_token, citation.form)
        item_keys = tuple(item.citation_key for item in citation.items)
        if parsed_keys is None or parsed_keys != item_keys:
            _fail(code, f"{location}.items do not exactly resolve the authored citation token")
        if len(item_keys) != len(set(item_keys)):
            _fail(code, f"{location}.items repeats one citation key")

    semantic_ranges = sorted(
        [(item.source_start, item.source_end) for item in document.references] + occurrence_ranges + citation_ranges
    )
    if any(left[1] > right[0] for left, right in pairwise(semantic_ranges)):
        _fail(code, "resolved reference, resource, and citation occurrences must not overlap")


def _parse_image_locator(token: str) -> str | None:
    markdown = parse_inline_link(token, 0, image=True)
    if markdown is not None and markdown.end == len(token):
        destination = parse_markdown_destination(markdown.target, allow_image_size=True)
        if destination is not None and destination.destination:
            return destination.destination
    wiki = _WIKI_EMBED_RE.fullmatch(token)
    if wiki is not None and wiki.group(1):
        return wiki.group(1)
    return None


def _parse_citation_keys(token: str, form: str) -> tuple[str, ...] | None:
    if form == "narrative":
        key = token[1:] if token.startswith("@") else ""
        return (key,) if _CITATION_KEY_RE.fullmatch(key) is not None else None
    if not token.startswith("[") or not token.endswith("]"):
        return None
    parts = token[1:-1].split(";")
    keys: list[str] = []
    for part in parts:
        candidate = part.strip()
        key = candidate[1:] if candidate.startswith("@") else ""
        if _CITATION_KEY_RE.fullmatch(key) is None:
            return None
        keys.append(key)
    return tuple(keys) or None


def _parse_cross_reference_token(
    token: str,
    location: str,
    code: str,
) -> tuple[str, str | None, str | None]:
    if not token.startswith("@[[") or not token.endswith("]]"):
        _fail(code, f"{location}.authored_token is not a semantic cross-reference")
    body = token[3:-2]
    if body.count("|") > 1:
        _fail(code, f"{location}.authored_token has a non-closed Alias form")
    destination, separator, raw_alias = body.partition("|")
    alias = raw_alias if separator else None
    if separator and not alias:
        _fail(code, f"{location}.authored_token has an empty Alias")
    _page_locator, marker, selector = destination.partition("#")
    if not marker or not selector:
        _fail(code, f"{location}.authored_token has no Heading selector")
    parts = selector.split("#")
    if any(not part for part in parts):
        _fail(code, f"{location}.authored_token has an empty Heading path segment")
    if selector.startswith("^"):
        source_id = selector[1:]
        if len(parts) != 1 or _SOURCE_ID_RE.fullmatch(source_id) is None:
            _fail(code, f"{location}.authored_token has an invalid stable selector")
        return "stable", source_id, alias
    if any(part.startswith("^") for part in parts):
        _fail(code, f"{location}.authored_token mixes a Heading path with a stable selector")
    return "soft", None, alias


def validate_plan(plan: ResolvedNumberingPlan) -> None:
    definitions = _validate_definitions(plan.heading_definitions)
    instances = _validate_instances(plan.heading_instances, definitions)
    target_keys = [target.occurrence_key for target in plan.targets]
    if target_keys != sorted(target_keys) or len(target_keys) != len(set(target_keys)):
        _invalid("plan targets must be unique and ordered by source range and kind")

    used_definitions: list[str] = []
    used_instances: list[str] = []
    heading_state: dict[str, dict[int, int]] = {}
    caption_formats: dict[str, str] = {}

    for index, target in enumerate(plan.targets):
        location = f"plan.targets[{index}]"
        if target.source_end <= target.source_start:
            _invalid(f"{location} has an empty or reversed source range")
        if not target.enabled:
            continue
        materialization = target.materialization
        if target.derived_number is None or materialization is None:
            _invalid(f"{location} enabled target is missing its number or materialization")
        if isinstance(materialization, HeadingListMaterialization):
            if target.kind != "heading":
                _invalid(f"{location} applies heading_list to a non-Heading target")
            if (
                type(materialization.level) is not int
                or not 1 <= materialization.level <= _MAX_AUTHORED_MARKDOWN_HEADING_LEVEL
            ):
                _invalid(f"{location} Heading materialization exceeds the authored Markdown boundary")
            definition = definitions.get(materialization.definition_id)
            instance = instances.get(materialization.instance_id)
            if definition is None or instance is None:
                _invalid(f"{location} references a missing Heading definition or instance")
            if instance.definition_id != definition.definition_id:
                _invalid(f"{location} Heading instance belongs to a different definition")
            levels = {level.level: level for level in definition.levels}
            level = levels.get(materialization.level)
            if level is None:
                _invalid(f"{location} references an undefined Heading level")
            _append_first_use(used_definitions, definition.definition_id)
            _append_first_use(used_instances, instance.instance_id)
            state = heading_state.setdefault(instance.instance_id, {})
            starts = {item.level: item.value for item in instance.starts}
            current = state.get(level.level)
            state[level.level] = starts.get(level.level, level.start) if current is None else current + 1
            for candidate in definition.levels:
                if candidate.restart_after_level == level.level and candidate.level != level.level:
                    state.pop(candidate.level, None)
            expected = _heading_display(definition, instance, level.level, state)
            if expected != target.derived_number:
                _invalid(f"{location}.derived_number {target.derived_number!r} does not equal simulated {expected!r}")
            continue

        if type(materialization) is not CaptionMaterialization or target.kind == "heading":
            _invalid(f"{location} has a kind/materialization mismatch")
        _validate_caption_materialization_boundary(materialization, location)
        expected_counter = _COUNTER_FOR_KIND[target.kind]
        if materialization.counter != expected_counter:
            _invalid(f"{location}.counter does not match target kind")
        prior_format = caption_formats.setdefault(materialization.counter, materialization.number_format)
        if prior_format != materialization.number_format:
            _unsupported(f"{location} changes the format of one reused SEQ counter")

        expected_number = materialization.sequence_cached_number
        if materialization.type == "chapter_seq":
            chapter_separator = materialization.chapter_separator
            chapter_cached_number = materialization.chapter_cached_number
            if chapter_separator is None or chapter_cached_number is None:
                _invalid(f"{location} chapter_seq is missing chapter binding")
            expected_number = chapter_cached_number + chapter_separator + materialization.sequence_cached_number
        elif materialization.chapter_cached_number is not None:
            _invalid(f"{location} simple_seq must have null chapter_cached_number")
        if expected_number != target.derived_number:
            _invalid(
                f"{location}.derived_number {target.derived_number!r} does not equal simulated {expected_number!r}"
            )

    if [item.definition_id for item in plan.heading_definitions] != used_definitions:
        _invalid("heading_definitions are not exactly ordered by first target use")
    if [item.instance_id for item in plan.heading_instances] != used_instances:
        _invalid("heading_instances are not exactly ordered by first target use")


def validate_port(
    document: ResolvedDocumentEnvelope,
    plan: NumberingExportPlanEnvelope,
) -> None:
    if (
        document.input_id != plan.input_id
        or document.source_sha256 != plan.source_sha256
        or document.plan_sha256 != plan.plan_sha256
    ):
        _invalid("resolved-document and numbering-plan identity pointers do not match")
    document_targets = {target.occurrence_key: target for target in document.document.targets}
    plan_targets = {target.occurrence_key: target for target in plan.plan.targets}
    if document_targets.keys() != plan_targets.keys():
        _invalid("numbering plan does not enumerate every document target exactly once")
    for key, document_target in document_targets.items():
        plan_target = plan_targets[key]
        if document_target.target_id != plan_target.target_id:
            _invalid("numbering target ID contradicts the resolved document")
        if document_target.kind == "heading":
            materialization = plan_target.materialization
            if (
                isinstance(materialization, HeadingListMaterialization)
                and materialization.level != document_target.heading_level
            ):
                _invalid("numbering Heading level contradicts the resolved document")
        elif isinstance(plan_target.materialization, HeadingListMaterialization):
            _invalid("caption target carries Heading materialization")
    _validate_caption_sequences(document.document, plan.plan)
    for reference in document.document.references:
        target = plan_targets[reference.target_occurrence_key]
        if not target.enabled or target.derived_number is None:
            _invalid("resolved reference points to an unnumbered target")
        if reference.cached_number != target.derived_number:
            _invalid("resolved reference cached number contradicts its target plan value")


def _validate_caption_sequences(
    document: ResolvedDocument,
    plan: ResolvedNumberingPlan,
) -> None:
    """Simulate SEQ/STYLEREF against every authenticated Heading occurrence.

    A ``SEQ \\s N`` scope is triggered by Heading-N style occurrences even when
    Structured Numbering is disabled for that Heading.  ``STYLEREF ... \\n``
    instead needs the latest Heading at its own level to have a proven list
    number.  The resolved document carries the levels for disabled Headings;
    the numbering plan deliberately does not invent materialization for them.
    """

    plan_targets = {target.occurrence_key: target for target in plan.targets}
    latest_headings: dict[int, NumberingTarget] = {}
    caption_state: dict[str, int] = {}
    restart_scopes: dict[tuple[str, int], tuple[int, int, str]] = {}

    for index, document_target in enumerate(document.targets):
        target = plan_targets[document_target.occurrence_key]
        if document_target.kind == "heading":
            heading_level = document_target.heading_level
            if type(heading_level) is not int or not 1 <= heading_level <= _MAX_AUTHORED_MARKDOWN_HEADING_LEVEL:
                _invalid(f"document.targets[{index}] Heading level exceeds the authored Markdown boundary")
            latest_headings[heading_level] = target
            continue
        if not target.enabled:
            continue
        location = f"plan.targets[{index}]"
        materialization = target.materialization
        if not isinstance(materialization, CaptionMaterialization):
            _invalid(f"{location} enabled caption is missing SEQ materialization")
        _validate_caption_materialization_boundary(materialization, location)

        if materialization.sequence_action == "continue":
            local_value = caption_state.get(materialization.counter, 0) + 1
        elif materialization.sequence_action == "reset_to_start":
            if materialization.start_value is None:
                _invalid(f"{location} reset_to_start is missing start_value")
            local_value = materialization.start_value
        else:
            restart_level = materialization.restart_heading_level
            if restart_level is None:
                _invalid(f"{location} restart action is missing its Heading level")
            scope = latest_headings.get(restart_level)
            if scope is None:
                _invalid(f"{location} has no preceding Heading at restart level {restart_level}")
            scope_key = (materialization.counter, restart_level)
            if restart_scopes.get(scope_key) != scope.occurrence_key:
                local_value = 1
                restart_scopes[scope_key] = scope.occurrence_key
            else:
                local_value = caption_state.get(materialization.counter, 0) + 1
        caption_state[materialization.counter] = local_value
        expected_sequence = _format_number(local_value, materialization.number_format)
        if materialization.sequence_cached_number != expected_sequence:
            _invalid(f"{location}.sequence_cached_number does not equal simulated counter value")

        if materialization.type != "chapter_seq":
            continue
        chapter_level = materialization.chapter_heading_level
        chapter_cached_number = materialization.chapter_cached_number
        if chapter_level is None or chapter_cached_number is None:
            _invalid(f"{location} chapter_seq is missing chapter binding")
        chapter = latest_headings.get(chapter_level)
        if chapter is None:
            _invalid(f"{location} has no preceding Heading at chapter level {chapter_level}")
        if not chapter.enabled or chapter.derived_number is None:
            _invalid(f"{location} latest chapter Heading is not numbered")
        if chapter_cached_number != chapter.derived_number:
            _invalid(f"{location}.chapter_cached_number does not equal the latest Heading number")


def _validate_caption_materialization_boundary(
    materialization: CaptionMaterialization,
    location: str,
) -> None:
    if materialization.type not in {"simple_seq", "chapter_seq"}:
        _invalid(f"{location} caption materialization type is outside the closed set")
    if materialization.type == "chapter_seq":
        chapter_level = materialization.chapter_heading_level
        if (
            type(chapter_level) is not int
            or not 1 <= chapter_level <= _MAX_AUTHORED_MARKDOWN_HEADING_LEVEL
            or materialization.chapter_heading_style != f"heading_{chapter_level}"
            or materialization.chapter_separator is None
            or materialization.chapter_cached_number is None
        ):
            _invalid(f"{location} chapter_seq has an invalid Heading binding")
    elif any(
        value is not None
        for value in (
            materialization.chapter_heading_level,
            materialization.chapter_heading_style,
            materialization.chapter_separator,
            materialization.chapter_cached_number,
        )
    ):
        _invalid(f"{location} simple_seq carries chapter fields")

    action = materialization.sequence_action
    if action == "continue":
        if any(
            value is not None
            for value in (
                materialization.start_value,
                materialization.restart_heading_level,
                materialization.restart_heading_style,
            )
        ):
            _invalid(f"{location} continue action carries reset/restart fields")
    elif action == "reset_to_start":
        if (
            type(materialization.start_value) is not int
            or materialization.start_value < 1
            or materialization.restart_heading_level is not None
            or materialization.restart_heading_style is not None
        ):
            _invalid(f"{location} reset_to_start has invalid reset/restart fields")
    elif action == "restart_by_heading_level":
        restart_level = materialization.restart_heading_level
        if (
            materialization.start_value != 1
            or type(restart_level) is not int
            or not 1 <= restart_level <= _MAX_AUTHORED_MARKDOWN_HEADING_LEVEL
            or materialization.restart_heading_style != f"heading_{restart_level}"
        ):
            _invalid(f"{location} restart action has an invalid Heading binding")
    else:
        _invalid(f"{location} sequence action is outside the closed set")


def _validate_definitions(definitions: tuple[HeadingDefinition, ...]) -> dict[str, HeadingDefinition]:
    result: dict[str, HeadingDefinition] = {}
    for definition_index, definition in enumerate(definitions):
        location = f"plan.heading_definitions[{definition_index}]"
        if definition.definition_id in result:
            _invalid(f"{location}.definition_id is duplicated")
        result[definition.definition_id] = definition
        levels = [level.level for level in definition.levels]
        if levels != sorted(levels) or len(levels) != len(set(levels)):
            _invalid(f"{location}.levels must be strictly increasing")
        by_level = {level.level: level for level in definition.levels}
        for level_index, level in enumerate(definition.levels):
            level_location = f"{location}.levels[{level_index}]"
            if level.restart_after_level is not None and (
                level.restart_after_level >= level.level or level.restart_after_level not in by_level
            ):
                _invalid(f"{level_location}.restart_after_level must name a defined lower level")
            counters = [item for item in level.display if isinstance(item, HeadingCounterSegment)]
            counter_levels = [item.level for item in counters]
            if counter_levels.count(level.level) != 1 or len(counter_levels) != len(set(counter_levels)):
                _unsupported(f"{level_location}.display has missing or repeated counter levels")
            for counter in counters:
                referenced = by_level.get(counter.level)
                if referenced is None or counter.level > level.level:
                    _unsupported(f"{level_location}.display contains a forward/unknown counter reference")
                if counter.number_format != referenced.number_format:
                    _unsupported(f"{level_location}.display changes a referenced counter format")
    return result


def _validate_instances(
    instances: tuple[HeadingInstance, ...], definitions: dict[str, HeadingDefinition]
) -> dict[str, HeadingInstance]:
    result: dict[str, HeadingInstance] = {}
    for index, instance in enumerate(instances):
        location = f"plan.heading_instances[{index}]"
        if instance.instance_id in result:
            _invalid(f"{location}.instance_id is duplicated")
        definition = definitions.get(instance.definition_id)
        if definition is None:
            _invalid(f"{location}.definition_id does not exist")
        starts = [item.level for item in instance.starts]
        if starts != sorted(starts) or len(starts) != len(set(starts)):
            _invalid(f"{location}.starts must be strictly increasing")
        defined_levels = {item.level for item in definition.levels}
        if any(level not in defined_levels for level in starts):
            _invalid(f"{location}.starts references an undefined level")
        result[instance.instance_id] = instance
    return result


def _heading_display(
    definition: HeadingDefinition,
    instance: HeadingInstance,
    level: int,
    state: dict[int, int],
) -> str:
    levels = {item.level: item for item in definition.levels}
    starts = {item.level: item.value for item in instance.starts}
    definition_level = levels[level]
    output: list[str] = []
    for segment in definition_level.display:
        if isinstance(segment, HeadingLiteralSegment):
            output.append(segment.literal)
        else:
            value = state.get(segment.level, starts.get(segment.level, levels[segment.level].start))
            output.append(_format_number(value, segment.number_format))
    return "".join(output)


def _format_number(value: int, number_format: HeadingNumberFormat | str) -> str:
    if number_format in {"chinese_lower", "chinese_upper"} and value > 99:
        _unsupported("Chinese counter values above 99 lack frozen three-host display parity")
    if number_format in {"roman_upper", "roman_lower"} and value > 3999:
        _unsupported("Roman counter values above 3999 lack frozen three-host display parity")
    if number_format == "arabic_circled" and value > 50:
        _unsupported("circled counter values above 50 lack frozen three-host display parity")
    converters = {
        "chinese_lower": number_to_chinese,
        "chinese_upper": number_to_chinese_upper,
        "arabic_half": str,
        "arabic_full": number_to_arabic_full,
        "arabic_circled": number_to_circled,
        "letter_upper": number_to_letter_upper,
        "letter_lower": number_to_letter_lower,
        "roman_upper": number_to_roman_upper,
        "roman_lower": number_to_roman_lower,
    }
    return str(converters[number_format](value))


def _append_first_use(items: list[str], value: str) -> None:
    if value not in items:
        items.append(value)


def _source_slice(source: str, start: int, end: int, location: str, code: str) -> str:
    if end <= start or end > len(source):
        _fail(code, f"{location} range is outside authored_markdown")
    return source[start:end]


def _unsupported(message: str) -> NoReturn:
    raise ResolvedNumberingPortError("docwen.numbering_export_plan.unsupported_materialization", message)


def _invalid(message: str) -> NoReturn:
    _fail("docwen.numbering_export_plan.invalid", message)


def _fail(code: str, message: str) -> NoReturn:
    raise ResolvedNumberingPortError(code, message)
