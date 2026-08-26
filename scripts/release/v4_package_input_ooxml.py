from __future__ import annotations

import io
import re
import zipfile
from collections import Counter
from collections.abc import Mapping
from dataclasses import dataclass
from typing import Any, NoReturn, cast
from xml.etree import ElementTree

from scripts.release import v4_package_input_ooxml_identity as identity_proof
from scripts.release import v4_package_input_ooxml_maps as map_proof
from scripts.release import v4_package_input_ooxml_numbering as numbering_proof
from scripts.release import v4_package_input_ooxml_opc as opc_proof
from scripts.release import v4_package_input_ooxml_targets as target_proof

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_W = f"{{{WML}}}"
_HEX64 = re.compile(r"^[0-9a-f]{64}$")
_BOOKMARK = re.compile(r"^(?:DW_T_[0-9a-f]{35}|_DWC_[0-9a-f]{35})$")
_FIELD_FORMATS = {
    "arabic_half": "ARABIC",
    "letter_upper": "ALPHABETIC",
    "letter_lower": "alphabetic",
    "roman_upper": "ROMAN",
    "roman_lower": "roman",
}
_CAPTION_KINDS = {"figure", "table", "equation", "code_block"}


class V4OoxmlProofError(ValueError):
    """A DOCX does not prove the exact resolved v4 physical plan."""


@dataclass(frozen=True)
class _Field:
    instruction: str
    cached_result: str
    paragraph: ElementTree.Element


@dataclass(frozen=True)
class _Bookmark:
    name: str
    bookmark_id: str
    start: ElementTree.Element
    end: ElementTree.Element


def _fail(code: str) -> NoReturn:
    raise V4OoxmlProofError(code)


def _object(value: object, code: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        _fail(code)
    return cast(dict[str, Any], value)


def _array(value: object, code: str) -> list[Any]:
    if not isinstance(value, list):
        _fail(code)
    return cast(list[Any], value)


def _xml(raw: bytes, *, part: str) -> ElementTree.Element:
    if b"<!DOCTYPE" in raw.upper() or b"<!ENTITY" in raw.upper():
        _fail(f"ooxml_xml_declaration_rejected:{part}")
    try:
        return ElementTree.fromstring(raw)
    except ElementTree.ParseError as exc:
        raise V4OoxmlProofError(f"ooxml_xml_invalid:{part}") from exc


def _identity(operation: str, *arguments: object) -> tuple[str, ...] | str:
    try:
        function = getattr(identity_proof, operation)
        return function(*arguments)
    except (identity_proof.V4OoxmlIdentityError, KeyError) as exc:
        raise V4OoxmlProofError(str(exc)) from exc


def _target_identity(kind: str, target_id: str) -> tuple[str, str]:
    return cast(tuple[str, str], _identity("target_identity", kind, target_id))


def _reference_tag(reference: Mapping[str, object], bookmark: str, source_sha256: str) -> str:
    return cast(str, _identity("reference_tag", reference, bookmark, source_sha256))


def _citation_identity(citation: Mapping[str, object], source_sha256: str) -> tuple[str, str, str]:
    return cast(tuple[str, str, str], _identity("citation_identity", citation, source_sha256))


def _parent_map(root: ElementTree.Element) -> dict[ElementTree.Element, ElementTree.Element]:
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
    values = [item.get(f"{_W}val") for item in sdt.findall(f"{_W}sdtPr/{_W}tag")]
    return values[0] if len(values) == 1 and values[0] else None


def _sdt_index(root: ElementTree.Element) -> dict[str, ElementTree.Element]:
    result: dict[str, ElementTree.Element] = {}
    for sdt in root.iter(f"{_W}sdt"):
        value = _sdt_tag(sdt)
        if value is None or not value.startswith("docwen-"):
            continue
        if value in result:
            _fail(f"ooxml_sdt_tag_duplicate:{value}")
        result[value] = sdt
    return result


def _direct_content_paragraph(sdt: ElementTree.Element, *, code: str) -> ElementTree.Element:
    contents = sdt.findall(f"{_W}sdtContent")
    if len(contents) != 1:
        _fail(code)
    paragraphs = contents[0].findall(f"{_W}p")
    if len(paragraphs) != 1:
        _fail(code)
    return paragraphs[0]


def _carrier_paragraph(
    sdt: ElementTree.Element,
    parents: Mapping[ElementTree.Element, ElementTree.Element],
    *,
    code: str,
) -> ElementTree.Element:
    inline = _ancestor(sdt, parents, f"{_W}p")
    return inline if inline is not None else _direct_content_paragraph(sdt, code=code)


def _visible_text(element: ElementTree.Element) -> str:
    return "".join(item.text or "" for item in element.iter(f"{_W}t"))


def _fields(
    root: ElementTree.Element,
    paragraphs: set[ElementTree.Element],
) -> tuple[_Field, ...]:
    result: list[_Field] = []
    for paragraph in root.iter(f"{_W}p"):
        if paragraph not in paragraphs:
            continue
        active = False
        separated = False
        instruction: list[str] = []
        cached: list[str] = []
        begin: ElementTree.Element | None = None
        for item in paragraph.iter():
            if item.tag == f"{_W}fldChar":
                field_type = item.get(f"{_W}fldCharType")
                if field_type == "begin":
                    if active:
                        _fail("ooxml_nested_field_rejected")
                    active, separated, begin = True, False, item
                    instruction, cached = [], []
                elif field_type == "separate":
                    if not active or separated:
                        _fail("ooxml_field_separator_invalid")
                    separated = True
                elif field_type == "end":
                    if not active or not separated or begin is None:
                        _fail("ooxml_field_end_invalid")
                    instruction_text = "".join(instruction)
                    cached_text = "".join(cached)
                    if not instruction_text or not cached_text:
                        _fail("ooxml_field_instruction_or_cache_empty")
                    result.append(_Field(instruction_text, cached_text, paragraph))
                    active, separated, begin = False, False, None
                else:
                    _fail("ooxml_field_type_invalid")
            elif item.tag == f"{_W}instrText":
                if not active or separated:
                    _fail("ooxml_instruction_outside_field")
                instruction.append(item.text or "")
            elif item.tag == f"{_W}t" and active and separated:
                cached.append(item.text or "")
        if active:
            _fail("ooxml_field_unterminated")
    return tuple(result)


def _bookmarks(
    root: ElementTree.Element,
    parents: Mapping[ElementTree.Element, ElementTree.Element],
) -> dict[str, _Bookmark]:
    starts: dict[str, ElementTree.Element] = {}
    ends: dict[str, ElementTree.Element] = {}
    names: dict[str, str] = {}
    for item in root.iter(f"{_W}bookmarkStart"):
        bookmark_id, name = item.get(f"{_W}id"), item.get(f"{_W}name")
        if (
            bookmark_id is None
            or name is None
            or bookmark_id in starts
            or name.casefold() in {x.casefold() for x in names.values()}
        ):
            _fail("ooxml_bookmark_start_invalid")
        if re.fullmatch(r"0|[1-9][0-9]*", bookmark_id) is None:
            _fail("ooxml_bookmark_identity_noncanonical")
        owned_signal = name.casefold().startswith(("dw_t_", "_dwc_"))
        if owned_signal and _BOOKMARK.fullmatch(name) is None:
            _fail("ooxml_owned_bookmark_identity_noncanonical")
        starts[bookmark_id], names[bookmark_id] = item, name
    for item in root.iter(f"{_W}bookmarkEnd"):
        bookmark_id = item.get(f"{_W}id")
        if bookmark_id is None or bookmark_id in ends:
            _fail("ooxml_bookmark_end_invalid")
        ends[bookmark_id] = item
    if set(starts) != set(ends):
        _fail("ooxml_bookmark_range_unbalanced")
    positions = {item: index for index, item in enumerate(root.iter())}
    result: dict[str, _Bookmark] = {}
    for bookmark_id, start in starts.items():
        end = ends[bookmark_id]
        if positions[start] >= positions[end]:
            _fail("ooxml_bookmark_range_reversed")
        name = names[bookmark_id]
        if _BOOKMARK.fullmatch(name) is not None:
            start_p = _ancestor(start, parents, f"{_W}p")
            end_p = _ancestor(end, parents, f"{_W}p")
            if start_p is None or start_p is not end_p:
                _fail("ooxml_owned_bookmark_cross_paragraph_rejected")
            result[name] = _Bookmark(name, bookmark_id, start, end)
    return result


def _prove_bookmark_scope(
    bookmark: _Bookmark,
    sdt: ElementTree.Element,
    parents: Mapping[ElementTree.Element, ElementTree.Element],
) -> None:
    if (
        _ancestor(bookmark.start, parents, f"{_W}sdt") is not sdt
        or _ancestor(bookmark.end, parents, f"{_W}sdt") is not sdt
    ):
        _fail(f"ooxml_bookmark_scope_invalid:{bookmark.name}")


def _style_index(styles: ElementTree.Element) -> dict[str, tuple[str, str]]:
    result: dict[str, tuple[str, str]] = {}
    for item in styles.findall(f"{_W}style"):
        style_id = item.get(f"{_W}styleId")
        style_type = item.get(f"{_W}type")
        names = item.findall(f"{_W}name")
        name = names[0].get(f"{_W}val") if len(names) == 1 else None
        if not style_id or not style_type or not name or style_id in result:
            _fail("ooxml_style_registry_invalid")
        result[style_id] = (style_type, name)
    return result


def _style_alias_ids(styles: ElementTree.Element) -> set[str]:
    return {
        style_id
        for item in styles.findall(f"{_W}style")
        if (style_id := item.get(f"{_W}styleId")) and item.find(f".//{_W}aliases") is not None
    }


def _val(parent: ElementTree.Element, name: str, *, required: bool = True) -> str | None:
    values = parent.findall(f"{_W}{name}")
    if len(values) > 1 or (required and len(values) != 1):
        _fail(f"ooxml_numbering_property_invalid:{name}")
    return None if not values else values[0].get(f"{_W}val")


def _expected_caption_fields(
    materialization: Mapping[str, object], derived_number: str, style_names: Mapping[int, str]
) -> list[tuple[str, str]]:
    field_format = _FIELD_FORMATS.get(str(materialization.get("number_format")))
    if field_format is None:
        _fail("ooxml_caption_field_format_invalid")
    result: list[tuple[str, str]] = []
    if materialization.get("type") == "chapter_seq":
        level = materialization.get("chapter_heading_level")
        if not isinstance(level, int) or isinstance(level, bool) or level not in style_names:
            _fail("ooxml_caption_chapter_style_invalid")
        result.append((f' STYLEREF "{style_names[level]}" \\n ', str(materialization.get("chapter_cached_number"))))
    elif materialization.get("type") != "simple_seq":
        _fail("ooxml_caption_materialization_invalid")
    counter = str(materialization.get("counter", ""))
    action = materialization.get("sequence_action")
    if action == "continue":
        instruction = f" SEQ {counter} \\* {field_format} "
    elif action == "reset_to_start":
        instruction = f" SEQ {counter} \\r {materialization.get('start_value')} \\* {field_format} "
    elif action == "restart_by_heading_level":
        instruction = f" SEQ {counter} \\s {materialization.get('restart_heading_level')} \\* {field_format} "
    else:
        _fail("ooxml_caption_sequence_action_invalid")
    sequence = str(materialization.get("sequence_cached_number", ""))
    if not sequence:
        _fail("ooxml_caption_sequence_cache_empty")
    if materialization.get("type") == "simple_seq" and sequence != derived_number:
        _fail("ooxml_caption_cache_contradicts_plan")
    result.append((instruction, sequence))
    return result


def inspect_resolved_docx(
    payload: bytes,
    expected: Mapping[str, int | bool],
    neutral_envelope: Mapping[str, object],
    plan_envelope: Mapping[str, object],
) -> dict[str, object]:
    """Prove the closed physical package against both resolved inputs."""

    try:
        with zipfile.ZipFile(io.BytesIO(payload)) as archive:
            infos = archive.infolist()
            names_in_order = [item.filename for item in infos]
            if len(names_in_order) != len(set(names_in_order)):
                _fail("ooxml_duplicate_zip_member")
            if len(infos) > 4096:
                _fail("ooxml_zip_member_limit_exceeded")
            for info in infos:
                if (
                    info.flag_bits & 0x1
                    or "\\" in info.filename
                    or ":" in info.filename
                    or info.filename.startswith("/")
                    or any(part in {"", ".", ".."} for part in info.filename.split("/"))
                    or info.file_size > 64 * 1024 * 1024
                ):
                    _fail("ooxml_zip_member_invalid")
            names = set(names_in_order)
            required = {
                "[Content_Types].xml",
                "_rels/.rels",
                "word/document.xml",
                "word/_rels/document.xml.rels",
                "word/styles.xml",
            }
            if not required <= names:
                _fail("ooxml_required_part_missing")
            document = _xml(archive.read("word/document.xml"), part="word/document.xml")
            styles = _xml(archive.read("word/styles.xml"), part="word/styles.xml")
            content_types = _xml(archive.read("[Content_Types].xml"), part="[Content_Types].xml")
            root_relationships = _xml(archive.read("_rels/.rels"), part="_rels/.rels")
            relationships = _xml(archive.read("word/_rels/document.xml.rels"), part="word/_rels/document.xml.rels")
            custom_parts = {name: archive.read(name) for name in names_in_order if name.startswith("customXml/")}
            numbering = (
                _xml(archive.read("word/numbering.xml"), part="word/numbering.xml")
                if "word/numbering.xml" in names
                else None
            )
    except (OSError, KeyError, zipfile.BadZipFile) as exc:
        raise V4OoxmlProofError("ooxml_package_invalid") from exc

    if document.tag != f"{_W}document" or styles.tag != f"{_W}styles":
        _fail("ooxml_word_root_invalid")
    neutral = _object(neutral_envelope.get("document"), "ooxml_neutral_document_invalid")
    plan = _object(plan_envelope.get("plan"), "ooxml_numbering_plan_invalid")
    source_sha256 = str(neutral_envelope.get("source_sha256", ""))
    if _HEX64.fullmatch(source_sha256) is None or source_sha256 != plan_envelope.get("source_sha256"):
        _fail("ooxml_source_identity_invalid")
    targets = [
        _object(item, "ooxml_neutral_target_invalid")
        for item in _array(neutral.get("targets"), "ooxml_neutral_targets_invalid")
    ]
    plan_targets = [
        _object(item, "ooxml_plan_target_invalid") for item in _array(plan.get("targets"), "ooxml_plan_targets_invalid")
    ]

    def target_keys(item: Mapping[str, object]) -> tuple[object, object, object, object]:
        return item.get("source_start"), item.get("source_end"), item.get("kind"), item.get("target_id")

    if [target_keys(item) for item in targets] != [target_keys(item) for item in plan_targets]:
        _fail("ooxml_target_inventory_mismatch")
    plan_sha256 = str(plan_envelope.get("plan_sha256", ""))
    if _HEX64.fullmatch(plan_sha256) is None or plan_sha256 != neutral_envelope.get("plan_sha256"):
        _fail("ooxml_plan_identity_invalid")
    parents = _parent_map(document)
    sdts = _sdt_index(document)
    bookmarks = _bookmarks(document, parents)
    styles_by_id = _style_index(styles)

    def physical_key(item: Mapping[str, object]) -> tuple[int, int, str]:
        start, end = item.get("source_start"), item.get("source_end")
        if not isinstance(start, int) or isinstance(start, bool) or not isinstance(end, int) or isinstance(end, bool):
            _fail("ooxml_target_source_identity_invalid")
        return start, end, str(item.get("kind"))

    target_tags: dict[tuple[int, int, str], str] = {}
    for target in targets:
        target_id = target.get("target_id")
        if target_id is None:
            continue
        if not isinstance(target_id, str) or not target_id:
            _fail("ooxml_target_id_invalid")
        target_tags[physical_key(target)] = _target_identity(str(target.get("kind")), target_id)[0]
    try:
        caption_maps = map_proof.inspect_caption_maps(
            custom_parts,
            relationships,
            content_types,
            targets,
            plan_targets,
            source_sha256,
            plan_sha256,
        )
        projection = target_proof.project_targets(
            document,
            numbering,
            targets,
            plan_targets,
            sdts,
            target_tags,
            caption_maps.styles_by_kind,
            caption_maps.occurrence_tags,
        )
    except (map_proof.V4OoxmlMapError, target_proof.V4OoxmlTargetError) as exc:
        raise V4OoxmlProofError(str(exc)) from exc
    for style_id, visible_name in caption_maps.styles_by_kind.values():
        if styles_by_id.get(style_id) != ("paragraph", visible_name):
            _fail("ooxml_caption_style_registry_mismatch")

    managed_style_ids = {style_id for style_id, _name in caption_maps.styles_by_kind.values()}
    for paragraph in projection.paragraphs.values():
        paragraph_style = paragraph.find(f"{_W}pPr/{_W}pStyle")
        if paragraph_style is not None:
            style_id = paragraph_style.get(f"{_W}val")
            if style_id:
                managed_style_ids.add(style_id)
    if numbering is not None:
        for level in numbering.findall(f"{_W}abstractNum/{_W}lvl"):
            style_id = _val(level, "pStyle")
            if style_id:
                managed_style_ids.add(style_id)
    if managed_style_ids & _style_alias_ids(styles):
        _fail("ooxml_managed_style_alias_rejected")

    heading_style_names: dict[int, str] = {}
    if numbering is not None:
        for node in numbering.findall(f"{_W}abstractNum/{_W}lvl"):
            raw_level = node.get(f"{_W}ilvl")
            style_id = _val(node, "pStyle")
            if raw_level is None or not raw_level.isdecimal() or not style_id or style_id not in styles_by_id:
                _fail("ooxml_heading_level_style_invalid")
            level = int(raw_level) + 1
            style_name = styles_by_id[style_id][1]
            if level in heading_style_names and heading_style_names[level] != style_name:
                _fail("ooxml_heading_level_style_ambiguous")
            heading_style_names[level] = style_name

    expected_names: set[str] = set()
    expected_fields: list[tuple[str, str, ElementTree.Element]] = []
    heading_bindings: list[tuple[dict[str, Any], dict[str, Any], ElementTree.Element, int, int]] = []
    for target, planned in zip(targets, plan_targets, strict=True):
        kind, target_id = str(target.get("kind")), target.get("target_id")
        key = physical_key(target)
        paragraph = projection.paragraphs[key]
        sdt = None if target_id is None else sdts.get(target_tags[key])
        if isinstance(target_id, str):
            tag, bookmark_name = _target_identity(kind, target_id)
            if sdt is None:
                _fail(f"ooxml_target_sdt_missing:{kind}:{target_id}")
            bookmark = bookmarks.get(bookmark_name)
            if bookmark is None:
                _fail(f"ooxml_target_bookmark_missing:{kind}:{target_id}")
            _prove_bookmark_scope(bookmark, sdt, parents)
            expected_names.add(bookmark_name)
        p_pr = paragraph.find(f"{_W}pPr")
        num_pr = None if p_pr is None else p_pr.find(f"{_W}numPr")
        enabled = planned.get("enabled") is True
        materialization = planned.get("materialization")
        derived_number = planned.get("derived_number")
        if kind == "heading":
            authored_heading = str(target.get("authored_text", ""))
            if not authored_heading or not _visible_text(paragraph).startswith(authored_heading):
                _fail("ooxml_heading_authored_text_changed")
            if enabled:
                materialized = _object(materialization, "ooxml_heading_materialization_missing")
                if materialized.get("type") != "heading_list" or num_pr is None:
                    _fail("ooxml_heading_numpr_missing")
                ilvl = _val(num_pr, "ilvl")
                num_id = _val(num_pr, "numId")
                if ilvl is None or num_id is None or not ilvl.isdecimal() or not num_id.isdecimal():
                    _fail("ooxml_heading_numpr_invalid")
                level = int(materialized.get("level", 0))
                if int(ilvl) != level - 1 or int(num_id) < 1:
                    _fail("ooxml_heading_numpr_plan_mismatch")
                heading_bindings.append((planned, materialized, paragraph, int(num_id), level))
            elif num_pr is not None or materialization is not None or derived_number is not None:
                _fail("ooxml_disabled_heading_materialized")
        elif kind in _CAPTION_KINDS:
            if num_pr is not None:
                _fail("ooxml_caption_has_numpr")
            style_id = _val(p_pr, "pStyle") if p_pr is not None else None
            expected_style = caption_maps.styles_by_kind.get(kind, (None, ""))[0]
            if style_id is None or style_id != expected_style:
                _fail("ooxml_caption_managed_style_missing")
            if enabled:
                materialized = _object(materialization, "ooxml_caption_materialization_missing")
                number = str(derived_number or "")
                if not number:
                    _fail("ooxml_caption_derived_number_missing")
                caption_pairs = _expected_caption_fields(materialized, number, heading_style_names)
                expected_fields.extend((instruction, cached, paragraph) for instruction, cached in caption_pairs)
                label = f"{materialized.get('localized_label', '')}{materialized.get('label_separator', '')}"
                if not _visible_text(paragraph).startswith(label):
                    _fail("ooxml_caption_label_missing")
                authored = str(target.get("authored_text", ""))
                if authored and authored not in _visible_text(paragraph):
                    _fail("ooxml_caption_authored_text_changed")
            elif materialization is not None or derived_number is not None:
                _fail("ooxml_disabled_caption_materialized")
        else:
            _fail("ooxml_target_kind_invalid")

    definitions = [
        _object(item, "ooxml_heading_definition_invalid")
        for item in _array(plan.get("heading_definitions"), "ooxml_heading_definitions_invalid")
    ]
    instances = [
        _object(item, "ooxml_heading_instance_invalid")
        for item in _array(plan.get("heading_instances"), "ooxml_heading_instances_invalid")
    ]
    try:
        opc_proof.prove_opc(
            set(names_in_order),
            content_types,
            root_relationships,
            relationships,
            numbering_present=bool(heading_bindings),
        )
    except opc_proof.V4OoxmlOpcError as exc:
        raise V4OoxmlProofError(str(exc)) from exc
    try:
        numbering_proof.prove_heading_numbering(
            numbering,
            definitions,
            instances,
            heading_bindings,
            styles_by_id,
        )
    except numbering_proof.V4OoxmlNumberingError as exc:
        raise V4OoxmlProofError(str(exc)) from exc

    references = [
        _object(item, "ooxml_reference_invalid")
        for item in _array(neutral.get("references"), "ooxml_references_invalid")
    ]
    target_bookmarks = {
        (str(item.get("kind")), str(item.get("target_id"))): _target_identity(
            str(item.get("kind")), str(item.get("target_id"))
        )[1]
        for item in targets
        if item.get("target_id")
    }
    managed_paragraphs = set(projection.paragraphs.values())
    reference_tags: set[str] = set()
    for reference in references:
        key = (str(reference.get("target_kind")), str(reference.get("target_id")))
        bookmark = target_bookmarks.get(key)
        if bookmark is None:
            _fail("ooxml_reference_target_missing")
        tag = _reference_tag(reference, bookmark, source_sha256)
        sdt = sdts.get(tag)
        if sdt is None:
            _fail("ooxml_reference_sdt_missing")
        reference_tags.add(tag)
        paragraph = _carrier_paragraph(sdt, parents, code="ooxml_reference_carrier_invalid")
        managed_paragraphs.add(paragraph)
        instruction = f" REF {bookmark} \\n \\h " if key[0] == "heading" else f" REF {bookmark} \\h "
        expected_fields.append((instruction, str(reference.get("cached_number")), paragraph))
        alias = reference.get("alias")
        if alias is not None and f" {alias}" not in _visible_text(sdt):
            _fail("ooxml_reference_alias_missing")

    if {tag for tag in sdts if tag.startswith("docwen-ref-occurrence-v1:")} != reference_tags:
        _fail("ooxml_reference_sdt_inventory_invalid")

    citations = [
        _object(item, "ooxml_citation_invalid") for item in _array(neutral.get("citations"), "ooxml_citations_invalid")
    ]
    citation_tags: set[str] = set()
    for citation in citations:
        tag, bookmark_name, instruction = _citation_identity(citation, source_sha256)
        sdt = sdts.get(tag)
        bookmark = bookmarks.get(bookmark_name)
        if sdt is None or bookmark is None:
            _fail("ooxml_citation_carrier_missing")
        _prove_bookmark_scope(bookmark, sdt, parents)
        expected_names.add(bookmark_name)
        citation_tags.add(tag)
        paragraph = _carrier_paragraph(sdt, parents, code="ooxml_citation_carrier_invalid")
        managed_paragraphs.add(paragraph)
        expected_fields.append((instruction, str(citation.get("cached_result")), paragraph))

    if {tag for tag in sdts if tag.startswith("docwen-citation-occurrence-v1:")} != citation_tags:
        _fail("ooxml_citation_sdt_inventory_invalid")

    managed_fields = list(_fields(document, managed_paragraphs))
    actual_fields = Counter((item.instruction, item.cached_result, item.paragraph) for item in managed_fields)
    if actual_fields != Counter(expected_fields):
        _fail("ooxml_field_inventory_or_cache_mismatch")
    owned_counters = {
        str(materialization.get("counter"))
        for planned in plan_targets
        if isinstance((materialization := planned.get("materialization")), dict) and materialization.get("counter")
    }
    for paragraph in document.iter(f"{_W}p"):
        if paragraph in managed_paragraphs:
            continue
        instruction = "".join(item.text or "" for item in paragraph.iter(f"{_W}instrText"))
        if (
            any(f" REF {name} " in instruction for name in expected_names)
            or re.search(r"(?:^|\s)REF\s+(?:DW_T_|_DWC_)", instruction, re.IGNORECASE) is not None
            or re.search(r"(?:^|\s)CITATION\s+DWCIT_", instruction, re.IGNORECASE) is not None
            or any(f" SEQ {counter} " in instruction for counter in owned_counters)
        ):
            _fail("ooxml_owned_field_outside_managed_carrier")
    if set(bookmarks) != expected_names:
        _fail("ooxml_bookmark_inventory_not_exact")

    actual_counts = {
        "abstractNumCount": 0 if numbering is None else len(numbering.findall(f"{_W}abstractNum")),
        "numCount": 0 if numbering is None else len(numbering.findall(f"{_W}num")),
        "bookmarkCount": len(bookmarks),
        "seqFieldCount": sum(item.instruction.startswith(" SEQ ") for item in managed_fields),
        "styleRefFieldCount": sum(item.instruction.startswith(" STYLEREF ") for item in managed_fields),
        "refFieldCount": sum(item.instruction.startswith(" REF ") for item in managed_fields),
        "citationFieldCount": sum(item.instruction.startswith(" CITATION ") for item in managed_fields),
    }
    if any(actual_counts.get(key) != expected.get(key) for key in actual_counts):
        _fail("ooxml_expected_count_mismatch")
    return {
        "bookmarkCount": actual_counts["bookmarkCount"],
        "seqFieldCount": actual_counts["seqFieldCount"],
        "refFieldCount": actual_counts["refFieldCount"],
        "violations": [],
    }
