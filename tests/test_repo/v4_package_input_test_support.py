from __future__ import annotations

import hashlib
import io
import json
import uuid
import zipfile
from collections.abc import Mapping
from typing import Any, cast
from xml.etree import ElementTree

from scripts.release import v4_package_input_ooxml_identity as identity

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CUSTOM_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
CAPTION_STYLE_NAMESPACE = "https://docwen.dev/schema/document-caption-style-binding-map/v1"
OCCURRENCE_NAMESPACE = "https://docwen.dev/schema/document-numbering-occurrence-map/v1"
MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_W = f"{{{WML}}}"
_R = f"{{{RELS}}}"
_CT = f"{{{CONTENT_TYPES}}}"

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
_FIELD_FORMATS = {
    "arabic_half": "ARABIC",
    "letter_upper": "ALPHABETIC",
    "letter_lower": "alphabetic",
    "roman_upper": "ROMAN",
    "roman_lower": "roman",
}
_CAPTION_LABEL = {"figure": "Figure", "table": "Table", "equation": "Equation", "code_block": "Code"}


def _sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def _json_bytes(value: object) -> bytes:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode()


def _target(kind: str, target_id: str, authored: str, start: int) -> dict[str, object]:
    return {
        "source_start": start,
        "source_end": start + 10,
        "source_slice_sha256": _sha(f"{kind}:{target_id}".encode()),
        "kind": kind,
        "target_id": target_id,
        "heading_level": 1 if kind == "heading" else None,
        "authored_text": authored,
    }


def _planned(target: Mapping[str, object], *, enabled: bool) -> dict[str, object]:
    kind = str(target["kind"])
    materialization: dict[str, object] | None
    if not enabled:
        materialization = None
    elif kind == "heading":
        materialization = {
            "definition_id": "heading-default",
            "instance_id": "heading-instance-1",
            "level": 1,
            "type": "heading_list",
        }
    else:
        label = _CAPTION_LABEL[kind]
        materialization = {
            "chapter_cached_number": None,
            "chapter_heading_level": None,
            "chapter_heading_style": None,
            "chapter_separator": None,
            "counter": label,
            "label_separator": " ",
            "localized_label": label,
            "number_format": "arabic_half",
            "restart_heading_level": None,
            "restart_heading_style": None,
            "sequence_action": "continue",
            "sequence_cached_number": "1",
            "start_value": None,
            "type": "simple_seq",
        }
    return {
        "derived_number": "1" if enabled else None,
        "enabled": enabled,
        "kind": kind,
        "materialization": materialization,
        "source_end": target["source_end"],
        "source_start": target["source_start"],
        "target_id": target["target_id"],
    }


def _reference(target: Mapping[str, object], *, start: int, alias: str | None = None) -> dict[str, object]:
    token = f"@[[#^{target['target_id']}]]" if alias is None else f"@[[#^{target['target_id']}|{alias}]]"
    return {
        "source_start": start,
        "source_end": start + len(token),
        "source_slice_sha256": _sha(token.encode()),
        "authored_token": token,
        "target_source_start": target["source_start"],
        "target_source_end": target["source_end"],
        "target_kind": target["kind"],
        "target_id": target["target_id"],
        "cached_number": "1",
        "alias": alias,
    }


def case_semantics(case_id: str, source: str) -> tuple[dict[str, Any], dict[str, Any], dict[str, int]]:
    targets: list[dict[str, object]] = []
    references: list[dict[str, object]] = []
    citations: list[dict[str, object]] = []
    if case_id == "rich-semantics-composite":
        for index, (kind, target_id, authored) in enumerate(
            (
                ("heading", "heading-one", "2.3 Authored heading"),
                ("figure", "figure-one", "Caption"),
                ("table", "table-one", "Caption"),
                ("equation", "equation-one", "Caption"),
                ("code_block", "code-one", "Caption"),
            )
        ):
            targets.append(_target(kind, target_id, authored, index * 20))
        references = [_reference(targets[0], start=200), _reference(targets[1], start=230, alias="Figure alias")]
        citation_token = "@cite-one"
        citations = [
            {
                "source_start": 260,
                "source_end": 260 + len(citation_token),
                "source_slice_sha256": _sha(citation_token.encode()),
                "authored_token": citation_token,
                "form": "narrative",
                "cluster_id": "citation-one",
                "items": [
                    {
                        "citation_key": "cite-one",
                        "record_id": "ref-one",
                        "record_sha256": _sha(b"reference-one"),
                        "presentation": "One (2026)",
                    }
                ],
                "cached_result": "One (2026)",
            }
        ]
        enabled = [True] * 5
    else:
        kind = "heading"
        for candidate in _CAPTION_LABEL:
            if f"numbering-{candidate.removesuffix('_block')}" in case_id:
                kind = candidate
        target_id = f"{kind.removesuffix('_block')}-one"
        authored = "2.3 Authored heading" if kind == "heading" else "Caption"
        targets = [_target(kind, target_id, authored, 0)]
        is_on = case_id.endswith("-on") or case_id == "heading-authored-number-preserved"
        enabled = [is_on]
        if is_on:
            references = [_reference(targets[0], start=100)]
    planned_targets = [_planned(target, enabled=value) for target, value in zip(targets, enabled, strict=True)]
    heading_enabled = any(item["kind"] == "heading" and item["enabled"] for item in planned_targets)
    definitions = (
        [
            {
                "definition_id": "heading-default",
                "levels": [
                    {
                        "display": [{"counter": {"level": 1, "number_format": "arabic_half"}}],
                        "level": 1,
                        "number_format": "arabic_half",
                        "restart_after_level": None,
                        "start": 1,
                        "suffix": "space",
                    }
                ],
            }
        ]
        if heading_enabled
        else []
    )
    instances = (
        [{"definition_id": "heading-default", "instance_id": "heading-instance-1", "starts": []}]
        if heading_enabled
        else []
    )
    plan_member = {"heading_definitions": definitions, "heading_instances": instances, "targets": planned_targets}
    source_raw = source.encode()
    source_sha = _sha(source_raw)
    plan_sha = _sha(_json_bytes(plan_member))
    shared = {
        "input_id": "positive-exact-two.md",
        "source_sha256": source_sha,
        "plan_sha256": plan_sha,
    }
    neutral: dict[str, Any] = {
        "$schema": "urn:docwen:schema:resolved-document:v1",
        "schema": "docwen.resolved_document.v1",
        **shared,
        "document": {
            "authored_markdown": source,
            "targets": targets,
            "references": references,
            "resource_occurrences": [],
            "citations": citations,
            "resources": [],
        },
    }
    plan: dict[str, Any] = {
        "$schema": "urn:docwen:schema:numbering-export-plan:v1",
        "schema": "docwen.numbering_export_plan.v1",
        **shared,
        "plan": plan_member,
    }
    expected = {
        "abstractNumCount": len(definitions),
        "numCount": len(instances),
        "bookmarkCount": len(targets) + len(citations),
        "seqFieldCount": sum(item["enabled"] and item["kind"] != "heading" for item in planned_targets),
        "styleRefFieldCount": 0,
        "refFieldCount": len(references),
        "citationFieldCount": len(citations),
    }
    return neutral, plan, cast(dict[str, int], expected)


def _sub(parent: ElementTree.Element, tag_name: str, **attributes: str) -> ElementTree.Element:
    return ElementTree.SubElement(
        parent,
        f"{_W}{tag_name}",
        {f"{_W}{key}": value for key, value in attributes.items()},
    )


def _text(parent: ElementTree.Element, value: str) -> None:
    run = _sub(parent, "r")
    node = _sub(run, "t")
    node.text = value


def _field(parent: ElementTree.Element, instruction: str, cached: str) -> None:
    begin = _sub(_sub(parent, "r"), "fldChar", fldCharType="begin")
    begin.set(f"{_W}dirty", "false")
    instruction_node = _sub(_sub(parent, "r"), "instrText")
    instruction_node.text = instruction
    _sub(_sub(parent, "r"), "fldChar", fldCharType="separate")
    _text(parent, cached)
    _sub(_sub(parent, "r"), "fldChar", fldCharType="end")


def _sdt(body: ElementTree.Element, tag: str) -> tuple[ElementTree.Element, ElementTree.Element]:
    sdt = _sub(body, "sdt")
    properties = _sub(sdt, "sdtPr")
    _sub(properties, "tag", val=tag)
    paragraph = _sub(_sub(sdt, "sdtContent"), "p")
    return sdt, paragraph


def _block_sdt(body: ElementTree.Element, tag: str) -> tuple[ElementTree.Element, ElementTree.Element]:
    sdt = _sub(body, "sdt")
    properties = _sub(sdt, "sdtPr")
    _sub(properties, "tag", val=tag)
    return sdt, _sub(sdt, "sdtContent")


def _logical_object(parent: ElementTree.Element, kind: str) -> ElementTree.Element:
    if kind == "table":
        table = _sub(parent, "tbl")
        _sub(_sub(_sub(table, "tr"), "tc"), "p")
        return table
    paragraph = _sub(parent, "p")
    if kind == "figure":
        _sub(_sub(paragraph, "r"), "drawing")
    elif kind == "equation":
        ElementTree.SubElement(paragraph, f"{{{MATH}}}oMath")
    elif kind == "code_block":
        _text(paragraph, "code")
    else:
        raise AssertionError(kind)
    return paragraph


def _occurrence_identity(target: Mapping[str, object], source_sha256: str, plan_sha256: str) -> tuple[str, str]:
    preimage = (
        "docwen-numbering-occurrence-map-v1\0"
        f"{source_sha256}\0{target['source_start']}\0{target['source_end']}\0{target['kind']}"
        f"\0false\0\0\0{plan_sha256}"
    )
    digest = _sha(preimage.encode())
    return f"docwen-numbering-occurrence-v1:{digest[:32]}", digest


def _caption_style_map() -> bytes:
    records = "".join(
        f'<binding semantic_key="{kind}_caption" resolved_style_id="Caption{kind.title().replace("_", "")}" '
        f'visible_name="DocWen {kind} caption"/>'
        for kind in _CAPTION_LABEL
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<documentCaptionStyleBindingMap xmlns="{CAPTION_STYLE_NAMESPACE}" version="1">'
        f"{records}</documentCaptionStyleBindingMap>\n"
    ).encode()


def _occurrence_map(
    records: list[tuple[Mapping[str, object], str, str]], source_sha256: str, plan_sha256: str
) -> bytes:
    entries = "".join(
        f'<occurrence tag="{tag}" source_sha256="{source_sha256}" '
        f'source_start="{target["source_start"]}" source_end="{target["source_end"]}" '
        f'kind="{target["kind"]}" enabled="false" target_id="" derived_number="" '
        f'plan_sha256="{plan_sha256}" sha256="{digest}"/>'
        for target, tag, digest in records
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<documentNumberingOccurrenceMap xmlns="{OCCURRENCE_NAMESPACE}" version="1" '
        f'plan_sha256="{plan_sha256}">{entries}</documentNumberingOccurrenceMap>\n'
    ).encode()


def _custom_xml_parts(
    maps: list[tuple[str, bytes]],
    content_types: ElementTree.Element,
    document_rels: ElementTree.Element,
) -> dict[str, bytes]:
    result: dict[str, bytes] = {}
    for number, (namespace, raw) in enumerate(maps, start=1):
        item_name = f"customXml/item{number}.xml"
        props_name = f"customXml/itemProps{number}.xml"
        item_rels_name = f"customXml/_rels/item{number}.xml.rels"
        item_uuid = "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, f"{namespace}\0{_sha(raw)}")).upper() + "}"
        props = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<ds:datastoreItem xmlns:ds="{CUSTOM_PROPS}" ds:itemID="{item_uuid}">'
            f'<ds:schemaRefs><ds:schemaRef ds:uri="{namespace}"/></ds:schemaRefs></ds:datastoreItem>\n'
        ).encode()
        item_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<Relationships xmlns="{RELS}"><Relationship Id="rId1" '
            f'Type="{OFFICE_RELS}/customXmlProps" Target="itemProps{number}.xml"/></Relationships>\n'
        ).encode()
        ElementTree.SubElement(
            document_rels,
            f"{_R}Relationship",
            Id=f"rIdCustom{number}",
            Type=f"{OFFICE_RELS}/customXml",
            Target=f"../{item_name}",
        )
        for part_name, content_type in (
            (f"/{item_name}", "application/xml"),
            (
                f"/{props_name}",
                "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
            ),
        ):
            ElementTree.SubElement(
                content_types,
                f"{_CT}Override",
                PartName=part_name,
                ContentType=content_type,
            )
        result.update({item_name: raw, props_name: props, item_rels_name: item_rels})
    return result


def _bookmark(
    paragraph: ElementTree.Element, bookmark_id: int, name: str
) -> tuple[ElementTree.Element, ElementTree.Element]:
    return _sub(paragraph, "bookmarkStart", id=str(bookmark_id), name=name), _sub(
        paragraph, "bookmarkEnd", id=str(bookmark_id)
    )


def _paragraph_style(paragraph: ElementTree.Element, style_id: str) -> ElementTree.Element:
    properties = _sub(paragraph, "pPr")
    _sub(properties, "pStyle", val=style_id)
    return properties


def _caption_field(materialized: Mapping[str, object]) -> str:
    field_format = _FIELD_FORMATS[str(materialized["number_format"])]
    action = materialized["sequence_action"]
    if action == "continue":
        return f" SEQ {materialized['counter']} \\* {field_format} "
    if action == "reset_to_start":
        return f" SEQ {materialized['counter']} \\r {materialized['start_value']} \\* {field_format} "
    return f" SEQ {materialized['counter']} \\s {materialized['restart_heading_level']} \\* {field_format} "


def _display(level: Mapping[str, object]) -> str:
    result = ""
    for segment in cast(list[dict[str, object]], level["display"]):
        result += str(segment.get("literal", f"%{cast(dict[str, object], segment['counter'])['level']}"))
    return result


def build_docx(neutral_envelope: Mapping[str, object], plan_envelope: Mapping[str, object]) -> bytes:
    document_data = cast(dict[str, Any], neutral_envelope["document"])
    plan_data = cast(dict[str, Any], plan_envelope["plan"])
    targets = cast(list[dict[str, object]], document_data["targets"])
    planned = cast(list[dict[str, object]], plan_data["targets"])
    document = ElementTree.Element(f"{_W}document")
    body = _sub(document, "body")
    styles_needed: dict[str, str] = {"Heading1": "Heading 1"}
    if any(target["kind"] in _CAPTION_LABEL for target in targets):
        styles_needed.update(
            {f"Caption{kind.title().replace('_', '')}": f"DocWen {kind} caption" for kind in _CAPTION_LABEL}
        )
    source_sha = str(neutral_envelope["source_sha256"])
    plan_sha = str(plan_envelope["plan_sha256"])
    occurrence_records: list[tuple[Mapping[str, object], str, str]] = []
    bookmark_id = 0
    for target, target_plan in zip(targets, planned, strict=True):
        kind = str(target["kind"])
        target_id = target.get("target_id")
        bookmark_name: str | None = None
        if kind == "heading":
            if target_id is None:
                paragraph = _sub(body, "p")
            else:
                tag, bookmark_name = identity.target_identity(kind, str(target_id))
                _carrier, paragraph = _sdt(body, tag)
            properties = _paragraph_style(paragraph, "Heading1")
            if target_plan["enabled"]:
                num_properties = _sub(properties, "numPr")
                _sub(num_properties, "ilvl", val="0")
                _sub(num_properties, "numId", val="1")
            if bookmark_name is None:
                _text(paragraph, str(target["authored_text"]))
            else:
                _start, end = _bookmark(paragraph, bookmark_id, bookmark_name)
                paragraph.remove(end)
                _text(paragraph, str(target["authored_text"]))
                paragraph.append(end)
                bookmark_id += 1
        else:
            style_id = f"Caption{kind.title().replace('_', '')}"
            if target_id is not None:
                tag, bookmark_name = identity.target_identity(kind, str(target_id))
                _carrier, container = _block_sdt(body, tag)
            elif target_plan["enabled"]:
                container = body
            else:
                tag, digest = _occurrence_identity(target, source_sha, plan_sha)
                occurrence_records.append((target, tag, digest))
                _carrier, container = _block_sdt(body, tag)
            if kind == "figure":
                _logical_object(container, kind)
                paragraph = _sub(container, "p")
            else:
                paragraph = _sub(container, "p")
                _logical_object(container, kind)
            _paragraph_style(paragraph, style_id)
            end = None
            if bookmark_name is not None:
                _start, end = _bookmark(paragraph, bookmark_id, bookmark_name)
                paragraph.remove(end)
            if target_plan["enabled"]:
                materialized = cast(dict[str, object], target_plan["materialization"])
                _text(paragraph, f"{materialized['localized_label']}{materialized['label_separator']}")
                _field(paragraph, _caption_field(materialized), str(materialized["sequence_cached_number"]))
                _text(paragraph, f" {target['authored_text']}")
            else:
                _text(paragraph, str(target["authored_text"]))
            if end is not None:
                paragraph.append(end)
                bookmark_id += 1
    target_by_key = {(str(item["kind"]), str(item["target_id"])): item for item in targets}
    for reference in cast(list[dict[str, object]], document_data["references"]):
        key = (str(reference["target_kind"]), str(reference["target_id"]))
        target = target_by_key[key]
        bookmark_name = identity.target_identity(str(target["kind"]), str(target["target_id"]))[1]
        tag = identity.reference_tag(reference, bookmark_name, source_sha)
        carrier, paragraph = _sdt(body, tag)
        instruction = f" REF {bookmark_name} \\n \\h " if key[0] == "heading" else f" REF {bookmark_name} \\h "
        _field(paragraph, instruction, str(reference["cached_number"]))
        if reference.get("alias") is not None:
            _text(paragraph, f" {reference['alias']}")
        del carrier
    for citation in cast(list[dict[str, object]], document_data["citations"]):
        tag, bookmark_name, instruction = identity.citation_identity(citation, source_sha)
        _carrier, paragraph = _sdt(body, tag)
        _start, end = _bookmark(paragraph, bookmark_id, bookmark_name)
        paragraph.remove(end)
        _field(paragraph, instruction, str(citation["cached_result"]))
        paragraph.append(end)
        bookmark_id += 1
    styles = ElementTree.Element(f"{_W}styles")
    for style_id, name in styles_needed.items():
        style = _sub(styles, "style", type="paragraph", styleId=style_id)
        _sub(style, "name", val=name)
    definitions = cast(list[dict[str, Any]], plan_data["heading_definitions"])
    instances = cast(list[dict[str, Any]], plan_data["heading_instances"])
    numbering: ElementTree.Element | None = None
    if definitions:
        numbering = ElementTree.Element(f"{_W}numbering")
        for abstract_id, definition in enumerate(definitions):
            abstract = _sub(numbering, "abstractNum", abstractNumId=str(abstract_id))
            _sub(abstract, "multiLevelType", val="multilevel")
            for level in definition["levels"]:
                node = _sub(abstract, "lvl", ilvl=str(int(level["level"]) - 1))
                _sub(node, "start", val=str(level["start"]))
                _sub(node, "numFmt", val=_NUMBER_FORMATS[str(level["number_format"])])
                _sub(
                    node,
                    "lvlRestart",
                    val="0" if level["restart_after_level"] is None else str(level["restart_after_level"]),
                )
                _sub(node, "suff", val=str(level["suffix"]))
                _sub(node, "lvlText", val=_display(level))
                _sub(node, "pStyle", val=f"Heading{level['level']}")
        definition_ids = {item["definition_id"]: index for index, item in enumerate(definitions)}
        for num_id, instance in enumerate(instances, start=1):
            num = _sub(numbering, "num", numId=str(num_id))
            _sub(num, "abstractNumId", val=str(definition_ids[instance["definition_id"]]))
            for start_value in instance["starts"]:
                override = _sub(num, "lvlOverride", ilvl=str(int(start_value["level"]) - 1))
                _sub(override, "startOverride", val=str(start_value["value"]))
    content_types = ElementTree.Element(f"{_CT}Types")
    ElementTree.SubElement(
        content_types,
        f"{_CT}Default",
        Extension="rels",
        ContentType="application/vnd.openxmlformats-package.relationships+xml",
    )
    ElementTree.SubElement(content_types, f"{_CT}Default", Extension="xml", ContentType="application/xml")
    ElementTree.SubElement(
        content_types,
        f"{_CT}Override",
        PartName="/word/document.xml",
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    )
    ElementTree.SubElement(
        content_types,
        f"{_CT}Override",
        PartName="/word/styles.xml",
        ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    )
    document_rels = ElementTree.Element(f"{_R}Relationships")
    ElementTree.SubElement(
        document_rels,
        f"{_R}Relationship",
        Id="rIdStyles",
        Type=f"{OFFICE_RELS}/styles",
        Target="styles.xml",
    )
    if numbering is not None:
        ElementTree.SubElement(
            content_types,
            f"{_CT}Override",
            PartName="/word/numbering.xml",
            ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        )
        ElementTree.SubElement(
            document_rels,
            f"{_R}Relationship",
            Id="rIdNumbering",
            Type=f"{OFFICE_RELS}/numbering",
            Target="numbering.xml",
        )
    maps: list[tuple[str, bytes]] = []
    if any(target["kind"] in _CAPTION_LABEL for target in targets):
        maps.append((CAPTION_STYLE_NAMESPACE, _caption_style_map()))
    if occurrence_records:
        occurrence_records.sort(key=lambda item: (item[0]["source_start"], item[0]["source_end"], item[0]["kind"]))
        maps.append((OCCURRENCE_NAMESPACE, _occurrence_map(occurrence_records, source_sha, plan_sha)))
    custom_parts = _custom_xml_parts(maps, content_types, document_rels)
    root_rels = ElementTree.Element(f"{_R}Relationships")
    ElementTree.SubElement(
        root_rels,
        f"{_R}Relationship",
        Id="rIdDocument",
        Type=f"{OFFICE_RELS}/officeDocument",
        Target="word/document.xml",
    )
    parts = {
        "[Content_Types].xml": ElementTree.tostring(content_types, encoding="utf-8", xml_declaration=True),
        "_rels/.rels": ElementTree.tostring(root_rels, encoding="utf-8", xml_declaration=True),
        "word/document.xml": ElementTree.tostring(document, encoding="utf-8", xml_declaration=True),
        "word/_rels/document.xml.rels": ElementTree.tostring(document_rels, encoding="utf-8", xml_declaration=True),
        "word/styles.xml": ElementTree.tostring(styles, encoding="utf-8", xml_declaration=True),
    }
    if numbering is not None:
        parts["word/numbering.xml"] = ElementTree.tostring(numbering, encoding="utf-8", xml_declaration=True)
    parts.update(custom_parts)
    target = io.BytesIO()
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, raw in parts.items():
            archive.writestr(name, raw)
    return target.getvalue()


def rewrite_docx(payload: bytes, mutation: Any) -> bytes:
    source = zipfile.ZipFile(io.BytesIO(payload))
    parts = {item.filename: source.read(item) for item in source.infolist()}
    source.close()
    mutation(parts)
    target = io.BytesIO()
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, raw in parts.items():
            archive.writestr(name, raw)
    return target.getvalue()
