"""Closed v4 resolved-citation WordprocessingML projection and proof."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import lxml.etree as etree

from docwen_core._docx_citation_authority import (
    _WORD_TAG_TOKEN_RE,
    CITATION_ITEM_MAP_NAMESPACE,
    CITATION_OCCURRENCE_MAP_NAMESPACE,
    CitationItemIdentity,
    CitationItemMap,
    CitationItemReferenceIdentity,
    CitationOccurrenceIdentity,
    CitationOccurrenceMap,
    ResolvedCitationOoxmlError,
    ResolvedCitationProjection,
    _validate_item_refs,
    _validate_single_occurrence,
    build_resolved_citation_projection,
    citation_item_map_xml,
    citation_occurrence_map_xml,
    derive_citation_item,
    derive_citation_item_reference,
    derive_citation_occurrence,
    parse_citation_item_map,
    parse_citation_occurrence_map,
    validate_citation_authorities,
    validate_citation_item_map,
    validate_citation_occurrence_map,
)
from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error
from docwen_core.docx_bookmarks import (
    BOOKMARK_ID_MAX,
    DocxBookmarkInventory,
    bookmark_id_key,
    build_docx_bookmark_inventory,
    prove_bookmark_name,
)
from docwen_core.models.resolved_numbering import ResolvedCitation, ResolvedCitationItem

_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def preflight_citation_document(document: Any, projection: ResolvedCitationProjection) -> None:
    """Reject template collisions before adding any citation-owned XML."""

    from docx.oxml.ns import qn

    validate_citation_authorities(projection.item_map, projection.occurrence_map)
    inventory = build_docx_bookmark_inventory(document)
    planned_bookmarks = {item.bookmark_name.casefold() for item in projection.occurrence_map.occurrences}
    if planned_bookmarks.intersection(inventory.used_name_keys):
        raise ResolvedCitationOoxmlError("citation bookmark collides with an existing package bookmark")
    planned_sdt_tags = {item.tag for item in projection.occurrence_map.occurrences}
    existing_sdt_tags = {
        item.get(qn("w:val"))
        for root in _word_xml_roots(document)
        for item in root.iter(qn("w:tag"))
        if item.get(qn("w:val")) is not None
    }
    if planned_sdt_tags.intersection(existing_sdt_tags):
        raise ResolvedCitationOoxmlError("citation occurrence tag collides with an existing SDT")
    planned_word_tags = {item.word_tag.casefold() for item in projection.item_map.items}
    for instruction in _document_instructions(document):
        if not _is_citation_instruction(instruction):
            continue
        if planned_word_tags.intersection(item.casefold() for item in _WORD_TAG_TOKEN_RE.findall(instruction)):
            raise ResolvedCitationOoxmlError("citation Word tag collides with an existing non-owned field")


def citation_occurrence_sdt(identity: CitationOccurrenceIdentity, *, bookmark_id: str) -> Any:
    """Create one canonical inline SDT with a locked, clean CITATION field."""

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    _validate_single_occurrence(identity)
    if (
        re.fullmatch(r"0|[1-9][0-9]*", bookmark_id) is None
        or bookmark_id_key(bookmark_id) != ("numeric", int(bookmark_id))
        or not 0 <= int(bookmark_id) <= BOOKMARK_ID_MAX
    ):
        raise ResolvedCitationOoxmlError("citation bookmark ID is not a portable canonical decimal")
    sdt = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    tag = OxmlElement("w:tag")
    tag.set(qn("w:val"), identity.tag)
    properties.append(tag)
    content = OxmlElement("w:sdtContent")
    bookmark_start = OxmlElement("w:bookmarkStart")
    bookmark_start.set(qn("w:id"), bookmark_id)
    bookmark_start.set(qn("w:name"), identity.bookmark_name)
    content.append(bookmark_start)
    _append_field_run(content, "begin", locked=True)
    instruction_run = OxmlElement("w:r")
    instruction_text = OxmlElement("w:instrText")
    instruction_text.set(_XML_SPACE, "preserve")
    instruction_text.text = citation_instruction(identity.item_refs)
    instruction_run.append(instruction_text)
    content.append(instruction_run)
    _append_field_run(content, "separate")
    cached_run = OxmlElement("w:r")
    cached_text = OxmlElement("w:t")
    if identity.cached_result[:1].isspace() or identity.cached_result[-1:].isspace():
        cached_text.set(_XML_SPACE, "preserve")
    cached_text.text = identity.cached_result
    cached_run.append(cached_text)
    content.append(cached_run)
    _append_field_run(content, "end")
    bookmark_end = OxmlElement("w:bookmarkEnd")
    bookmark_end.set(qn("w:id"), bookmark_id)
    content.append(bookmark_end)
    sdt.extend((properties, content))
    return sdt


def citation_instruction(item_refs: tuple[CitationItemReferenceIdentity, ...]) -> str:
    _validate_item_refs(item_refs)
    tokens = ["CITATION", item_refs[0].word_tag]
    for item in item_refs[1:]:
        tokens.extend((r"\m", item.word_tag))
    return f" {' '.join(tokens)} "


def prove_citation_occurrence_sdt(
    sdt: Any,
    identity: CitationOccurrenceIdentity,
    *,
    bookmark_inventory: DocxBookmarkInventory,
) -> None:
    """Authenticate one map-bound direct-paragraph SDT and its field cache."""

    from docx.oxml.ns import qn

    _validate_single_occurrence(identity)
    parent = sdt.getparent()
    properties = sdt.find(qn("w:sdtPr"))
    content = sdt.find(qn("w:sdtContent"))
    if (
        parent is None
        or parent.tag != qn("w:p")
        or [item.tag for item in sdt] != [qn("w:sdtPr"), qn("w:sdtContent")]
        or properties is None
        or content is None
        or properties.text is not None
        or properties.tail is not None
        or content.text is not None
        or content.tail is not None
    ):
        raise ResolvedCitationOoxmlError("citation occurrence SDT envelope is not canonical")
    property_children = list(properties)
    if (
        len(property_children) != 1
        or property_children[0].tag != qn("w:tag")
        or tuple(property_children[0].attrib) != (qn("w:val"),)
        or property_children[0].get(qn("w:val")) != identity.tag
        or property_children[0].text is not None
        or property_children[0].tail is not None
        or len(property_children[0]) != 0
    ):
        raise ResolvedCitationOoxmlError("citation occurrence SDT tag is not canonical")
    children = list(content)
    expected_tags = [
        qn("w:bookmarkStart"),
        qn("w:r"),
        qn("w:r"),
        qn("w:r"),
        qn("w:r"),
        qn("w:r"),
        qn("w:bookmarkEnd"),
    ]
    if [item.tag for item in children] != expected_tags or any(item.tail is not None for item in children):
        raise ResolvedCitationOoxmlError("citation occurrence field topology is not canonical")
    bookmark_start, begin_run, instruction_run, separate_run, cached_run, end_run, bookmark_end = children
    if (
        tuple(bookmark_start.attrib) != (qn("w:id"), qn("w:name"))
        or bookmark_start.get(qn("w:name")) != identity.bookmark_name
        or bookmark_start.text is not None
        or len(bookmark_start) != 0
        or tuple(bookmark_end.attrib) != (qn("w:id"),)
        or bookmark_end.get(qn("w:id")) != bookmark_start.get(qn("w:id"))
        or bookmark_end.text is not None
        or len(bookmark_end) != 0
    ):
        raise ResolvedCitationOoxmlError("citation occurrence bookmark range is not canonical")
    proof = prove_bookmark_name(bookmark_inventory, identity.bookmark_name, scope_element=sdt)
    if not proof.valid or proof.start is None or proof.end is None:
        raise ResolvedCitationOoxmlError("citation occurrence bookmark is not globally proven")
    if proof.start.element is not bookmark_start or proof.end.element is not bookmark_end:
        raise ResolvedCitationOoxmlError("citation occurrence bookmark does not bind its SDT")
    _prove_field_run(begin_run, "begin", locked=True)
    _prove_instruction_run(instruction_run, citation_instruction(identity.item_refs))
    _prove_field_run(separate_run, "separate")
    _prove_cached_run(cached_run, identity.cached_result)
    _prove_field_run(end_run, "end")


def prove_citation_projection(document: Any, projection: ResolvedCitationProjection) -> None:
    """Prove one-to-one map/SDT/bookmark/field ownership after reopen."""

    from docx.oxml.ns import qn

    validate_citation_authorities(projection.item_map, projection.occurrence_map)
    inventory = build_docx_bookmark_inventory(document)
    expected_by_tag = {item.tag: item for item in projection.occurrence_map.occurrences}
    found: dict[str, list[Any]] = {tag: [] for tag in expected_by_tag}
    owned_sdt_elements: set[int] = set()
    physical_tags: list[str] = []
    roots = _word_xml_roots(document)
    for root in roots:
        for tag_element in root.iter(qn("w:tag")):
            value = tag_element.get(qn("w:val")) or ""
            if not value.startswith("docwen-citation-occurrence-v1:"):
                continue
            sdt_properties = tag_element.getparent()
            sdt = None if sdt_properties is None else sdt_properties.getparent()
            if root is not document.element or value not in expected_by_tag or sdt is None:
                raise ResolvedCitationOoxmlError("unmapped citation occurrence SDT is present")
            found[value].append(sdt)
            owned_sdt_elements.add(id(sdt))
            physical_tags.append(value)
    for tag, identity in expected_by_tag.items():
        if len(found[tag]) != 1:
            raise ResolvedCitationOoxmlError("citation occurrence SDT is missing or duplicated")
        prove_citation_occurrence_sdt(found[tag][0], identity, bookmark_inventory=inventory)
    if physical_tags != [item.tag for item in projection.occurrence_map.occurrences]:
        raise ResolvedCitationOoxmlError("citation occurrence physical order differs from source authority")
    planned_word_tags = {item.word_tag for item in projection.item_map.items}
    planned_word_tag_keys = {item.casefold() for item in planned_word_tags}
    for root in roots:
        for instruction_element in root.iter(qn("w:instrText")):
            instruction = instruction_element.text or ""
            word_tag_keys = {item.casefold() for item in _WORD_TAG_TOKEN_RE.findall(instruction)}
            if not word_tag_keys.intersection(planned_word_tag_keys):
                continue
            owner = _ancestor(instruction_element, qn("w:sdt"))
            if root is not document.element or owner is None or id(owner) not in owned_sdt_elements:
                raise ResolvedCitationOoxmlError("citation Word tag is reused by a non-owned field")


def recover_resolved_citations(
    item_map: CitationItemMap,
    occurrence_map: CitationOccurrenceMap,
) -> tuple[ResolvedCitation, ...]:
    """Restore authenticated neutral Citation records after physical proof."""

    validate_citation_authorities(item_map, occurrence_map)
    items_by_tag = {item.word_tag: item for item in item_map.items}
    citations: list[ResolvedCitation] = []
    for occurrence in occurrence_map.occurrences:
        resolved_items = tuple(
            ResolvedCitationItem(
                citation_key=item_ref.citation_key,
                record_id=items_by_tag[item_ref.word_tag].record_id,
                record_sha256=items_by_tag[item_ref.word_tag].record_sha256,
                presentation=items_by_tag[item_ref.word_tag].presentation,
            )
            for item_ref in occurrence.item_refs
        )
        citations.append(
            ResolvedCitation(
                source_start=occurrence.source_start,
                source_end=occurrence.source_end,
                source_slice_sha256=occurrence.source_slice_sha256,
                authored_token=occurrence.authored_token,
                form=occurrence.form,  # type: ignore[arg-type]
                cluster_id=occurrence.cluster_id,
                items=resolved_items,
                cached_result=occurrence.cached_result,
            )
        )
    return tuple(citations)


def read_proven_resolved_citations(path: str | Path) -> tuple[ResolvedCitation, ...]:
    """Read Citations only when both maps, OPC trios, and physical fields prove them."""

    from docx import Document

    from docwen_core._docx_semantics_v3_package import read_owned_map_parts, verify_custom_xml_support

    package_path = Path(path)
    try:
        with ZipFile(package_path) as package:
            owned = read_owned_map_parts(package)
            has_items = CITATION_ITEM_MAP_NAMESPACE in owned
            has_occurrences = CITATION_OCCURRENCE_MAP_NAMESPACE in owned
            if has_items != has_occurrences:
                raise ResolvedCitationOoxmlError("resolved Citation requires exactly one of both canonical maps")
            if not has_items:
                if _package_has_owned_citation_signal(package):
                    raise ResolvedCitationOoxmlError("citation-owned physical signal has no authority maps")
                return ()
            item_number, item_root = owned[CITATION_ITEM_MAP_NAMESPACE]
            occurrence_number, occurrence_root = owned[CITATION_OCCURRENCE_MAP_NAMESPACE]
            verify_custom_xml_support(package, item_number, CITATION_ITEM_MAP_NAMESPACE)
            verify_custom_xml_support(package, occurrence_number, CITATION_OCCURRENCE_MAP_NAMESPACE)
            item_map = parse_citation_item_map(item_root)
            occurrence_map = parse_citation_occurrence_map(occurrence_root)
            validate_citation_authorities(item_map, occurrence_map)
        document = Document(str(package_path))
        prove_citation_projection(document, ResolvedCitationProjection(item_map, occurrence_map))
        return recover_resolved_citations(item_map, occurrence_map)
    except ResolvedCitationOoxmlError:
        raise
    except DocxSemanticsV3Error as exc:
        raise ResolvedCitationOoxmlError(f"resolved-citation package proof failed: {exc}") from exc


def _append_field_run(container: Any, field_type: str, *, locked: bool = False) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    run = OxmlElement("w:r")
    field = OxmlElement("w:fldChar")
    field.set(qn("w:fldCharType"), field_type)
    if locked:
        field.set(qn("w:fldLock"), "true")
    run.append(field)
    container.append(run)


def _prove_field_run(run: Any, field_type: str, *, locked: bool = False) -> None:
    from docx.oxml.ns import qn

    children = list(run)
    expected_attributes = (qn("w:fldCharType"), qn("w:fldLock")) if locked else (qn("w:fldCharType"),)
    if (
        run.attrib
        or len(children) != 1
        or children[0].tag != qn("w:fldChar")
        or tuple(children[0].attrib) != expected_attributes
        or children[0].get(qn("w:fldCharType")) != field_type
        or children[0].get(qn("w:fldLock")) != ("true" if locked else None)
        or children[0].get(qn("w:dirty")) is not None
        or children[0].text is not None
        or children[0].tail is not None
        or len(children[0]) != 0
    ):
        raise ResolvedCitationOoxmlError("citation complex field marker is not canonical")


def _prove_instruction_run(run: Any, expected: str) -> None:
    from docx.oxml.ns import qn

    children = list(run)
    if (
        run.attrib
        or len(children) != 1
        or children[0].tag != qn("w:instrText")
        or tuple(children[0].attrib) != (_XML_SPACE,)
        or children[0].get(_XML_SPACE) != "preserve"
        or children[0].text != expected
        or children[0].tail is not None
        or len(children[0]) != 0
    ):
        raise ResolvedCitationOoxmlError("citation field instruction is not exact")


def _prove_cached_run(run: Any, expected: str) -> None:
    from docx.oxml.ns import qn

    children = list(run)
    preserve = expected[:1].isspace() or expected[-1:].isspace()
    expected_attributes = (_XML_SPACE,) if preserve else ()
    if (
        run.attrib
        or len(children) != 1
        or children[0].tag != qn("w:t")
        or tuple(children[0].attrib) != expected_attributes
        or children[0].get(_XML_SPACE) != ("preserve" if preserve else None)
        or children[0].text != expected
        or children[0].tail is not None
        or len(children[0]) != 0
    ):
        raise ResolvedCitationOoxmlError("citation cached field result is not exact")


def _package_has_owned_citation_signal(package: ZipFile) -> bool:
    if "word/document.xml" not in package.namelist():
        return False
    from docx.oxml.ns import qn

    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    try:
        root = etree.fromstring(package.read("word/document.xml"), parser)
    except etree.XMLSyntaxError as exc:
        raise ResolvedCitationOoxmlError("main document XML is malformed") from exc
    if any(
        (item.get(qn("w:val")) or "").startswith("docwen-citation-occurrence-v1:") for item in root.iter(qn("w:tag"))
    ):
        return True
    return any(
        _is_citation_instruction(item.text or "") and _WORD_TAG_TOKEN_RE.search(item.text or "") is not None
        for item in root.iter(qn("w:instrText"))
    )


def _document_instructions(document: Any) -> tuple[str, ...]:
    from docx.oxml.ns import qn

    return tuple(item.text or "" for root in _word_xml_roots(document) for item in root.iter(qn("w:instrText")))


def _word_xml_roots(document: Any) -> tuple[Any, ...]:
    roots: list[Any] = [document.element]
    seen = {str(document.part.partname)}
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    for part in document.part.package.parts:
        part_name = str(part.partname)
        content_type = str(part.content_type).lower()
        if (
            part_name in seen
            or not part_name.lower().startswith("/word/")
            or not part_name.lower().endswith(".xml")
            or not content_type.endswith("xml")
        ):
            continue
        seen.add(part_name)
        element = getattr(part, "element", None)
        if element is None:
            element = getattr(part, "_element", None)
        if element is None:
            try:
                blob = part.blob
                if not isinstance(blob, bytes) or not blob:
                    raise ValueError
                element = etree.fromstring(blob, parser)
            except (AttributeError, TypeError, ValueError, etree.XMLSyntaxError) as exc:
                raise ResolvedCitationOoxmlError(
                    f"cannot inspect WordprocessingML part {part_name} for citation collisions"
                ) from exc
        roots.append(element)
    return tuple(roots)


def _is_citation_instruction(instruction: str) -> bool:
    tokens = instruction.split()
    return bool(tokens) and tokens[0].casefold() == "citation"


def _ancestor(element: Any, tag: str) -> Any | None:
    current = element.getparent()
    while current is not None:
        if current.tag == tag:
            return current
        current = current.getparent()
    return None


__all__ = [
    "CITATION_ITEM_MAP_NAMESPACE",
    "CITATION_OCCURRENCE_MAP_NAMESPACE",
    "CitationItemIdentity",
    "CitationItemMap",
    "CitationItemReferenceIdentity",
    "CitationOccurrenceIdentity",
    "CitationOccurrenceMap",
    "ResolvedCitationOoxmlError",
    "ResolvedCitationProjection",
    "build_resolved_citation_projection",
    "citation_instruction",
    "citation_item_map_xml",
    "citation_occurrence_map_xml",
    "citation_occurrence_sdt",
    "derive_citation_item",
    "derive_citation_item_reference",
    "derive_citation_occurrence",
    "parse_citation_item_map",
    "parse_citation_occurrence_map",
    "preflight_citation_document",
    "prove_citation_occurrence_sdt",
    "prove_citation_projection",
    "read_proven_resolved_citations",
    "recover_resolved_citations",
    "validate_citation_authorities",
    "validate_citation_item_map",
    "validate_citation_occurrence_map",
]
