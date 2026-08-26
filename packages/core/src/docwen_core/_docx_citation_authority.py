"""Closed v4 resolved-citation carriers and WordprocessingML projection."""

from __future__ import annotations

import base64
import binascii
import hashlib
import re
from dataclasses import dataclass
from itertools import pairwise
from typing import Any

from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error
from docwen_core.models.resolved_numbering import ResolvedCitation

CITATION_ITEM_MAP_NAMESPACE = "https://docwen.dev/schema/document-citation-item-map/v1"
CITATION_OCCURRENCE_MAP_NAMESPACE = "https://docwen.dev/schema/document-citation-occurrence-map/v1"

_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_WORD_TAG_RE = re.compile(r"^DWCIT_[0-9a-f]{32}$")
_OCCURRENCE_TAG_RE = re.compile(r"^docwen-citation-occurrence-v1:[0-9a-f]{32}$")
_BOOKMARK_RE = re.compile(r"^_DWC_[0-9a-f]{35}$")
_CITATION_KEY_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,127}$")
_RECORD_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$")
_CLUSTER_ID_RE = re.compile(r"^[A-Za-z][A-Za-z0-9-]{0,63}$")
_WORD_TAG_TOKEN_RE = re.compile(r"\bDWCIT_[0-9a-f]{32}\b", re.IGNORECASE)


class ResolvedCitationOoxmlError(DocxSemanticsV3Error):
    """Resolved citation authority or its physical projection is invalid."""


@dataclass(frozen=True, slots=True)
class CitationItemIdentity:
    source_sha256: str
    word_tag: str
    record_id: str
    record_sha256: str
    presentation: str
    presentation_sha256: str
    sha256: str

    @property
    def semantic_key(self) -> tuple[str, str, str]:
        return self.record_id, self.record_sha256, self.presentation


@dataclass(frozen=True, slots=True)
class CitationItemReferenceIdentity:
    citation_key: str
    word_tag: str
    item_sha256: str
    sha256: str


@dataclass(frozen=True, slots=True)
class CitationOccurrenceIdentity:
    tag: str
    bookmark_name: str
    source_sha256: str
    source_start: int
    source_end: int
    source_slice_sha256: str
    authored_token: str
    form: str
    cluster_id: str
    cached_result: str
    cached_result_sha256: str
    item_refs: tuple[CitationItemReferenceIdentity, ...]
    sha256: str

    @property
    def occurrence_key(self) -> tuple[int, int]:
        return self.source_start, self.source_end


@dataclass(frozen=True, slots=True)
class CitationItemMap:
    source_sha256: str
    items: tuple[CitationItemIdentity, ...]


@dataclass(frozen=True, slots=True)
class CitationOccurrenceMap:
    source_sha256: str
    occurrences: tuple[CitationOccurrenceIdentity, ...]


@dataclass(frozen=True, slots=True)
class ResolvedCitationProjection:
    item_map: CitationItemMap
    occurrence_map: CitationOccurrenceMap


def derive_citation_item(
    *,
    source_sha256: str,
    record_id: str,
    record_sha256: str,
    presentation: str,
) -> CitationItemIdentity:
    """Derive the opaque Word tag without narrowing the provider identity."""

    _require_sha256(source_sha256, "citation source")
    _require_record_id(record_id)
    _require_sha256(record_sha256, "citation record")
    _require_xml_text(presentation, "citation presentation", minimum=1, maximum=4096)
    presentation_sha256 = _sha256_text(presentation)
    digest = _sha256_text(
        "\0".join(
            (
                "docwen-citation-item-map-v1",
                source_sha256,
                record_id,
                record_sha256,
                presentation_sha256,
            )
        )
    )
    return CitationItemIdentity(
        source_sha256=source_sha256,
        word_tag=f"DWCIT_{digest[:32]}",
        record_id=record_id,
        record_sha256=record_sha256,
        presentation=presentation,
        presentation_sha256=presentation_sha256,
        sha256=digest,
    )


def derive_citation_item_reference(
    *,
    citation_key: str,
    word_tag: str,
    item_sha256: str,
) -> CitationItemReferenceIdentity:
    _require_citation_key(citation_key)
    if _WORD_TAG_RE.fullmatch(word_tag) is None:
        raise ResolvedCitationOoxmlError("citation item reference has a non-canonical Word tag")
    _require_sha256(item_sha256, "citation item")
    digest = _sha256_text("\0".join(("docwen-citation-item-ref-v1", citation_key, word_tag, item_sha256)))
    return CitationItemReferenceIdentity(citation_key, word_tag, item_sha256, digest)


def derive_citation_occurrence(
    *,
    source_sha256: str,
    source_start: int,
    source_end: int,
    source_slice_sha256: str,
    authored_token: str,
    form: str,
    cluster_id: str,
    cached_result: str,
    item_refs: tuple[CitationItemReferenceIdentity, ...],
) -> CitationOccurrenceIdentity:
    _require_sha256(source_sha256, "citation source")
    _require_source_range(source_start, source_end)
    _require_sha256(source_slice_sha256, "citation source slice")
    _require_xml_text(authored_token, "citation authored token", minimum=2, maximum=4096)
    if _sha256_text(authored_token) != source_slice_sha256:
        raise ResolvedCitationOoxmlError("citation authored token does not match its source-slice digest")
    if form not in {"narrative", "parenthetical"}:
        raise ResolvedCitationOoxmlError("citation occurrence form is outside the closed set")
    if _CLUSTER_ID_RE.fullmatch(cluster_id) is None:
        raise ResolvedCitationOoxmlError("citation cluster identity is outside the closed set")
    _require_xml_text(cached_result, "citation cached result", minimum=1, maximum=4096)
    if not 1 <= len(item_refs) <= 64:
        raise ResolvedCitationOoxmlError("citation occurrence must contain 1..64 item references")
    _validate_item_refs(item_refs)
    cached_result_sha256 = _sha256_text(cached_result)
    digest = _sha256_text(
        "\0".join(
            (
                "docwen-citation-occurrence-map-v1",
                source_sha256,
                str(source_start),
                str(source_end),
                source_slice_sha256,
                form,
                cluster_id,
                cached_result_sha256,
                ",".join(item.sha256 for item in item_refs),
            )
        )
    )
    return CitationOccurrenceIdentity(
        tag=f"docwen-citation-occurrence-v1:{digest[:32]}",
        bookmark_name=f"_DWC_{digest[:35]}",
        source_sha256=source_sha256,
        source_start=source_start,
        source_end=source_end,
        source_slice_sha256=source_slice_sha256,
        authored_token=authored_token,
        form=form,
        cluster_id=cluster_id,
        cached_result=cached_result,
        cached_result_sha256=cached_result_sha256,
        item_refs=item_refs,
        sha256=digest,
    )


def build_resolved_citation_projection(
    source_sha256: str,
    citations: tuple[ResolvedCitation, ...],
) -> ResolvedCitationProjection | None:
    """Build and collision-check both closed authorities from the typed snapshot."""

    _require_sha256(source_sha256, "citation source")
    if not citations:
        return None
    items_by_semantic_key: dict[tuple[str, str, str], CitationItemIdentity] = {}
    occurrences: list[CitationOccurrenceIdentity] = []
    for citation in citations:
        item_refs: list[CitationItemReferenceIdentity] = []
        for resolved_item in citation.items:
            semantic_key = (
                resolved_item.record_id,
                resolved_item.record_sha256,
                resolved_item.presentation,
            )
            item = items_by_semantic_key.get(semantic_key)
            if item is None:
                item = derive_citation_item(
                    source_sha256=source_sha256,
                    record_id=resolved_item.record_id,
                    record_sha256=resolved_item.record_sha256,
                    presentation=resolved_item.presentation,
                )
                items_by_semantic_key[semantic_key] = item
            item_refs.append(
                derive_citation_item_reference(
                    citation_key=resolved_item.citation_key,
                    word_tag=item.word_tag,
                    item_sha256=item.sha256,
                )
            )
        occurrences.append(
            derive_citation_occurrence(
                source_sha256=source_sha256,
                source_start=citation.source_start,
                source_end=citation.source_end,
                source_slice_sha256=citation.source_slice_sha256,
                authored_token=citation.authored_token,
                form=citation.form,
                cluster_id=citation.cluster_id,
                cached_result=citation.cached_result,
                item_refs=tuple(item_refs),
            )
        )
    projection = ResolvedCitationProjection(
        CitationItemMap(source_sha256, tuple(sorted(items_by_semantic_key.values(), key=lambda item: item.word_tag))),
        CitationOccurrenceMap(
            source_sha256,
            tuple(sorted(occurrences, key=lambda item: (item.source_start, item.source_end, item.tag))),
        ),
    )
    validate_citation_authorities(projection.item_map, projection.occurrence_map)
    return projection


def citation_item_map_xml(item_map: CitationItemMap) -> bytes:
    validated = validate_citation_item_map(item_map)
    entries = "".join(
        (
            f'<item word_tag="{item.word_tag}" record_id="{_xml_attr(item.record_id)}" '
            f'record_sha256="{item.record_sha256}" '
            f'presentation_base64="{_encode_text(item.presentation)}" '
            f'presentation_sha256="{item.presentation_sha256}" sha256="{item.sha256}"/>'
        )
        for item in validated.items
    )
    root = (
        f'<documentCitationItemMap xmlns="{CITATION_ITEM_MAP_NAMESPACE}" '
        f'version="1" source_sha256="{validated.source_sha256}">{entries}</documentCitationItemMap>'
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def citation_occurrence_map_xml(occurrence_map: CitationOccurrenceMap) -> bytes:
    validated = validate_citation_occurrence_map(occurrence_map)
    entries = "".join(_occurrence_xml(item) for item in validated.occurrences)
    root = (
        f'<documentCitationOccurrenceMap xmlns="{CITATION_OCCURRENCE_MAP_NAMESPACE}" '
        f'version="1" source_sha256="{validated.source_sha256}">{entries}</documentCitationOccurrenceMap>'
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def parse_citation_item_map(root: Any) -> CitationItemMap:
    namespace = f"{{{CITATION_ITEM_MAP_NAMESPACE}}}"
    if (
        root.tag != f"{namespace}documentCitationItemMap"
        or tuple(root.attrib) != ("version", "source_sha256")
        or root.get("version") != "1"
        or root.text is not None
        or root.tail is not None
    ):
        raise ResolvedCitationOoxmlError("citation-item map root is not closed and canonical")
    source_sha256 = root.get("source_sha256", "")
    _require_sha256(source_sha256, "citation source")
    items: list[CitationItemIdentity] = []
    expected_attributes = (
        "word_tag",
        "record_id",
        "record_sha256",
        "presentation_base64",
        "presentation_sha256",
        "sha256",
    )
    for element in root:
        if (
            element.tag != f"{namespace}item"
            or tuple(element.attrib) != expected_attributes
            or element.text is not None
            or element.tail is not None
            or len(element) != 0
        ):
            raise ResolvedCitationOoxmlError("citation-item record is not closed and canonical")
        presentation = _decode_text(
            element.get("presentation_base64", ""),
            context="citation presentation",
            minimum=1,
            maximum=4096,
        )
        expected = derive_citation_item(
            source_sha256=source_sha256,
            record_id=element.get("record_id", ""),
            record_sha256=element.get("record_sha256", ""),
            presentation=presentation,
        )
        if (
            element.get("word_tag") != expected.word_tag
            or element.get("presentation_sha256") != expected.presentation_sha256
            or element.get("sha256") != expected.sha256
        ):
            raise ResolvedCitationOoxmlError("citation-item identity does not recompute")
        items.append(expected)
    return validate_citation_item_map(CitationItemMap(source_sha256, tuple(items)))


def parse_citation_occurrence_map(root: Any) -> CitationOccurrenceMap:
    namespace = f"{{{CITATION_OCCURRENCE_MAP_NAMESPACE}}}"
    if (
        root.tag != f"{namespace}documentCitationOccurrenceMap"
        or tuple(root.attrib) != ("version", "source_sha256")
        or root.get("version") != "1"
        or root.text is not None
        or root.tail is not None
    ):
        raise ResolvedCitationOoxmlError("citation-occurrence map root is not closed and canonical")
    source_sha256 = root.get("source_sha256", "")
    _require_sha256(source_sha256, "citation source")
    occurrences: list[CitationOccurrenceIdentity] = []
    expected_attributes = (
        "tag",
        "bookmark_name",
        "source_sha256",
        "source_start",
        "source_end",
        "source_slice_sha256",
        "authored_token_base64",
        "form",
        "cluster_id",
        "cached_result_base64",
        "cached_result_sha256",
        "sha256",
    )
    for element in root:
        if (
            element.tag != f"{namespace}citationOccurrence"
            or tuple(element.attrib) != expected_attributes
            or element.text is not None
            or element.tail is not None
            or not 1 <= len(element) <= 64
            or element.get("source_sha256") != source_sha256
        ):
            raise ResolvedCitationOoxmlError("citation-occurrence record is not closed and canonical")
        item_refs = tuple(_parse_item_ref(item, namespace) for item in element)
        authored_token = _decode_text(
            element.get("authored_token_base64", ""),
            context="citation authored token",
            minimum=2,
            maximum=4096,
        )
        cached_result = _decode_text(
            element.get("cached_result_base64", ""),
            context="citation cached result",
            minimum=1,
            maximum=4096,
        )
        source_start = _parse_decimal(element.get("source_start", ""), "citation source_start")
        source_end = _parse_decimal(element.get("source_end", ""), "citation source_end")
        expected = derive_citation_occurrence(
            source_sha256=source_sha256,
            source_start=source_start,
            source_end=source_end,
            source_slice_sha256=element.get("source_slice_sha256", ""),
            authored_token=authored_token,
            form=element.get("form", ""),
            cluster_id=element.get("cluster_id", ""),
            cached_result=cached_result,
            item_refs=item_refs,
        )
        if (
            element.get("tag") != expected.tag
            or element.get("bookmark_name") != expected.bookmark_name
            or element.get("cached_result_sha256") != expected.cached_result_sha256
            or element.get("sha256") != expected.sha256
        ):
            raise ResolvedCitationOoxmlError("citation-occurrence identity does not recompute")
        occurrences.append(expected)
    return validate_citation_occurrence_map(CitationOccurrenceMap(source_sha256, tuple(occurrences)))


def validate_citation_item_map(item_map: CitationItemMap) -> CitationItemMap:
    _require_sha256(item_map.source_sha256, "citation source")
    if not item_map.items:
        raise ResolvedCitationOoxmlError("citation-item map must not be empty")
    if item_map.items != tuple(sorted(item_map.items, key=lambda item: item.word_tag)):
        raise ResolvedCitationOoxmlError("citation-item records are not canonically ordered")
    semantic_keys: set[tuple[str, str, str]] = set()
    full_digests: dict[str, tuple[str, str, str]] = {}
    truncated_tags: dict[str, str] = {}
    for item in item_map.items:
        if item.source_sha256 != item_map.source_sha256:
            raise ResolvedCitationOoxmlError("citation-item records mix source identities")
        semantic_key = item.semantic_key
        if semantic_key in semantic_keys:
            raise ResolvedCitationOoxmlError("citation-item semantic tuple is duplicated")
        semantic_keys.add(semantic_key)
        previous_preimage = full_digests.setdefault(item.sha256, semantic_key)
        if previous_preimage != semantic_key:
            raise ResolvedCitationOoxmlError("citation-item full digest collides")
        previous_digest = truncated_tags.setdefault(item.word_tag, item.sha256)
        if previous_digest != item.sha256:
            raise ResolvedCitationOoxmlError("citation-item truncated Word tag collides")
        expected = derive_citation_item(
            source_sha256=item.source_sha256,
            record_id=item.record_id,
            record_sha256=item.record_sha256,
            presentation=item.presentation,
        )
        if item != expected:
            raise ResolvedCitationOoxmlError("citation-item record is not canonically derived")
    return item_map


def validate_citation_occurrence_map(occurrence_map: CitationOccurrenceMap) -> CitationOccurrenceMap:
    _require_sha256(occurrence_map.source_sha256, "citation source")
    if not occurrence_map.occurrences:
        raise ResolvedCitationOoxmlError("citation-occurrence map must not be empty")
    canonical = tuple(
        sorted(
            occurrence_map.occurrences,
            key=lambda item: (item.source_start, item.source_end, item.tag),
        )
    )
    if occurrence_map.occurrences != canonical:
        raise ResolvedCitationOoxmlError("citation-occurrence records are not canonically ordered")
    full_digests: dict[str, tuple[int, int, str]] = {}
    tags: dict[str, str] = {}
    bookmarks: dict[str, str] = {}
    cluster_ids: set[str] = set()
    for item in occurrence_map.occurrences:
        if item.source_sha256 != occurrence_map.source_sha256:
            raise ResolvedCitationOoxmlError("citation occurrences mix source identities")
        key = (item.source_start, item.source_end, item.cluster_id)
        previous_preimage = full_digests.setdefault(item.sha256, key)
        if previous_preimage != key:
            raise ResolvedCitationOoxmlError("citation-occurrence full digest collides")
        previous_tag_digest = tags.setdefault(item.tag, item.sha256)
        if previous_tag_digest != item.sha256:
            raise ResolvedCitationOoxmlError("citation-occurrence truncated SDT tag collides")
        previous_bookmark_digest = bookmarks.setdefault(item.bookmark_name.casefold(), item.sha256)
        if previous_bookmark_digest != item.sha256:
            raise ResolvedCitationOoxmlError("citation-occurrence truncated bookmark collides")
        if item.cluster_id in cluster_ids:
            raise ResolvedCitationOoxmlError("citation cluster identity is duplicated")
        cluster_ids.add(item.cluster_id)
        expected = derive_citation_occurrence(
            source_sha256=item.source_sha256,
            source_start=item.source_start,
            source_end=item.source_end,
            source_slice_sha256=item.source_slice_sha256,
            authored_token=item.authored_token,
            form=item.form,
            cluster_id=item.cluster_id,
            cached_result=item.cached_result,
            item_refs=item.item_refs,
        )
        if item != expected:
            raise ResolvedCitationOoxmlError("citation-occurrence record is not canonically derived")
    for previous, current in pairwise(occurrence_map.occurrences):
        if current.source_start < previous.source_end:
            raise ResolvedCitationOoxmlError("citation occurrence source ranges overlap")
    return occurrence_map


def validate_citation_authorities(
    item_map: CitationItemMap,
    occurrence_map: CitationOccurrenceMap,
) -> None:
    validate_citation_item_map(item_map)
    validate_citation_occurrence_map(occurrence_map)
    if item_map.source_sha256 != occurrence_map.source_sha256:
        raise ResolvedCitationOoxmlError("citation maps disagree on source identity")
    items_by_tag = {item.word_tag: item for item in item_map.items}
    used_tags: set[str] = set()
    for occurrence in occurrence_map.occurrences:
        for item_ref in occurrence.item_refs:
            item = items_by_tag.get(item_ref.word_tag)
            if item is None or item.sha256 != item_ref.item_sha256:
                raise ResolvedCitationOoxmlError("citation item reference is dangling or cross-linked")
            used_tags.add(item_ref.word_tag)
    if used_tags != set(items_by_tag):
        raise ResolvedCitationOoxmlError("citation-item map contains an unused record")


def _occurrence_xml(item: CitationOccurrenceIdentity) -> str:
    item_refs = "".join(
        (
            f'<itemRef citation_key="{item_ref.citation_key}" word_tag="{item_ref.word_tag}" '
            f'item_sha256="{item_ref.item_sha256}" sha256="{item_ref.sha256}"/>'
        )
        for item_ref in item.item_refs
    )
    return (
        f'<citationOccurrence tag="{item.tag}" bookmark_name="{item.bookmark_name}" '
        f'source_sha256="{item.source_sha256}" source_start="{item.source_start}" '
        f'source_end="{item.source_end}" source_slice_sha256="{item.source_slice_sha256}" '
        f'authored_token_base64="{_encode_text(item.authored_token)}" form="{item.form}" '
        f'cluster_id="{item.cluster_id}" cached_result_base64="{_encode_text(item.cached_result)}" '
        f'cached_result_sha256="{item.cached_result_sha256}" sha256="{item.sha256}">{item_refs}'
        "</citationOccurrence>"
    )


def _parse_item_ref(element: Any, namespace: str) -> CitationItemReferenceIdentity:
    if (
        element.tag != f"{namespace}itemRef"
        or tuple(element.attrib) != ("citation_key", "word_tag", "item_sha256", "sha256")
        or element.text is not None
        or element.tail is not None
        or len(element) != 0
    ):
        raise ResolvedCitationOoxmlError("citation item reference is not closed and canonical")
    expected = derive_citation_item_reference(
        citation_key=element.get("citation_key", ""),
        word_tag=element.get("word_tag", ""),
        item_sha256=element.get("item_sha256", ""),
    )
    if element.get("sha256") != expected.sha256:
        raise ResolvedCitationOoxmlError("citation item-reference digest does not recompute")
    return expected


def _validate_item_refs(item_refs: tuple[CitationItemReferenceIdentity, ...]) -> None:
    if len({item.citation_key for item in item_refs}) != len(item_refs):
        raise ResolvedCitationOoxmlError("citation occurrence repeats an authored key")
    for item in item_refs:
        expected = derive_citation_item_reference(
            citation_key=item.citation_key,
            word_tag=item.word_tag,
            item_sha256=item.item_sha256,
        )
        if item != expected:
            raise ResolvedCitationOoxmlError("citation item reference is not canonically derived")


def _validate_single_occurrence(identity: CitationOccurrenceIdentity) -> None:
    if _OCCURRENCE_TAG_RE.fullmatch(identity.tag) is None or _BOOKMARK_RE.fullmatch(identity.bookmark_name) is None:
        raise ResolvedCitationOoxmlError("citation occurrence has a non-portable physical identity")
    expected = derive_citation_occurrence(
        source_sha256=identity.source_sha256,
        source_start=identity.source_start,
        source_end=identity.source_end,
        source_slice_sha256=identity.source_slice_sha256,
        authored_token=identity.authored_token,
        form=identity.form,
        cluster_id=identity.cluster_id,
        cached_result=identity.cached_result,
        item_refs=identity.item_refs,
    )
    if identity != expected:
        raise ResolvedCitationOoxmlError("citation occurrence is not canonically derived")


def _encode_text(value: str) -> str:
    return base64.b64encode(value.encode("utf-8")).decode("ascii")


def _decode_text(value: str, *, context: str, minimum: int, maximum: int) -> str:
    try:
        raw = base64.b64decode(value, validate=True)
        if base64.b64encode(raw).decode("ascii") != value:
            raise ValueError
        decoded = raw.decode("utf-8")
    except (binascii.Error, UnicodeDecodeError, ValueError) as exc:
        raise ResolvedCitationOoxmlError(f"{context} is not canonical UTF-8 RFC 4648 base64") from exc
    _require_xml_text(decoded, context, minimum=minimum, maximum=maximum)
    return decoded


def _sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _require_sha256(value: str, context: str) -> None:
    if _SHA256_RE.fullmatch(value) is None:
        raise ResolvedCitationOoxmlError(f"{context} SHA-256 is not lowercase hexadecimal")


def _require_record_id(value: str) -> None:
    if _RECORD_ID_RE.fullmatch(value) is None:
        raise ResolvedCitationOoxmlError("citation record identity is outside the closed set")


def _require_citation_key(value: str) -> None:
    if _CITATION_KEY_RE.fullmatch(value) is None:
        raise ResolvedCitationOoxmlError("citation key is outside the closed set")


def _require_source_range(source_start: int, source_end: int) -> None:
    if isinstance(source_start, bool) or isinstance(source_end, bool) or not 0 <= source_start < source_end:
        raise ResolvedCitationOoxmlError("citation source range is invalid")


def _parse_decimal(value: str, context: str) -> int:
    if re.fullmatch(r"0|[1-9][0-9]*", value) is None:
        raise ResolvedCitationOoxmlError(f"{context} is not a canonical decimal")
    return int(value)


def _require_xml_text(value: str, context: str, *, minimum: int, maximum: int) -> None:
    if not minimum <= len(value) <= maximum or not all(
        character in {"\t", "\n", "\r"}
        or "\u0020" <= character <= "\ud7ff"
        or "\ue000" <= character <= "\ufffd"
        or "\U00010000" <= character <= "\U0010ffff"
        for character in value
    ):
        raise ResolvedCitationOoxmlError(f"{context} is not bounded XML 1.0 text")


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
    "citation_item_map_xml",
    "citation_occurrence_map_xml",
    "derive_citation_item",
    "derive_citation_item_reference",
    "derive_citation_occurrence",
    "parse_citation_item_map",
    "parse_citation_occurrence_map",
    "validate_citation_authorities",
    "validate_citation_item_map",
    "validate_citation_occurrence_map",
]
