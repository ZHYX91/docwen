"""Closed embedded resources and resolved citations for the v4 neutral document."""

from __future__ import annotations

import base64
import binascii
import hashlib
import re
import warnings
from io import BytesIO
from typing import Literal, cast

from PIL import Image, ImageSequence, UnidentifiedImageError

from docwen_core.models._resolved_numbering_validation import (
    _array,
    _definition_id,
    _exact_keys,
    _fail,
    _integer,
    _is_xml_10_text,
    _object,
    _sha256,
    _string,
)
from docwen_core.models.resolved_numbering import (
    MAX_RESOLVED_DOCUMENT_EMBEDDED_BYTES,
    ResolvedCitation,
    ResolvedCitationItem,
    ResolvedEmbeddedResource,
    ResolvedNumberingPortError,
    ResolvedResourceOccurrence,
)
from docwen_core.semantic_bibliography import (
    SemanticBibliographyResourceError,
    parse_semantic_bibliography,
)

_CITATION_KEY_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,127}$")
_RECORD_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$")
_LINKED_MEDIA_TYPES = frozenset({"image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp"})
_BIBLIOGRAPHY_MEDIA_TYPE = "application/vnd.docwen.semantic-bibliography+json"
_PIL_FORMAT_FOR_MEDIA_TYPE = {
    "image/png": "PNG",
    "image/jpeg": "JPEG",
    "image/gif": "GIF",
    "image/bmp": "BMP",
    "image/webp": "WEBP",
}


def parse_document_extras(
    document_raw: dict[str, object], code: str
) -> tuple[
    tuple[ResolvedResourceOccurrence, ...],
    tuple[ResolvedCitation, ...],
    tuple[ResolvedEmbeddedResource, ...],
]:
    occurrences = tuple(
        _parse_resource_occurrence(item, index, code)
        for index, item in enumerate(
            _array(document_raw["resource_occurrences"], "document.resource_occurrences", code)
        )
    )
    citations = tuple(
        _parse_citation(item, index, code)
        for index, item in enumerate(_array(document_raw["citations"], "document.citations", code))
    )
    resources = tuple(
        _parse_resource(item, index, code)
        for index, item in enumerate(_array(document_raw["resources"], "document.resources", code))
    )
    if sum(item.size_bytes for item in resources) > MAX_RESOLVED_DOCUMENT_EMBEDDED_BYTES:
        _fail(code, "embedded resource bytes exceed the 6,000,000-byte v1 admission limit")
    return occurrences, citations, resources


def _parse_resource_occurrence(raw: object, index: int, code: str) -> ResolvedResourceOccurrence:
    location = f"document.resource_occurrences[{index}]"
    item = _object(raw, location, code)
    _exact_keys(
        item,
        {
            "source_start",
            "source_end",
            "source_slice_sha256",
            "authored_token",
            "authored_locator",
            "resource_id",
        },
        code,
    )
    return ResolvedResourceOccurrence(
        source_start=_integer(item["source_start"], f"{location}.source_start", code, 0),
        source_end=_integer(item["source_end"], f"{location}.source_end", code, 1),
        source_slice_sha256=_sha256(item["source_slice_sha256"], f"{location}.source_slice_sha256", code),
        authored_token=_string(
            item["authored_token"],
            f"{location}.authored_token",
            code,
            minimum=5,
            maximum=8192,
        ),
        authored_locator=_string(
            item["authored_locator"],
            f"{location}.authored_locator",
            code,
            minimum=1,
            maximum=4096,
        ),
        resource_id=_definition_id(item["resource_id"], f"{location}.resource_id", code),
    )


def _parse_resource(raw: object, index: int, code: str) -> ResolvedEmbeddedResource:
    location = f"document.resources[{index}]"
    item = _object(raw, location, code)
    _exact_keys(
        item,
        {"resource_id", "role", "media_type", "size_bytes", "sha256", "content_base64"},
        code,
    )
    role = _string(item["role"], f"{location}.role", code, minimum=1)
    if role not in {"linked_resource", "bibliography"}:
        _fail(code, f"{location}.role is not a closed resource role")
    media_type = _string(item["media_type"], f"{location}.media_type", code, minimum=3, maximum=255)
    if role == "bibliography":
        if media_type != _BIBLIOGRAPHY_MEDIA_TYPE:
            _fail(code, f"{location} bibliography media type is invalid")
    elif media_type not in _LINKED_MEDIA_TYPES:
        _fail(code, f"{location} linked resource media type is unsupported")
    size_bytes = _integer(
        item["size_bytes"],
        f"{location}.size_bytes",
        code,
        1,
        MAX_RESOLVED_DOCUMENT_EMBEDDED_BYTES,
    )
    content_base64 = _string(
        item["content_base64"],
        f"{location}.content_base64",
        code,
        minimum=0,
        maximum=8_000_000,
    )
    try:
        content = base64.b64decode(content_base64, validate=True)
    except (ValueError, UnicodeEncodeError, binascii.Error) as exc:
        _fail(code, f"{location}.content_base64 is not canonical base64")
        raise AssertionError from exc
    if base64.b64encode(content).decode("ascii") != content_base64:
        _fail(code, f"{location}.content_base64 is not canonical base64")
    sha256 = _sha256(item["sha256"], f"{location}.sha256", code)
    if len(content) != size_bytes or hashlib.sha256(content).hexdigest() != sha256:
        _fail(code, f"{location} size/hash does not authenticate embedded bytes")
    if role == "linked_resource":
        _validate_raster(content, media_type, location, code)
    else:
        try:
            parse_semantic_bibliography(content)
        except SemanticBibliographyResourceError as exc:
            _fail(code, f"{location} is not a valid closed semantic bibliography: {exc.code}")
    return ResolvedEmbeddedResource(
        resource_id=_definition_id(item["resource_id"], f"{location}.resource_id", code),
        role=cast(Literal["linked_resource", "bibliography"], role),
        media_type=media_type,
        size_bytes=size_bytes,
        sha256=sha256,
        content=content,
    )


def _parse_citation(raw: object, index: int, code: str) -> ResolvedCitation:
    location = f"document.citations[{index}]"
    item = _object(raw, location, code)
    _exact_keys(
        item,
        {
            "source_start",
            "source_end",
            "source_slice_sha256",
            "authored_token",
            "form",
            "cluster_id",
            "items",
            "cached_result",
        },
        code,
    )
    form = _string(item["form"], f"{location}.form", code, minimum=1)
    if form not in {"narrative", "parenthetical"}:
        _fail(code, f"{location}.form is not a closed citation form")
    items_raw = _array(item["items"], f"{location}.items", code)
    if not 1 <= len(items_raw) <= 64:
        _fail(code, f"{location}.items must contain 1..64 records")
    items = tuple(_parse_citation_item(value, item_index, code, location) for item_index, value in enumerate(items_raw))
    cached_result = _string(item["cached_result"], f"{location}.cached_result", code, minimum=1, maximum=4096)
    if not _is_xml_10_text(cached_result):
        _fail(code, f"{location}.cached_result is not XML 1.0 text")
    return ResolvedCitation(
        source_start=_integer(item["source_start"], f"{location}.source_start", code, 0),
        source_end=_integer(item["source_end"], f"{location}.source_end", code, 1),
        source_slice_sha256=_sha256(item["source_slice_sha256"], f"{location}.source_slice_sha256", code),
        authored_token=_string(
            item["authored_token"],
            f"{location}.authored_token",
            code,
            minimum=2,
            maximum=4096,
        ),
        form=cast(Literal["narrative", "parenthetical"], form),
        cluster_id=_definition_id(item["cluster_id"], f"{location}.cluster_id", code),
        items=items,
        cached_result=cached_result,
    )


def _parse_citation_item(raw: object, index: int, code: str, parent: str) -> ResolvedCitationItem:
    location = f"{parent}.items[{index}]"
    item = _object(raw, location, code)
    _exact_keys(item, {"citation_key", "record_id", "record_sha256", "presentation"}, code)
    citation_key = _string(item["citation_key"], f"{location}.citation_key", code, minimum=1, maximum=128)
    record_id = _string(item["record_id"], f"{location}.record_id", code, minimum=1, maximum=256)
    if _CITATION_KEY_RE.fullmatch(citation_key) is None or _RECORD_ID_RE.fullmatch(record_id) is None:
        _fail(code, f"{location} has an invalid citation key or record identity")
    presentation = _string(item["presentation"], f"{location}.presentation", code, minimum=1, maximum=4096)
    if not _is_xml_10_text(presentation):
        _fail(code, f"{location}.presentation is not XML 1.0 text")
    return ResolvedCitationItem(
        citation_key=citation_key,
        record_id=record_id,
        record_sha256=_sha256(item["record_sha256"], f"{location}.record_sha256", code),
        presentation=presentation,
    )


def _validate_raster(content: bytes, media_type: str, location: str, code: str) -> None:
    expected_format = _PIL_FORMAT_FOR_MEDIA_TYPE[media_type]
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(BytesIO(content)) as image:
                if image.format != expected_format or image.width < 1 or image.height < 1:
                    _fail(code, f"{location} bytes do not match the declared raster media type")
                image.verify()
            with Image.open(BytesIO(content)) as image:
                for frame in ImageSequence.Iterator(image):
                    frame.load()
    except ResolvedNumberingPortError:
        raise
    except (
        Image.DecompressionBombError,
        Image.DecompressionBombWarning,
        UnidentifiedImageError,
        OSError,
        SyntaxError,
        ValueError,
    ) as exc:
        _fail(code, f"{location} is not a complete safe raster image")
        raise AssertionError from exc


__all__ = ["parse_document_extras"]
