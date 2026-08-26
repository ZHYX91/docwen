"""Closed custom-XML codec for exact authored fenced-block occurrences."""

from __future__ import annotations

import base64
import binascii
from collections.abc import Mapping
from itertools import pairwise
from typing import Any

from docwen_core._docx_semantics_v3_fenced import (
    _BODY_PREFIX_BYTES_MAX,
    _BODY_PREFIX_COUNT_MAX,
    _MAX_DECIMAL,
    _SMALL_VALUE_BYTES_MAX,
    FENCED_SOURCE_MAP_NAMESPACE,
    FencedSourceIdentityV3,
    derive_fenced_source_identity_v3,
)
from docwen_core._docx_semantics_v3_model import DocxSemanticsV3Error

_XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
_RECORD_COUNT_MAX = 16_384
_MAP_DECODED_BYTES_MAX = 16_777_216
_MAP_ENCODED_BYTES_MAX = 25_165_824
_RECORD_ATTRIBUTES = (
    "tag",
    "source_sha256",
    "source_start",
    "source_end",
    "identity_sha256",
    "block_sha256",
    "body_sha256",
    "fence_character",
    "opening_length",
    "opening_prefix_b64",
    "info_b64",
    "opening_eol",
    "body_prefix_count",
    "body_prefixes_b64",
    "closing_state",
    "closing_length",
    "closing_prefix_b64",
    "closing_suffix_b64",
    "closing_eol",
)


def fenced_source_identity_from_mapping_v3(record: Mapping[str, Any]) -> FencedSourceIdentityV3:
    """Parse the exact source-oracle/wire spelling of one fenced record."""

    if set(record) != set(_RECORD_ATTRIBUTES):
        raise DocxSemanticsV3Error("fenced-source record fields are not closed")
    try:
        source_start = _canonical_decimal(record["source_start"], "source_start")
        source_end = _canonical_decimal(record["source_end"], "source_end")
        opening_length = _canonical_decimal(record["opening_length"], "opening_length")
        prefix_count = _canonical_decimal(record["body_prefix_count"], "body_prefix_count")
        closing_length = _canonical_decimal(record["closing_length"], "closing_length")
        opening_prefix = _decode_text_b64(record["opening_prefix_b64"], _SMALL_VALUE_BYTES_MAX)
        info = _decode_text_b64(record["info_b64"], _SMALL_VALUE_BYTES_MAX)
        body_prefixes = _decode_prefixes(record["body_prefixes_b64"], prefix_count)
        closing_prefix = _decode_text_b64(record["closing_prefix_b64"], _SMALL_VALUE_BYTES_MAX)
        closing_suffix = _decode_text_b64(record["closing_suffix_b64"], _SMALL_VALUE_BYTES_MAX)
    except (KeyError, TypeError) as exc:  # pragma: no cover - closed fields above
        raise DocxSemanticsV3Error("fenced-source record scalar has the wrong type") from exc
    identity = derive_fenced_source_identity_v3(
        source_sha256=_require_text(record["source_sha256"]),
        source_start=source_start,
        source_end=source_end,
        block_sha256=_require_text(record["block_sha256"]),
        body_sha256=_require_text(record["body_sha256"]),
        fence_character=_require_text(record["fence_character"]),
        opening_length=opening_length,
        opening_prefix=opening_prefix,
        info=info,
        opening_eol=_require_text(record["opening_eol"]),
        body_prefixes=body_prefixes,
        closing_state=_require_text(record["closing_state"]),  # type: ignore[arg-type]
        closing_length=closing_length,
        closing_prefix=closing_prefix,
        closing_suffix=closing_suffix,
        closing_eol=_require_text(record["closing_eol"]),
    )
    if record["identity_sha256"] != identity.identity_sha256 or record["tag"] != identity.tag:
        raise DocxSemanticsV3Error("fenced-source identity or tag does not recompute")
    return identity


def fenced_source_mapping_v3(identity: FencedSourceIdentityV3) -> dict[str, str | int]:
    """Return the canonical closed source-oracle spelling."""

    return {
        "tag": identity.tag,
        "source_sha256": identity.source_sha256,
        "source_start": identity.source_start,
        "source_end": identity.source_end,
        "identity_sha256": identity.identity_sha256,
        "block_sha256": identity.block_sha256,
        "body_sha256": identity.body_sha256,
        "fence_character": identity.fence_character,
        "opening_length": identity.opening_length,
        "opening_prefix_b64": _encode_text_b64(identity.opening_prefix),
        "info_b64": _encode_text_b64(identity.info),
        "opening_eol": identity.opening_eol,
        "body_prefix_count": len(identity.body_prefixes),
        "body_prefixes_b64": _encode_prefixes(identity.body_prefixes),
        "closing_state": identity.closing_state,
        "closing_length": identity.closing_length,
        "closing_prefix_b64": _encode_text_b64(identity.closing_prefix),
        "closing_suffix_b64": _encode_text_b64(identity.closing_suffix),
        "closing_eol": identity.closing_eol,
    }


def fenced_source_map_xml(records: list[FencedSourceIdentityV3]) -> bytes:
    """Serialize one canonical closed custom-XML map."""

    _prove_record_inventory(records)
    entries = "".join(_fenced_source_xml_record(item) for item in records)
    root = (
        f'<documentFencedSourceMap xmlns="{FENCED_SOURCE_MAP_NAMESPACE}" version="1">'
        f"{entries}</documentFencedSourceMap>"
    )
    return f"{_XML_DECLARATION}\n{root}\n".encode()


def parse_fenced_source_map(root: Any) -> list[FencedSourceIdentityV3]:
    """Parse and rederive every record from a closed XML topology."""

    namespace = f"{{{FENCED_SOURCE_MAP_NAMESPACE}}}"
    if (
        root.tag != f"{namespace}documentFencedSourceMap"
        or tuple(root.attrib.items()) != (("version", "1"),)
        or root.text is not None
        or root.tail is not None
    ):
        raise DocxSemanticsV3Error("fenced-source root is not closed and canonical")
    if len(root) > _RECORD_COUNT_MAX:
        raise DocxSemanticsV3Error("fenced-source record count exceeds the closed bound")
    records: list[FencedSourceIdentityV3] = []
    encoded_payload_bytes = 0
    for item in root:
        if (
            item.tag != f"{namespace}fencedSource"
            or tuple(item.attrib) != _RECORD_ATTRIBUTES
            or item.text is not None
            or item.tail is not None
            or len(item) != 0
        ):
            raise DocxSemanticsV3Error("fenced-source record is not closed and canonical")
        encoded_payload_bytes += sum(
            len(item.get(name, ""))
            for name in (
                "opening_prefix_b64",
                "info_b64",
                "body_prefixes_b64",
                "closing_prefix_b64",
                "closing_suffix_b64",
            )
        )
        if encoded_payload_bytes > _MAP_ENCODED_BYTES_MAX:
            raise DocxSemanticsV3Error("fenced-source map payload exceeds the closed bound")
        records.append(fenced_source_identity_from_mapping_v3(item.attrib))
    _prove_record_inventory(records)
    return records


def _fenced_source_xml_record(identity: FencedSourceIdentityV3) -> str:
    mapping = fenced_source_mapping_v3(identity)
    attributes = " ".join(f'{name}="{_xml_attr(str(value))}"' for name, value in mapping.items())
    return f"<fencedSource {attributes}/>"


def _prove_record_inventory(records: list[FencedSourceIdentityV3]) -> None:
    if not records:
        raise DocxSemanticsV3Error("fenced-source map requires at least one occurrence")
    if len(records) > _RECORD_COUNT_MAX:
        raise DocxSemanticsV3Error("fenced-source record count exceeds the closed bound")
    expected = sorted(records, key=lambda item: (item.source_start, item.source_end, item.tag))
    if records != expected:
        raise DocxSemanticsV3Error("fenced-source records are not canonically ordered")
    if len({item.tag for item in records}) != len(records):
        raise DocxSemanticsV3Error("fenced-source tags are not unique")
    if len({item.source_sha256 for item in records}) > 1:
        raise DocxSemanticsV3Error("fenced-source records do not share one source identity")
    decoded_payload_bytes = sum(
        len(value.encode("utf-8"))
        for item in records
        for value in (
            item.opening_prefix,
            item.info,
            *item.body_prefixes,
            item.closing_prefix,
            item.closing_suffix,
        )
    )
    if decoded_payload_bytes > _MAP_DECODED_BYTES_MAX:
        raise DocxSemanticsV3Error("fenced-source map payload exceeds the closed bound")
    for previous, current in pairwise(records):
        if current.source_start < previous.source_end:
            raise DocxSemanticsV3Error("fenced-source ranges overlap")


def _encode_text_b64(value: str) -> str:
    return base64.b64encode(value.encode("utf-8")).decode("ascii")


def _encode_prefixes(prefixes: tuple[str, ...]) -> str:
    return base64.b64encode(b"\0".join(item.encode("utf-8") for item in prefixes)).decode("ascii")


def _decode_text_b64(value: Any, maximum: int) -> str:
    if not isinstance(value, str):
        raise DocxSemanticsV3Error("fenced-source base64 scalar must be text")
    try:
        encoded = value.encode("ascii")
        if len(encoded) > ((maximum + 2) // 3) * 4:
            raise DocxSemanticsV3Error("fenced-source base64 exceeds the closed bound")
        decoded = base64.b64decode(encoded, validate=True)
        text = decoded.decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError, binascii.Error) as exc:
        raise DocxSemanticsV3Error("fenced-source base64 is not canonical UTF-8 RFC 4648") from exc
    if len(decoded) > maximum or base64.b64encode(decoded) != encoded:
        raise DocxSemanticsV3Error("fenced-source base64 exceeds bounds or is not canonical")
    return text


def _decode_prefixes(value: Any, count: int) -> tuple[str, ...]:
    text = _decode_text_b64(value, _BODY_PREFIX_BYTES_MAX)
    if count < 0 or count > _BODY_PREFIX_COUNT_MAX:
        raise DocxSemanticsV3Error("fenced-source body-prefix count exceeds the closed bound")
    if count == 0:
        if text:
            raise DocxSemanticsV3Error("zero fenced-source body prefixes have a non-empty payload")
        return ()
    prefixes = tuple(text.split("\0"))
    if len(prefixes) != count:
        raise DocxSemanticsV3Error("fenced-source body-prefix count does not match its payload")
    return prefixes


def _canonical_decimal(value: Any, name: str) -> int:
    if isinstance(value, bool):
        raise DocxSemanticsV3Error(f"fenced-source {name} is not a canonical decimal")
    if isinstance(value, int):
        if value < 0 or value > _MAX_DECIMAL:
            raise DocxSemanticsV3Error(f"fenced-source {name} exceeds the closed decimal bound")
        return value
    if (
        not isinstance(value, str)
        or not value
        or len(value) > 19
        or (value != "0" and value.startswith("0"))
        or not value.isascii()
    ):
        raise DocxSemanticsV3Error(f"fenced-source {name} is not a canonical decimal")
    if not value.isdecimal():
        raise DocxSemanticsV3Error(f"fenced-source {name} is not a canonical decimal")
    parsed = int(value)
    if parsed > _MAX_DECIMAL:
        raise DocxSemanticsV3Error(f"fenced-source {name} exceeds the closed decimal bound")
    return parsed


def _require_text(value: Any) -> str:
    if not isinstance(value, str):
        raise DocxSemanticsV3Error("fenced-source scalar must be text")
    return value


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
    "fenced_source_identity_from_mapping_v3",
    "fenced_source_map_xml",
    "fenced_source_mapping_v3",
    "parse_fenced_source_map",
]
