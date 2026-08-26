"""Exact source-occurrence carrier helpers for v4 fenced Markdown blocks.

The semantic parser owns occurrence discovery and absolute source ranges.  This
module owns the closed, hash-bound projection record and the inverse proof used
by the runtime adapter.  Keeping those operations together prevents a DOCX
round trip from inventing fence syntax from a normalized Markdown AST.
"""

from __future__ import annotations

import base64
import binascii
import hashlib
from collections.abc import Mapping, Sequence
from typing import Any

_IDENTITY_DOMAIN = "docwen-fenced-source-map-v1"


def project_fenced_source_v3(
    source: str,
    *,
    source_sha256: str,
    source_start: int,
    source_end: int,
    fence: str,
    fence_start: int,
    opening_eol: str,
    body_line_coordinates: Sequence[tuple[int, int, int, int]],
    closing_present: bool,
    closing_fence_start: int,
    closing_fence_end: int,
    closing_line_start: int,
    closing_content_end: int,
    closing_eol: str,
) -> dict[str, Any]:
    """Build one canonical fenced-source record from parser-owned coordinates."""

    authored = source[source_start:source_end]
    opening_prefix = source[source_start:fence_start]
    info_start = fence_start + len(fence)
    opening_content_end = source.find(opening_eol, info_start, source_end) if opening_eol else source_end
    if opening_content_end < 0:
        raise ValueError("fenced opener EOL is not source-bound")
    info = source[info_start:opening_content_end]

    body_prefixes: list[str] = []
    logical_body_parts: list[str] = []
    for physical_start, logical_start, content_end, line_end in body_line_coordinates:
        body_prefixes.append(source[physical_start:logical_start])
        logical_body_parts.append(source[logical_start:content_end] + source[content_end:line_end])
    logical_body = "".join(logical_body_parts)

    closing_length = 0
    closing_prefix = ""
    closing_suffix = ""
    canonical_closing_eol = ""
    closing_state = "present" if closing_present else "omitted_eof"
    if closing_present:
        closing_length = closing_fence_end - closing_fence_start
        closing_prefix = source[closing_line_start:closing_fence_start]
        closing_suffix = source[closing_fence_end:closing_content_end]
        canonical_closing_eol = closing_eol

    block_sha256 = _sha256(authored)
    body_sha256 = _sha256(logical_body)
    identity_preimage = (
        f"{_IDENTITY_DOMAIN}\0{source_sha256}\0{source_start}\0{source_end}\0{block_sha256}\0{body_sha256}"
    )
    identity_sha256 = _sha256(identity_preimage)
    return {
        "tag": f"docwen-fenced-source-v1:{identity_sha256[:32]}",
        "source_sha256": source_sha256,
        "source_start": source_start,
        "source_end": source_end,
        "identity_sha256": identity_sha256,
        "block_sha256": block_sha256,
        "body_sha256": body_sha256,
        "fence_character": fence[0],
        "opening_length": len(fence),
        "opening_prefix_b64": _encode(opening_prefix),
        "info_b64": _encode(info),
        "opening_eol": opening_eol,
        "body_prefix_count": len(body_prefixes),
        "body_prefixes_b64": base64.b64encode(b"\0".join(item.encode("utf-8") for item in body_prefixes)).decode(
            "ascii"
        ),
        "closing_state": closing_state,
        "closing_length": closing_length,
        "closing_prefix_b64": _encode(closing_prefix),
        "closing_suffix_b64": _encode(closing_suffix),
        "closing_eol": canonical_closing_eol,
    }


def fenced_source_info_insertion_offset_v3(source: str, record: Mapping[str, Any]) -> int:
    """Return the authenticated zero-width marker position after exact info."""

    _require_authenticated_block(source, record)
    start = int(record["source_start"])
    opening_prefix = _decode_text(record, "opening_prefix_b64")
    info = _decode_text(record, "info_b64")
    opening_length = int(record["opening_length"])
    offset = start + len(opening_prefix) + opening_length + len(info)
    fence = str(record["fence_character"]) * opening_length
    if source[start:offset] != opening_prefix + fence + info:
        raise ValueError("fenced opener framing is not source-bound")
    return offset


def recover_fenced_logical_body_v3(source: str, record: Mapping[str, Any]) -> str:
    """Recover and authenticate the de-containerized logical fenced body."""

    block = _require_authenticated_block(source, record)
    opening_prefix = _decode_text(record, "opening_prefix_b64")
    info = _decode_text(record, "info_b64")
    opening_eol = str(record["opening_eol"])
    opening_size = len(opening_prefix) + int(record["opening_length"]) + len(info) + len(opening_eol)

    closing_size = 0
    if record["closing_state"] == "present":
        closing_size = (
            len(_decode_text(record, "closing_prefix_b64"))
            + int(record["closing_length"])
            + len(_decode_text(record, "closing_suffix_b64"))
            + len(str(record["closing_eol"]))
        )
    if opening_size + closing_size > len(block):
        raise ValueError("fenced framing exceeds the authenticated block")
    physical_body = block[opening_size : len(block) - closing_size if closing_size else None]
    physical_lines = physical_body.splitlines(keepends=True)
    prefixes = _decode_prefixes(record)
    if len(physical_lines) != len(prefixes):
        raise ValueError("fenced body prefix count does not match physical lines")
    logical_lines: list[str] = []
    for line, prefix in zip(physical_lines, prefixes, strict=True):
        if not line.startswith(prefix):
            raise ValueError("fenced body prefix is not source-bound")
        logical_lines.append(line[len(prefix) :])
    logical_body = "".join(logical_lines)
    if _sha256(logical_body) != record["body_sha256"]:
        raise ValueError("fenced logical body hash mismatch")
    return logical_body


def _require_authenticated_block(source: str, record: Mapping[str, Any]) -> str:
    if _sha256(source) != record["source_sha256"]:
        raise ValueError("fenced source record does not bind this source")
    start = int(record["source_start"])
    end = int(record["source_end"])
    if start < 0 or end <= start or end > len(source):
        raise ValueError("fenced source range is invalid")
    block = source[start:end]
    if _sha256(block) != record["block_sha256"]:
        raise ValueError("fenced block hash mismatch")
    return block


def _decode_prefixes(record: Mapping[str, Any]) -> tuple[str, ...]:
    count = int(record["body_prefix_count"])
    raw = _decode_bytes(record, "body_prefixes_b64")
    if count == 0:
        if raw:
            raise ValueError("zero fenced body prefixes require an empty payload")
        return ()
    parts = raw.split(b"\0")
    if len(parts) != count:
        raise ValueError("fenced body prefix payload count mismatch")
    try:
        return tuple(item.decode("utf-8") for item in parts)
    except UnicodeDecodeError as exc:
        raise ValueError("fenced body prefixes must be UTF-8") from exc


def _decode_text(record: Mapping[str, Any], key: str) -> str:
    try:
        return _decode_bytes(record, key).decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ValueError(f"{key} must encode UTF-8") from exc


def _decode_bytes(record: Mapping[str, Any], key: str) -> bytes:
    value = record[key]
    if not isinstance(value, str):
        raise ValueError(f"{key} must be a base64 string")
    try:
        return base64.b64decode(value, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ValueError(f"{key} is not canonical RFC 4648 base64") from exc


def _encode(value: str) -> str:
    return base64.b64encode(value.encode("utf-8")).decode("ascii")


def _sha256(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


__all__ = [
    "fenced_source_info_insertion_offset_v3",
    "project_fenced_source_v3",
    "recover_fenced_logical_body_v3",
]
