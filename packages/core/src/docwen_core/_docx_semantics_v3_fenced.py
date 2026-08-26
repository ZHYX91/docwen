"""Lossless, source-bound fenced-block carrier for Markdown semantics v3."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from typing import Any, Literal

from docwen_core._docx_semantics_v3_model import (
    DocxSemanticsV3Error,
    require_sha256,
)

FENCED_SOURCE_MAP_NAMESPACE = "https://docwen.dev/schema/document-fenced-source-map/v1"
FENCED_SOURCE_TAG_PREFIX = "docwen-fenced-source-v1:"

_IDENTITY_DOMAIN = "docwen-fenced-source-map-v1"
_SMALL_VALUE_BYTES_MAX = 16_384
_BODY_PREFIX_BYTES_MAX = 1_048_576
_BODY_PREFIX_COUNT_MAX = 65_536
_MAX_DECIMAL = 9_223_372_036_854_775_807

type ClosingStateV3 = Literal["present", "omitted_eof"]


@dataclass(frozen=True, slots=True)
class FencedSourceIdentityV3:
    """One exact authored fence occurrence, without semantic-target identity."""

    tag: str
    source_sha256: str
    source_start: int
    source_end: int
    identity_sha256: str
    block_sha256: str
    body_sha256: str
    fence_character: str
    opening_length: int
    opening_prefix: str
    info: str
    opening_eol: str
    body_prefixes: tuple[str, ...]
    closing_state: ClosingStateV3
    closing_length: int
    closing_prefix: str
    closing_suffix: str
    closing_eol: str


@dataclass(frozen=True, slots=True)
class FencedSourceBindingV3:
    identity: FencedSourceIdentityV3
    paragraph_element: Any


def derive_fenced_source_identity_v3(
    *,
    source_sha256: str,
    source_start: int,
    source_end: int,
    block_sha256: str,
    body_sha256: str,
    fence_character: str,
    opening_length: int,
    opening_prefix: str,
    info: str,
    opening_eol: str,
    body_prefixes: tuple[str, ...],
    closing_state: ClosingStateV3,
    closing_length: int,
    closing_prefix: str,
    closing_suffix: str,
    closing_eol: str,
) -> FencedSourceIdentityV3:
    """Validate closed framing scalars and derive the occurrence identity."""

    require_sha256(source_sha256)
    require_sha256(block_sha256)
    require_sha256(body_sha256)
    if any(type(value) is not int for value in (source_start, source_end, opening_length, closing_length)):
        raise DocxSemanticsV3Error("fenced-source numeric scalars must be integers")
    if source_start < 0 or source_end <= source_start:
        raise DocxSemanticsV3Error("fenced-source range must be non-empty and ordered")
    if source_end > _MAX_DECIMAL:
        raise DocxSemanticsV3Error("fenced-source range exceeds the closed decimal bound")
    if fence_character not in {"`", "~"} or opening_length < 3:
        raise DocxSemanticsV3Error("fenced-source opener is not a CommonMark fence")
    if opening_length > 65_536:
        raise DocxSemanticsV3Error("fenced-source opening length exceeds the closed bound")
    if fence_character == "`" and "`" in info:
        raise DocxSemanticsV3Error("backtick fenced-source info contains a backtick")
    body_prefixes = _require_body_prefixes(body_prefixes)
    _require_line_fragment(opening_prefix, "opening prefix", _SMALL_VALUE_BYTES_MAX)
    _require_line_fragment(info, "info string", _SMALL_VALUE_BYTES_MAX)
    _require_line_fragment(closing_prefix, "closing prefix", _SMALL_VALUE_BYTES_MAX)
    _require_line_fragment(closing_suffix, "closing suffix", _SMALL_VALUE_BYTES_MAX)
    if closing_suffix.strip(" \t"):
        raise DocxSemanticsV3Error("fenced-source closing suffix must contain spaces or tabs only")
    for prefix in (opening_prefix, closing_prefix, *body_prefixes):
        if any(character not in " \t>+-*0123456789.)[]xX" for character in prefix):
            raise DocxSemanticsV3Error("fenced-source container prefix contains a non-container character")
    _require_eol(opening_eol, context="opening")
    _require_eol(closing_eol, context="closing")
    if closing_state == "present":
        if closing_length < opening_length or opening_eol == "":
            raise DocxSemanticsV3Error("present fenced-source closer is shorter than its opener")
        if closing_length > 65_536:
            raise DocxSemanticsV3Error("fenced-source closing length exceeds the closed bound")
    elif closing_state == "omitted_eof":
        if closing_length != 0 or closing_prefix or closing_suffix or closing_eol:
            raise DocxSemanticsV3Error("omitted-EOF fenced source must not synthesize a closer")
    else:
        raise DocxSemanticsV3Error("fenced-source closing state is not closed")
    if body_prefixes and not opening_eol:
        raise DocxSemanticsV3Error("fenced-source body requires an opening EOL")
    preimage = f"{_IDENTITY_DOMAIN}\0{source_sha256}\0{source_start}\0{source_end}\0{block_sha256}\0{body_sha256}"
    identity_sha256 = hashlib.sha256(preimage.encode("utf-8")).hexdigest()
    return FencedSourceIdentityV3(
        tag=f"{FENCED_SOURCE_TAG_PREFIX}{identity_sha256[:32]}",
        source_sha256=source_sha256,
        source_start=source_start,
        source_end=source_end,
        identity_sha256=identity_sha256,
        block_sha256=block_sha256,
        body_sha256=body_sha256,
        fence_character=fence_character,
        opening_length=opening_length,
        opening_prefix=opening_prefix,
        info=info,
        opening_eol=opening_eol,
        body_prefixes=body_prefixes,
        closing_state=closing_state,
        closing_length=closing_length,
        closing_prefix=closing_prefix,
        closing_suffix=closing_suffix,
        closing_eol=closing_eol,
    )


def wrap_fenced_paragraph_payload(paragraph_element: Any, tag: str) -> Any:
    """Wrap every visible code payload run in one exact inline SDT."""

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    if paragraph_element.tag != qn("w:p"):
        raise DocxSemanticsV3Error("fenced-source carrier requires one Word paragraph")
    if paragraph_element.find(f"./{qn('w:sdt')}") is not None:
        raise DocxSemanticsV3Error("fenced-source paragraph already contains an inline SDT")
    children = list(paragraph_element)
    start = 1 if children and children[0].tag == qn("w:pPr") else 0
    payload = children[start:]
    if any(child.tag != qn("w:r") for child in payload):
        raise DocxSemanticsV3Error("fenced-source payload must contain runs only")
    sdt = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    tag_element = OxmlElement("w:tag")
    tag_element.set(qn("w:val"), tag)
    properties.append(tag_element)
    content = OxmlElement("w:sdtContent")
    for child in payload:
        paragraph_element.remove(child)
        content.append(child)
    sdt.extend((properties, content))
    paragraph_element.append(sdt)
    return sdt


def write_canonical_fenced_body_v3(paragraph_element: Any, logical_body: str) -> None:
    """Replace a code paragraph payload with the closed text/tab/break topology."""

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    if paragraph_element.tag != qn("w:p"):
        raise DocxSemanticsV3Error("fenced-source carrier requires one Word paragraph")
    children = list(paragraph_element)
    start = 1 if children and children[0].tag == qn("w:pPr") else 0
    if any(child.tag != qn("w:r") for child in children[start:]):
        raise DocxSemanticsV3Error("fenced-source paragraph contains non-run payload")
    for child in children[start:]:
        paragraph_element.remove(child)
    text_buffer: list[str] = []

    def flush_text() -> None:
        if not text_buffer:
            return
        value = "".join(text_buffer)
        text_buffer.clear()
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        if value[:1].isspace() or value[-1:].isspace():
            text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = value
        run.append(text)
        paragraph_element.append(run)

    try:
        index = 0
        while index < len(logical_body):
            character = logical_body[index]
            if character not in {"\r", "\n", "\t"}:
                text_buffer.append(character)
                index += 1
                continue
            flush_text()
            run = OxmlElement("w:r")
            if character == "\r":
                if index + 1 >= len(logical_body) or logical_body[index + 1] != "\n":
                    raise DocxSemanticsV3Error("fenced-source body contains a bare carriage return")
                payload_name = "w:cr"
                index += 2
            else:
                payload_name = "w:br" if character == "\n" else "w:tab"
                index += 1
            run.append(OxmlElement(payload_name))
            paragraph_element.append(run)
        flush_text()
    except (TypeError, ValueError) as exc:
        raise DocxSemanticsV3Error("fenced-source body contains an unrepresentable Word character") from exc


def prove_fenced_source_paragraph(
    paragraph_element: Any,
    identity: FencedSourceIdentityV3,
) -> str:
    """Authenticate the inline carrier and return the exact authored block."""

    from docx.oxml.ns import qn

    children = list(paragraph_element)
    payload = children[1:] if children and children[0].tag == qn("w:pPr") else children
    if len(payload) != 1 or payload[0].tag != qn("w:sdt"):
        raise DocxSemanticsV3Error("fenced-source SDT does not cover the complete paragraph payload")
    sdt = payload[0]
    if sdt.attrib or sdt.text is not None or sdt.tail is not None:
        raise DocxSemanticsV3Error("fenced-source SDT envelope is not canonical")
    sdt_children = list(sdt)
    if [item.tag for item in sdt_children] != [qn("w:sdtPr"), qn("w:sdtContent")]:
        raise DocxSemanticsV3Error("fenced-source SDT envelope is not canonical")
    _prove_single_tag(sdt_children[0], identity.tag)
    content = sdt_children[1]
    if content.attrib or content.text is not None or content.tail is not None:
        raise DocxSemanticsV3Error("fenced-source SDT content is not canonical")
    logical_body = _prove_and_read_body_runs(content)
    if hashlib.sha256(logical_body.encode("utf-8")).hexdigest() != identity.body_sha256:
        raise DocxSemanticsV3Error("fenced-source visible body hash does not match its map record")
    authored = reconstruct_fenced_source_v3(identity, logical_body)
    if len(authored) != identity.source_end - identity.source_start:
        raise DocxSemanticsV3Error("fenced-source block length does not match its source range")
    if hashlib.sha256(authored.encode("utf-8")).hexdigest() != identity.block_sha256:
        raise DocxSemanticsV3Error("fenced-source reconstructed block hash does not match its map record")
    if (
        paragraph_element.find(f".//{qn('w:bookmarkStart')}") is not None
        or paragraph_element.find(f".//{qn('w:bookmarkEnd')}") is not None
        or paragraph_element.find(f".//{qn('w:instrText')}") is not None
        or paragraph_element.find(f".//{qn('w:fldChar')}") is not None
    ):
        raise DocxSemanticsV3Error("fenced-source carrier must not contain bookmark, SEQ, or REF machinery")
    return authored


def bind_fenced_source_document_v3(
    body: Any,
    records: list[FencedSourceIdentityV3],
) -> dict[Any, tuple[FencedSourceIdentityV3, str]]:
    """Bind every map record to exactly one inline carrier in physical order."""

    from docx.oxml.ns import qn

    by_tag = {item.tag: item for item in records}
    seen: set[str] = set()
    physical_tags: list[str] = []
    output: dict[Any, tuple[FencedSourceIdentityV3, str]] = {}
    for sdt in body.iter(qn("w:sdt")):
        tag = _sdt_tag(sdt)
        if not tag or not tag.startswith(FENCED_SOURCE_TAG_PREFIX):
            continue
        record = by_tag.get(tag)
        if record is None or tag in seen:
            raise DocxSemanticsV3Error("fenced-source SDT is missing, duplicated, or unmapped")
        paragraph = sdt.getparent()
        if paragraph is None or paragraph.tag != qn("w:p"):
            raise DocxSemanticsV3Error("fenced-source SDT must be a direct paragraph child")
        physical_tags.append(tag)
        authored = prove_fenced_source_paragraph(paragraph, record)
        output[paragraph] = (record, authored)
        seen.add(tag)
    if seen != set(by_tag):
        raise DocxSemanticsV3Error("fenced-source map record is missing its exact inline SDT")
    expected_tags = [
        item.tag for item in sorted(records, key=lambda item: (item.source_start, item.source_end, item.tag))
    ]
    if physical_tags != expected_tags:
        raise DocxSemanticsV3Error("fenced-source physical order differs from authenticated source order")
    return output


def reconstruct_fenced_source_v3(identity: FencedSourceIdentityV3, logical_body: str) -> str:
    """Reconstruct exact authored Markdown from visible body plus framing."""

    if identity.body_prefixes:
        lines = _split_logical_body(logical_body)
        if len(lines) != len(identity.body_prefixes):
            raise DocxSemanticsV3Error("fenced-source body line count differs from its prefix inventory")
    elif logical_body:
        raise DocxSemanticsV3Error("fenced-source body exists without a prefix inventory")
    else:
        lines = ()
    opener = (
        identity.opening_prefix
        + identity.fence_character * identity.opening_length
        + identity.info
        + identity.opening_eol
    )
    if identity.closing_state == "present":
        if logical_body and not logical_body.endswith(("\n", "\r\n")):
            raise DocxSemanticsV3Error("fenced-source body does not end before its present closer")
        body_parts = [prefix + line for prefix, line in zip(identity.body_prefixes, lines, strict=True)]
        closer = (
            identity.closing_prefix
            + identity.fence_character * identity.closing_length
            + identity.closing_suffix
            + identity.closing_eol
        )
        return opener + "".join(body_parts) + closer
    if not lines:
        return opener
    body_parts = [prefix + line for prefix, line in zip(identity.body_prefixes, lines, strict=True)]
    return opener + "".join(body_parts)


def _split_logical_body(logical_body: str) -> tuple[str, ...]:
    """Split exact body lines while retaining every LF/CRLF spelling."""

    output: list[str] = []
    start = 0
    index = 0
    while index < len(logical_body):
        if logical_body[index] == "\r":
            if index + 1 >= len(logical_body) or logical_body[index + 1] != "\n":
                raise DocxSemanticsV3Error("fenced-source body contains a bare carriage return")
            output.append(logical_body[start : index + 2])
            index += 2
            start = index
            continue
        if logical_body[index] == "\n":
            output.append(logical_body[start : index + 1])
            index += 1
            start = index
            continue
        index += 1
    if start < len(logical_body):
        output.append(logical_body[start:])
    return tuple(output)


def _prove_and_read_body_runs(content: Any) -> str:
    from docx.oxml.ns import qn

    output: list[str] = []
    previous_was_text = False
    for run in content:
        if run.tag != qn("w:r") or run.attrib or run.xpath("text()") or run.tail is not None:
            raise DocxSemanticsV3Error("fenced-source content contains a non-canonical run")
        payload = list(run)
        if len(payload) != 1 or payload[0].tag not in {
            qn("w:t"),
            qn("w:br"),
            qn("w:cr"),
            qn("w:tab"),
        }:
            raise DocxSemanticsV3Error("fenced-source run has non-text payload")
        item = payload[0]
        if item.tag in {qn("w:br"), qn("w:cr"), qn("w:tab")}:
            if item.attrib or item.text is not None or item.tail is not None or len(item) != 0:
                raise DocxSemanticsV3Error("fenced-source break/tab is not canonical")
            output.append("\n" if item.tag == qn("w:br") else "\r\n" if item.tag == qn("w:cr") else "\t")
            previous_was_text = False
            continue
        text = item.text or ""
        if not text or any(character in {"\r", "\n", "\t"} for character in text) or previous_was_text:
            raise DocxSemanticsV3Error("fenced-source text runs are not canonical")
        expected_attributes = (
            (("{http://www.w3.org/XML/1998/namespace}space", "preserve"),)
            if text[:1].isspace() or text[-1:].isspace()
            else ()
        )
        if tuple(item.attrib.items()) != expected_attributes or item.tail is not None or len(item) != 0:
            raise DocxSemanticsV3Error("fenced-source text run is not canonical")
        output.append(text)
        previous_was_text = True
    return "".join(output)


def _prove_single_tag(properties: Any, tag: str) -> None:
    from docx.oxml.ns import qn

    children = list(properties)
    if (
        properties.attrib
        or properties.text is not None
        or properties.tail is not None
        or len(children) != 1
        or children[0].tag != qn("w:tag")
        or tuple(children[0].attrib.items()) != ((qn("w:val"), tag),)
        or children[0].text is not None
        or children[0].tail is not None
        or len(children[0]) != 0
    ):
        raise DocxSemanticsV3Error("fenced-source SDT properties are not canonical")


def _sdt_tag(sdt: Any) -> str | None:
    from docx.oxml.ns import qn

    properties = sdt.find(qn("w:sdtPr"))
    if properties is None:
        return None
    tags = properties.findall(qn("w:tag"))
    if len(tags) != 1:
        return None
    return tags[0].get(qn("w:val"))


def _require_body_prefixes(value: Any) -> tuple[str, ...]:
    if not isinstance(value, tuple) or len(value) > _BODY_PREFIX_COUNT_MAX:
        raise DocxSemanticsV3Error("fenced-source body line count exceeds the closed bound")
    prefixes: list[str] = []
    encoded_bytes = max(0, len(value) - 1)
    for prefix in value:
        _require_line_fragment(prefix, "body prefix", _SMALL_VALUE_BYTES_MAX)
        prefixes.append(prefix)
        encoded_bytes += len(prefix.encode("utf-8"))
    if encoded_bytes > _BODY_PREFIX_BYTES_MAX:
        raise DocxSemanticsV3Error("fenced-source body-prefix payload exceeds the closed bound")
    return tuple(prefixes)


def _require_line_fragment(value: Any, context: str, maximum: int) -> None:
    if not isinstance(value, str) or "\0" in value or "\r" in value or "\n" in value:
        raise DocxSemanticsV3Error(f"fenced-source {context} is not a single-line fragment")
    if len(value.encode("utf-8")) > maximum:
        raise DocxSemanticsV3Error(f"fenced-source {context} exceeds the closed bound")


def _require_eol(value: str, *, context: str) -> None:
    if value not in {"", "\n", "\r\n"}:
        raise DocxSemanticsV3Error(f"fenced-source {context} EOL is not closed")


__all__ = [
    "FENCED_SOURCE_MAP_NAMESPACE",
    "FENCED_SOURCE_TAG_PREFIX",
    "FencedSourceBindingV3",
    "FencedSourceIdentityV3",
    "bind_fenced_source_document_v3",
    "derive_fenced_source_identity_v3",
    "prove_fenced_source_paragraph",
    "reconstruct_fenced_source_v3",
    "wrap_fenced_paragraph_payload",
    "write_canonical_fenced_body_v3",
]
