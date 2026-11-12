"""Request-owned native SEQ identifiers, separate from visible caption labels."""

from __future__ import annotations

import hashlib
import re
from typing import Any

from lxml import etree

_WORD_NAMESPACE = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_SCOPE_RE = re.compile(r"[0-9a-f]{24}\Z")
_IDENTIFIER_RE = re.compile(r"DW_([FTEC])_([0-9a-f]{24})\Z")
_KIND_CODES = {"figure": "F", "table": "T", "equation": "E", "code_block": "C"}
_SEQ_RE = re.compile(r'^\s*SEQ\s+(?:"([^"]+)"|([^\s\\]+))', re.IGNORECASE)


def caption_sequence_identifier(scope: str, kind: str) -> str:
    """Return the physical counter for a logical kind in one render request."""
    if not isinstance(scope, str) or _SCOPE_RE.fullmatch(scope) is None:
        raise ValueError("caption sequence scope must be 24 lowercase hexadecimal characters")
    if kind not in _KIND_CODES:
        raise ValueError("caption sequence kind is unsupported")
    return f"DW_{_KIND_CODES[kind]}_{scope}"


def caption_sequence_kind(identifier: str) -> str | None:
    """Recognize counter syntax without granting caption ownership."""
    match = _IDENTIFIER_RE.fullmatch(identifier)
    if match is None:
        return None
    return next(kind for kind, code in _KIND_CODES.items() if code == match[1])


def _sequence_identifiers(root: Any) -> set[str]:
    identifiers: set[str] = set()

    def collect(instruction: str) -> None:
        match = _SEQ_RE.match(instruction)
        if match is not None:
            identifiers.add((match[1] or match[2]).casefold())

    fields: list[list[str] | None] = []
    for element in root.iter():
        if element.tag == f"{_WORD_NAMESPACE}fldSimple":
            collect(element.get(f"{_WORD_NAMESPACE}instr", ""))
        elif element.tag == f"{_WORD_NAMESPACE}fldChar":
            field_type = element.get(f"{_WORD_NAMESPACE}fldCharType")
            if field_type == "begin":
                fields.append([])
            elif field_type in {"separate", "end"} and fields:
                if fields[-1] is not None:
                    collect("".join(fields[-1]))
                if field_type == "end":
                    fields.pop()
                else:
                    fields[-1] = None
        elif element.tag == f"{_WORD_NAMESPACE}instrText":
            text = element.text or ""
            collect(text)
            if fields and fields[-1] is not None:
                fields[-1].append(text)
    for fragments in fields:
        if fragments is not None:
            collect("".join(fragments))
    return identifiers


def allocate_caption_sequence_scope(document: Any, seed: str) -> str:
    """Choose a deterministic scope absent from every existing package story."""
    occupied: set[str] = set()
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    for part in document.part.package.parts:
        if not part.content_type.lower().endswith(("+xml", "/xml")):
            continue
        try:
            root = etree.fromstring(part.blob, parser)
        except etree.XMLSyntaxError:
            continue
        occupied.update(_sequence_identifiers(root))
    attempt = 0
    while True:
        scope = hashlib.sha256(f"docwen-caption-sequence-v1\0{seed}\0{attempt}".encode()).hexdigest()[:24]
        if all(caption_sequence_identifier(scope, kind).casefold() not in occupied for kind in _KIND_CODES):
            return scope
        attempt += 1
