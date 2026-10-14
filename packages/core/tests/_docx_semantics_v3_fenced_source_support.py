"""Shared constructors and exact package mutations for fenced-source tests."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any, Literal
from zipfile import ZipFile

import lxml.etree as etree
from docx.oxml.ns import qn

from docwen_core.docx_semantics_v3 import (
    FENCED_SOURCE_MAP_NAMESPACE,
    DocxSemanticsV3Session,
    FencedSourceIdentityV3,
    derive_fenced_source_identity_v3,
)


def _identity(
    block: str,
    logical_body: str,
    *,
    fence_character: str,
    opening_length: int,
    opening_prefix: str,
    info: str,
    opening_eol: str,
    body_prefixes: tuple[str, ...],
    closing_state: Literal["present", "omitted_eof"],
    closing_length: int,
    closing_prefix: str,
    closing_suffix: str,
    closing_eol: str,
    source_document: str | None = None,
    source_start: int = 0,
) -> FencedSourceIdentityV3:
    source = block if source_document is None else source_document
    return derive_fenced_source_identity_v3(
        source_sha256=hashlib.sha256(source.encode()).hexdigest(),
        source_start=source_start,
        source_end=source_start + len(block),
        block_sha256=hashlib.sha256(block.encode()).hexdigest(),
        body_sha256=hashlib.sha256(logical_body.encode()).hexdigest(),
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


def _write(session: DocxSemanticsV3Session, document: Any, output: Path) -> Path:
    session.finalize_document()
    document.save(str(output))
    session.write_package(output)
    session.prove_package(output)
    return output


def _map_part(package: ZipFile) -> str:
    return next(
        name
        for name in package.namelist()
        if name.startswith("customXml/item")
        and name.endswith(".xml")
        and FENCED_SOURCE_MAP_NAMESPACE.encode() in package.read(name)
    )


def _sdt_tag(sdt: etree._Element) -> str | None:
    tag = sdt.find(f"./{qn('w:sdtPr')}/{qn('w:tag')}")
    return None if tag is None else tag.get(qn("w:val"))


def _mutate_document_xml(path: Path, mutate) -> None:
    _mutate_xml_member(path, "word/document.xml", mutate)


def _mutate_xml_member(path: Path, member: str, mutate) -> None:
    with ZipFile(path) as package:
        infos = package.infolist()
        payloads = {item.filename: package.read(item.filename) for item in infos}
    root = etree.fromstring(payloads[member])
    mutate(root)
    payloads[member] = etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )
    temporary = path.with_suffix(".mutated.docx")
    with ZipFile(temporary, "w") as output:
        for info in infos:
            output.writestr(info, payloads[info.filename])
    temporary.replace(path)
