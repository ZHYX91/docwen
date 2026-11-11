"""Saved Heading evidence retains the template's actual style identity."""

from __future__ import annotations

import hashlib
from copy import deepcopy
from pathlib import Path
from typing import Any, cast
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.styles.style import ParagraphStyle
from lxml import etree

from docwen_core._docx_semantics_v3_package import read_owned_map_parts
from docwen_core.docx_semantics_v3 import (
    HEADING_STYLE_BINDING_MAP_NAMESPACE,
    DocxSemanticsV3Error,
    DocxSemanticsV3Recovery,
    DocxSemanticsV3Session,
)

from .test_docx_semantics_v3_postreview import _mutate_xml_member

pytestmark = pytest.mark.contract


def _saved_heading(tmp_path: Path, style_id: str = '42_标题&"') -> tuple[Any, Path, bytes]:
    document = Document()
    style = cast(ParagraphStyle, document.styles["Heading 1"])
    style.element.set(qn("w:styleId"), style_id)
    name = style.element.find(qn("w:name"))
    assert name is not None
    name.set(qn("w:val"), "模板原有标题")
    aliases = OxmlElement("w:aliases")
    aliases.set(qn("w:val"), "标题旧称,别名")
    style.element.append(aliases)
    original = etree.tostring(style.element, method="c14n")
    fixed = document.add_paragraph("固定模板标题", style=style)
    fixed_properties = etree.tostring(fixed._p, method="c14n")
    session = DocxSemanticsV3Session(document, source_sha256=hashlib.sha256(b"# Generated ^intro").hexdigest())
    generated = document.add_paragraph("Generated", style=style)
    session.bind_heading(generated, {"kind": "heading", "id": "intro"}, heading_level=1)
    session.finalize_document()
    path = tmp_path / "heading.docx"
    document.save(str(path))
    session.write_package(path)
    assert etree.tostring(fixed._p, method="c14n") == fixed_properties
    return session, path, original


@pytest.mark.parametrize("style_id", ["42", "标题_原样式", '42_标题&"'])
def test_saved_heading_uses_recorded_template_style_identity(tmp_path: Path, style_id: str) -> None:
    session, path, original = _saved_heading(tmp_path, style_id)
    recovery = session.prove_package(path)
    binding = recovery.heading_style_bindings[0]
    assert (binding.source_id, binding.level, binding.resolved_style_id, binding.visible_name) == (
        "intro",
        1,
        style_id,
        "模板原有标题",
    )
    reopened = Document(str(path))
    matches = [style for style in reopened.styles.element if style.get(qn("w:styleId")) == style_id]
    assert len(matches) == 1
    assert etree.tostring(matches[0], method="c14n") == original
    assert len(DocxSemanticsV3Recovery.load(path, reopened).heading_style_bindings) == 1
    with ZipFile(path) as package:
        assert HEADING_STYLE_BINDING_MAP_NAMESPACE in read_owned_map_parts(package)


@pytest.mark.parametrize("mutation", ["paragraph", "name", "duplicate", "wrong_type"])
def test_heading_identity_tampering_is_rejected(tmp_path: Path, mutation: str) -> None:
    _session, path, _original = _saved_heading(tmp_path)

    def mutate(root: Any) -> None:
        if mutation == "paragraph":
            references = root.findall(f".//{qn('w:sdtContent')}/{qn('w:p')}/{qn('w:pPr')}/{qn('w:pStyle')}")
            assert len(references) == 1
            references[0].set(qn("w:val"), "Heading2")
            return
        style = next(item for item in root if item.get(qn("w:styleId")) == '42_标题&"')
        if mutation == "name":
            style.find(qn("w:name")).set(qn("w:val"), "replacement")
        elif mutation == "duplicate":
            root.append(deepcopy(style))
        else:
            style.set(qn("w:type"), "character")

    _mutate_xml_member(path, "word/document.xml" if mutation == "paragraph" else "word/styles.xml", mutate)
    with pytest.raises(DocxSemanticsV3Error, match="Heading"):
        DocxSemanticsV3Recovery.load(path, Document(str(path)))


def test_body_only_semantics_do_not_emit_heading_style_map(tmp_path: Path) -> None:
    document = Document()
    session = DocxSemanticsV3Session(document, source_sha256=hashlib.sha256(b"body ^body").hexdigest())
    session.bind_paragraph_anchor(document.add_paragraph("body"), {"block_kind": "paragraph", "id": "body"})
    session.finalize_document()
    path = tmp_path / "body.docx"
    document.save(str(path))
    session.write_package(path)
    session.prove_package(path)
    with ZipFile(path) as package:
        assert HEADING_STYLE_BINDING_MAP_NAMESPACE not in read_owned_map_parts(package)


def test_heading_source_binding_cannot_be_inferred_from_a_style_name() -> None:
    document = Document()
    normal = cast(ParagraphStyle, document.styles["Normal"])
    name = normal.element.find(qn("w:name"))
    assert name is not None
    name.set(qn("w:val"), "Heading 1")
    session = DocxSemanticsV3Session(document, source_sha256=hashlib.sha256(b"source").hexdigest())
    paragraph = document.add_paragraph("Not an outline heading", style=normal)
    with pytest.raises(DocxSemanticsV3Error, match=r"direct paragraph style|inherited outline level"):
        session.bind_heading(paragraph, {"kind": "heading", "id": "intro"})
