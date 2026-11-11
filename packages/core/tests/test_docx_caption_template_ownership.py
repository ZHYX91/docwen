"""Only generated caption groups are recovered from shared template styles."""

from __future__ import annotations

import hashlib
from copy import deepcopy
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from docwen_core._docx_semantics_v3_caption_groups import PAIRED_CAPTION_MAP_NAMESPACE, PAIRED_CAPTION_TAG_PREFIX
from docwen_core.docx_semantics_v3 import DocxSemanticsV3Error, DocxSemanticsV3Recovery, DocxSemanticsV3Session

from .test_docx_semantics_v3_caption_styles import _add_caption_styles, _add_table_caption
from .test_docx_semantics_v3_postreview import _mutate_xml_member

pytestmark = pytest.mark.contract


def _output(tmp_path: Path, *, template_has_seq: bool) -> tuple[Any, Path]:
    document = Document()
    bindings, styles = _add_caption_styles(document)
    if template_has_seq:
        fixed, _table = _add_table_caption(document, styles["table_caption"], "Fixed", number="9")
    else:
        fixed = document.add_paragraph("Fixed template caption", style=styles["table_caption"])
        document.add_table(rows=1, cols=1)
    original = etree.tostring(fixed._p, method="c14n")
    session = DocxSemanticsV3Session(
        document, source_sha256=hashlib.sha256(b"generated captions").hexdigest(), caption_style_bindings=bindings
    )
    caption, table = _add_table_caption(document, styles["table_caption"], "Generated")
    session.bind_ordinary_anchor((table._tbl,), {"block_kind": "table", "id": "generated-table"})
    session.bind_caption(caption, (table._tbl,), {"kind": "table", "number": "1", "title": "Generated"})
    named, other = _add_table_caption(document, styles["table_caption"], "Addressable", number="2")
    session.bind_caption(named, (other._tbl,), {"kind": "table", "id": "named", "number": "2", "title": "Addressable"})
    session.finalize_document()
    path = tmp_path / "caption-ownership.docx"
    document.save(str(path))
    session.write_package(path)
    assert etree.tostring(fixed._p, method="c14n") == original
    return session, path


@pytest.mark.parametrize("template_has_seq", [False, True])
def test_fixed_template_caption_is_not_claimed_as_generated(tmp_path: Path, template_has_seq: bool) -> None:
    session, path = _output(tmp_path, template_has_seq=template_has_seq)
    recovery = session.prove_package(path)
    assert recovery.caption_signatures == (("table", None, "Generated", "1"), ("table", "named", "Addressable", "2"))
    loaded = Document(str(path))
    reopened = DocxSemanticsV3Recovery.load(path, loaded)
    body = loaded.element.find(qn("w:body"))
    assert body is not None
    fixed = next(element for element in body if element.tag == qn("w:p"))
    assert not reopened.is_caption_element(fixed)
    assert fixed in reopened.logical_body_elements(loaded)
    generated = reopened.recovered_captions[0]
    anchor = reopened.source_anchor(generated.object_elements[0])
    assert anchor is not None and anchor.source_id == "generated-table"


@pytest.mark.parametrize("mutation", ["missing_seq", "unwrapped", "duplicate", "missing_map"])
def test_generated_idless_caption_damage_is_not_silently_ignored(tmp_path: Path, mutation: str) -> None:
    _session, path = _output(tmp_path, template_has_seq=False)
    if mutation == "missing_map":
        with ZipFile(path) as package:
            member = next(
                name
                for name in package.namelist()
                if name.startswith("customXml/item")
                and name.endswith(".xml")
                and etree.QName(etree.fromstring(package.read(name))).namespace == PAIRED_CAPTION_MAP_NAMESPACE
            )

        def mutate_map(root: Any) -> None:
            root.tag = "unrelatedMap"

        _mutate_xml_member(path, member, mutate_map)
    else:

        def mutate(root: Any) -> None:
            outer = next(
                node
                for node in root.iter(qn("w:sdt"))
                if (node.find(f"{qn('w:sdtPr')}/{qn('w:tag')}").get(qn("w:val"), "")).startswith(
                    PAIRED_CAPTION_TAG_PREFIX
                )
            )
            if mutation == "missing_seq":
                instruction = next(outer.iter(qn("w:instrText")))
                instruction.text = " PAGE "
            elif mutation == "duplicate":
                outer.addnext(deepcopy(outer))
            else:
                parent = outer.getparent()
                position = parent.index(outer)
                for child in reversed(list(outer.find(qn("w:sdtContent")))):
                    parent.insert(position, child)
                parent.remove(outer)

        _mutate_xml_member(path, "word/document.xml", mutate)
    with pytest.raises(DocxSemanticsV3Error):
        DocxSemanticsV3Recovery.load(path, Document(str(path)))


@pytest.mark.parametrize("extra", ["unbound_caption", "template_copy"])
def test_new_unbound_caption_cannot_inherit_template_allowance(tmp_path: Path, extra: str) -> None:
    document = Document()
    bindings, styles = _add_caption_styles(document)
    fixed, _fixed_table = _add_table_caption(document, styles["table_caption"], "Fixed")
    session = DocxSemanticsV3Session(
        document, source_sha256=hashlib.sha256(b"source").hexdigest(), caption_style_bindings=bindings
    )
    caption, table = _add_table_caption(document, styles["table_caption"], "Generated")
    session.bind_caption(caption, (table._tbl,), {"kind": "table", "number": "1", "title": "Generated"})
    if extra == "unbound_caption":
        _add_table_caption(document, styles["table_caption"], "Forgotten")
    else:
        caption._p.addprevious(deepcopy(fixed._p))
    with pytest.raises(DocxSemanticsV3Error, match="inventory differs from bound identities"):
        session.finalize_document()
    assert not list(tmp_path.iterdir())


def test_new_caption_after_capture_is_rejected_by_saved_session_proof(tmp_path: Path) -> None:
    session, path = _output(tmp_path, template_has_seq=True)
    reopened = Document(str(path))
    style = next(style for style in reopened.styles if style.style_id == "DocWenTableCaption")
    _add_table_caption(reopened, style, "Unexpected")
    reopened.save(str(path))
    with pytest.raises(DocxSemanticsV3Error, match="inventory differs after reopen"):
        session.prove_package(path)


def test_template_caption_change_after_capture_is_rejected(tmp_path: Path) -> None:
    session, path = _output(tmp_path, template_has_seq=False)

    def mutate(root: Any) -> None:
        fixed = root.find(f"{qn('w:body')}/{qn('w:p')}")
        next(fixed.iter(qn("w:t"))).text = "Changed after finalization"

    _mutate_xml_member(path, "word/document.xml", mutate)
    with pytest.raises(DocxSemanticsV3Error, match="template caption paragraph changed"):
        session.prove_package(path)
