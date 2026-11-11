"""Request-level v4 numbering DOCX session gates."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile, ZipInfo

import lxml.etree as etree
import pytest
from docx import Document
from docx.oxml.ns import qn

from docwen_core.docx_numbering_occurrence import NUMBERING_OCCURRENCE_MAP_NAMESPACE
from docwen_core.docx_resolved_numbering import (
    ResolvedNumberingDocxError,
    ResolvedNumberingDocxSession,
)

from ._docx_resolved_numbering_support import _caption_bindings, _mutate_zip_member, _port

pytestmark = pytest.mark.unit


def test_session_writes_heading_ref_and_disabled_idless_occurrence(tmp_path: Path) -> None:
    document = Document()
    bindings = _caption_bindings(document)
    session = ResolvedNumberingDocxSession(
        document,
        _port(),
        heading_style_ids={1: "Heading1"},
        heading_style_names={"heading_1": "Heading 1"},
        caption_style_bindings=bindings,
    )
    heading = document.add_heading("2.3 标题", level=1)
    reference = document.add_paragraph()
    caption = document.add_paragraph(style="Table Caption V4")
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "b"
    targets = session.port.document.targets
    session.bind_heading(
        heading,
        source_start=targets[0].source_start,
        source_end=targets[0].source_end,
    )
    session.bind_caption(
        caption,
        (table._element,),
        source_start=targets[1].source_start,
        source_end=targets[1].source_end,
        kind="table",
    )
    resolved_reference = session.port.document.references[0]
    session.render_reference(
        reference,
        source_start=resolved_reference.source_start,
        source_end=resolved_reference.source_end,
    )
    session.finalize_document()
    output = tmp_path / "resolved.docx"
    document.save(str(output))
    session.write_package(output)
    session.prove_package(output)

    assert heading.text == "2.3 标题"
    with ZipFile(output) as package:
        numbering = etree.fromstring(package.read("word/numbering.xml"))
        document_xml = etree.fromstring(package.read("word/document.xml"))
        assert any(
            NUMBERING_OCCURRENCE_MAP_NAMESPACE.encode() in package.read(name)
            for name in package.namelist()
            if name.startswith("customXml/item") and name.endswith(".xml")
        )
    assert numbering.find(qn("w:abstractNum")) is not None
    heading_paragraph = next(
        paragraph
        for paragraph in document_xml.iter(qn("w:p"))
        if "".join(item.text or "" for item in paragraph.iter(qn("w:t"))) == "2.3 标题"
    )
    assert heading_paragraph.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is not None
    assert not list(heading_paragraph.iter(qn("w:instrText")))
    instructions = [item.text for item in document_xml.iter(qn("w:instrText"))]
    assert any(value and " REF " in value and r"\n \h" in value for value in instructions)
    assert "1" in [item.text for item in document_xml.iter(qn("w:t"))]
    assert " 标题" in [item.text for item in document_xml.iter(qn("w:t"))]


def test_session_rejects_incomplete_bindings_before_package_mutation(tmp_path: Path) -> None:
    document = Document()
    session = ResolvedNumberingDocxSession(
        document,
        _port(),
        heading_style_ids={1: "Heading1"},
        heading_style_names={"heading_1": "Heading 1"},
        caption_style_bindings=_caption_bindings(document),
    )
    target = session.port.document.targets[0]
    session.bind_heading(
        document.add_heading("2.3 标题", level=1),
        source_start=target.source_start,
        source_end=target.source_end,
    )
    output = tmp_path / "must-not-exist.docx"

    with pytest.raises(ResolvedNumberingDocxError, match="not every resolved target"):
        session.write_package(output)
    assert not output.exists()


def test_session_never_uses_authored_heading_prefix_as_number() -> None:
    document = Document()
    session = ResolvedNumberingDocxSession(
        document,
        _port(),
        heading_style_ids={1: "Heading1"},
        heading_style_names={"heading_1": "Heading 1"},
        caption_style_bindings=_caption_bindings(document),
    )
    target = session.port.document.targets[0]
    heading = document.add_heading("2.3 标题", level=1)
    session.bind_heading(
        heading,
        source_start=target.source_start,
        source_end=target.source_end,
    )

    assert heading.text == "2.3 标题"
    assert all((item.text or "") != "1" for item in heading._p.iter(qn("w:t")))


def test_heading_merge_snapshot_preserves_authored_prefix_and_full_rendered_text(tmp_path: Path) -> None:
    document = Document()
    session = ResolvedNumberingDocxSession(
        document,
        _port(),
        heading_style_ids={1: "Heading1"},
        heading_style_names={"heading_1": "Heading 1"},
        caption_style_bindings=_caption_bindings(document),
    )
    heading = document.add_heading("2.3 标题：合并正文", level=1)
    targets = session.port.document.targets
    session.bind_heading(
        heading,
        source_start=targets[0].source_start,
        source_end=targets[0].source_end,
    )
    caption = document.add_paragraph(style="Table Caption V4")
    table = document.add_table(rows=1, cols=1)
    session.bind_caption(
        caption,
        (table._element,),
        source_start=targets[1].source_start,
        source_end=targets[1].source_end,
        kind="table",
    )
    reference = document.add_paragraph()
    resolved_reference = session.port.document.references[0]
    session.render_reference(
        reference,
        source_start=resolved_reference.source_start,
        source_end=resolved_reference.source_end,
    )
    output = tmp_path / "merged-heading.docx"
    session.write_package(output)
    session.prove_package(output)

    _replace_zip_member(
        output,
        "word/document.xml",
        "合并正文".encode(),
        "篡改正文".encode(),
    )
    with pytest.raises(ValueError, match="rendered text changed"):
        session.prove_package(output)


def test_heading_bind_rejects_rendered_text_without_authored_prefix() -> None:
    document = Document()
    session = ResolvedNumberingDocxSession(
        document,
        _port(),
        heading_style_ids={1: "Heading1"},
        heading_style_names={"heading_1": "Heading 1"},
        caption_style_bindings=_caption_bindings(document),
    )
    target = session.port.document.targets[0]

    with pytest.raises(ResolvedNumberingDocxError, match="authored title prefix"):
        session.bind_heading(
            document.add_heading("不同标题", level=1),
            source_start=target.source_start,
            source_end=target.source_end,
        )


@pytest.mark.parametrize(
    ("member", "old", "new"),
    [
        ("word/numbering.xml", None, None),
        ("word/document.xml", None, None),
        ("word/document.xml", b">" + "数据".encode() + b"<", b">" + "篡改".encode() + b"<"),
        ("word/document.xml", b">1<", b">9<"),
        ("word/document.xml", b"docwen-target-v1:", b"docwen-target-v2:"),
    ],
)
def test_session_reopen_proof_rejects_numbering_caption_ref_and_target_tamper(
    tmp_path: Path,
    member: str,
    old: bytes | None,
    new: bytes | None,
) -> None:
    session, output = _render_resolved_package(tmp_path)
    if member == "word/numbering.xml":
        _mutate_owned_number_format(output, session)
    elif old is None:
        owned_num_id = str(session.projection.num_id("document")).encode()
        _replace_zip_member(
            output,
            member,
            b'w:numId w:val="' + owned_num_id + b'"',
            b'w:numId w:val="99"',
        )
    else:
        assert new is not None
        _replace_zip_member(output, member, old, new)

    with pytest.raises(ValueError):
        session.prove_package(output)


def _render_resolved_package(tmp_path: Path) -> tuple[ResolvedNumberingDocxSession, Path]:
    document = Document()
    session = ResolvedNumberingDocxSession(
        document,
        _port(),
        heading_style_ids={1: "Heading1"},
        heading_style_names={"heading_1": "Heading 1"},
        caption_style_bindings=_caption_bindings(document),
    )
    heading = document.add_heading("2.3 标题", level=1)
    reference = document.add_paragraph()
    caption = document.add_paragraph(style="Table Caption V4")
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "b"
    targets = session.port.document.targets
    session.bind_heading(heading, source_start=targets[0].source_start, source_end=targets[0].source_end)
    session.bind_caption(
        caption,
        (table._element,),
        source_start=targets[1].source_start,
        source_end=targets[1].source_end,
        kind="table",
    )
    resolved_reference = session.port.document.references[0]
    session.render_reference(
        reference,
        source_start=resolved_reference.source_start,
        source_end=resolved_reference.source_end,
    )
    output = tmp_path / "resolved-proof.docx"
    document.save(str(output))
    session.write_package(output)
    return session, output


def _replace_zip_member(path: Path, member: str, old: bytes, new: bytes) -> None:
    with ZipFile(path) as package:
        infos = package.infolist()
        data = {info.filename: package.read(info.filename) for info in infos}
    assert old in data[member]
    data[member] = data[member].replace(old, new, 1)
    temporary = path.with_suffix(".rewrite.docx")
    with ZipFile(temporary, "w") as output:
        for info in infos:
            output.writestr(info, data[info.filename])
        known = {info.filename for info in infos}
        for name in sorted(set(data) - known):
            output.writestr(ZipInfo(name), data[name])
    temporary.replace(path)


def _mutate_owned_number_format(path: Path, session: ResolvedNumberingDocxSession) -> None:
    abstract_id = str(session.projection.abstract_id("main"))

    def mutate(payload: bytes) -> bytes:
        root = etree.fromstring(payload)
        candidates = [
            item for item in root.findall(qn("w:abstractNum")) if item.get(qn("w:abstractNumId")) == abstract_id
        ]
        assert len(candidates) == 1
        number_format = candidates[0].find("./" + qn("w:lvl") + "/" + qn("w:numFmt"))
        assert number_format is not None
        assert number_format.get(qn("w:val")) == "decimal"
        number_format.set(qn("w:val"), "upperRoman")
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    _mutate_zip_member(path, "word/numbering.xml", mutate)
