"""Empty-field layout cleanup preserves original non-text document structures."""

from __future__ import annotations

from io import BytesIO
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from docwen_plugin_markdown.gongwen_fields import GongwenFieldPlacementError
from docwen_plugin_markdown.template_filler import fill_template
from docwen_plugin_markdown.template_utils import scan_placeholders

pytestmark = pytest.mark.contract

_EXTENSION = "{urn:docwen:test:template-fidelity}retain"
_EXTENSION_ATTRIBUTE = "{urn:docwen:test:template-fidelity}semanticToken"


def _save(document) -> tuple[object, dict[str, bytes]]:
    stream = BytesIO()
    document.save(stream)
    payload = stream.getvalue()
    with ZipFile(BytesIO(payload)) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    return Document(BytesIO(payload)), parts


def _fill(document) -> None:
    fill_template(document, {}, [], None, scan_placeholders(document))


def _bookmark(paragraph) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), "41")
    start.set(qn("w:name"), "TemplateTarget")
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), "41")
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def _content_control(node, *, block: bool = False):
    control = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    identity = OxmlElement("w:id")
    identity.set(qn("w:val"), "73")
    properties.append(identity)
    control.append(properties)
    content = OxmlElement("w:sdtContent")
    control.append(content)
    if block:
        node.addprevious(control)
    else:
        node.getparent().append(control)
    content.append(node)
    return control


@pytest.mark.parametrize("structure", ["bookmark", "comment", "inline-sdt", "block-sdt", "extension", "run-extension"])
def test_empty_field_keeps_semantic_content_in_saved_document(structure: str) -> None:
    document = Document()
    paragraph = document.add_paragraph("{{主送机关}}")
    if structure == "bookmark":
        _bookmark(paragraph)
        retained = paragraph._p.find(qn("w:bookmarkStart"))
    elif structure == "comment":
        document.add_comment(paragraph.runs[0], "Existing template comment", author="Template")
        retained = paragraph._p.find(qn("w:commentRangeStart"))
    elif structure in {"inline-sdt", "block-sdt"}:
        retained = _content_control(
            paragraph._p if structure == "block-sdt" else paragraph.runs[0]._r,
            block=structure == "block-sdt",
        ).find(qn("w:sdtPr"))
    else:
        retained = etree.Element(_EXTENSION, setting="keep")
        if structure == "extension":
            paragraph._p.append(retained)
        else:
            paragraph.runs[0]._r.get_or_add_rPr().append(retained)
    assert retained is not None
    expected = etree.tostring(retained, method="c14n")
    _, original_parts = _save(document)

    _fill(document)
    _, result_parts = _save(document)
    root = etree.fromstring(result_parts["word/document.xml"])
    kept = list(root.iter(retained.tag))
    assert len(kept) == 1
    assert etree.tostring(kept[0], method="c14n") == expected
    assert not any("{{" in (text.text or "") for text in root.iter(qn("w:t")))
    assert len(root.findall(f".//{qn('w:p')}")) == 1
    if structure == "bookmark":
        assert root.find(f".//{qn('w:bookmarkEnd')}") is not None
    if structure == "comment":
        assert root.find(f".//{qn('w:commentRangeEnd')}") is not None
        assert root.find(f".//{qn('w:commentReference')}") is not None
        assert result_parts["word/comments.xml"] == original_parts["word/comments.xml"]


def test_empty_field_row_keeps_anchors_in_another_cell() -> None:
    document = Document()
    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "{{抄送机关}}"
    _bookmark(table.cell(0, 1).paragraphs[0])

    _fill(document)
    _, parts = _save(document)
    root = etree.fromstring(parts["word/document.xml"])
    assert len(root.findall(f".//{qn('w:tr')}")) == 1
    assert root.find(f".//{qn('w:bookmarkStart')}") is not None
    assert root.find(f".//{qn('w:bookmarkEnd')}") is not None
    assert not any(text.text for text in root.iter(qn("w:t")))


def test_last_empty_row_does_not_discard_table_metadata() -> None:
    document = Document()
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "{{抄送机关}}"
    extension = etree.Element(_EXTENSION, setting="table metadata")
    table._tbl.tblPr.append(extension)

    _fill(document)
    _, parts = _save(document)
    root = etree.fromstring(parts["word/document.xml"])
    assert len(root.findall(f".//{qn('w:tbl')}")) == 1
    assert len(root.findall(f".//{qn('w:tr')}")) == 1
    assert root.find(f".//{_EXTENSION}") is not None
    assert not any(text.text for text in root.iter(qn("w:t")))


@pytest.mark.parametrize("location", ["paragraph", "run", "formatting", "row", "table", "table-properties"])
@pytest.mark.parametrize("attribute", [_EXTENSION_ATTRIBUTE, "semanticToken", qn("w14:semanticToken")])
def test_empty_field_keeps_unknown_attributes_in_saved_document(location: str, attribute: str) -> None:
    document = Document()
    if location in {"row", "table", "table-properties"}:
        table = document.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "{{抄送机关}}"
        retained = {"row": table.rows[0]._tr, "table": table._tbl, "table-properties": table._tbl.tblPr}[location]
    else:
        paragraph = document.add_paragraph("{{主送机关}}")
        if location == "formatting":
            paragraph.runs[0].bold = True
            properties = paragraph.runs[0]._r.rPr
            assert properties is not None
            retained = properties.find(qn("w:b"))
            assert retained is not None
        else:
            retained = paragraph._p if location == "paragraph" else paragraph.runs[0]._r
    retained.set(attribute, "template-owned")

    _fill(document)
    _, parts = _save(document)
    root = etree.fromstring(parts["word/document.xml"])
    matches = [node for node in root.iter() if node.get(attribute) == "template-owned"]
    assert len(matches) == 1
    assert matches[0].tag == retained.tag
    assert not any(text.text for text in root.iter(qn("w:t")))


def test_plain_empty_field_layout_still_disappears() -> None:
    document = Document()
    paragraph = document.add_paragraph("{{主送机关}}")
    paragraph.runs[0].bold = True
    paragraph.paragraph_format.keep_with_next = True
    paragraph._p.set(qn("w:rsidR"), "00112233")
    paragraph._p.set(qn("w14:paraId"), "00112234")
    paragraph._p.set(qn("w14:textId"), "00112235")
    text = paragraph.runs[0]._r.find(qn("w:t"))
    assert text is not None
    text.set(qn("xml:space"), "preserve")
    table = document.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "抄送：{{抄送机关}}"

    _fill(document)
    _, parts = _save(document)
    root = etree.fromstring(parts["word/document.xml"])
    assert root.find(f".//{qn('w:p')}") is None
    assert root.find(f".//{qn('w:tbl')}") is None


def test_attachment_expansion_refuses_unknown_content_before_clearing_it() -> None:
    document = Document()
    paragraph = document.add_paragraph("{{附件说明}}")
    paragraph.runs[0]._r.get_or_add_rPr().append(etree.Element(_EXTENSION, setting="retain"))
    original = document.element.xml

    with pytest.raises(GongwenFieldPlacementError):
        _fill(document)

    assert document.element.xml == original


@pytest.mark.parametrize(
    ("control", "offset"),
    [("w:tab", 1), ("w:tab", 2), ("w:tab", 4), ("w:tab", 7), ("w:cr", 4), ("w:drawing", 4), ("m:oMath", 4)],
)
def test_broken_reserved_attachment_marker_is_rejected_without_becoming_fillable(control: str, offset: int) -> None:
    document = Document()
    paragraph = document.add_paragraph()
    marker = "{{附件说明}}"
    paragraph.add_run(marker[:offset])._r.append(OxmlElement(control))
    paragraph.add_run(marker[offset:])
    placeholders = scan_placeholders(document)
    assert "附件说明" not in placeholders
    original = document.element.xml

    with pytest.raises(GongwenFieldPlacementError, match="附件说明占位符必须独占"):
        fill_template(document, {}, [], None, placeholders)

    assert document.element.xml == original


@pytest.mark.parametrize("wrapper", ["w:hyperlink", "w:fldSimple", "w:sdt"])
def test_reserved_attachment_marker_cannot_cross_semantic_wrappers(wrapper: str) -> None:
    document = Document()
    paragraph = document.add_paragraph("{{附件")
    suffix = paragraph.add_run("说明}}")
    if wrapper == "w:sdt":
        _content_control(suffix._r)
    else:
        container = OxmlElement(wrapper)
        if wrapper == "w:fldSimple":
            container.set(qn("w:instr"), "DATE")
        paragraph._p.append(container)
        container.append(suffix._r)
    original = document.element.xml

    with pytest.raises(GongwenFieldPlacementError, match="附件说明占位符必须独占"):
        _fill(document)

    assert document.element.xml == original


@pytest.mark.parametrize("container", ["paragraph", "table"])
def test_rendered_broken_attachment_marker_remains_literal(container: str) -> None:
    document = Document()
    body = document.add_paragraph("{{body}}")
    placeholders = scan_placeholders(document)
    if container == "table":
        rendered = document.add_table(rows=1, cols=1)
        paragraph = rendered.cell(0, 0).paragraphs[0]
    else:
        rendered = paragraph = document.add_paragraph()
    paragraph.add_run("{{附件").add_tab()
    paragraph.add_run("说明}}")

    fill_template(document, {}, [rendered], body, placeholders)
    _, parts = _save(document)
    root = etree.fromstring(parts["word/document.xml"])
    assert "".join(text.text or "" for text in root.iter(qn("w:t"))) == "{{附件说明}}"
    assert root.find(f".//{qn('w:tab')}") is not None


def test_reserved_attachment_marker_can_span_ordinary_formatting_runs() -> None:
    document = Document()
    paragraph = document.add_paragraph("{{ 附件")
    paragraph.add_run("说明 }}").bold = True

    fill_template(document, {"附件说明": ["材料"]}, [], None, scan_placeholders(document))
    _, parts = _save(document)
    root = etree.fromstring(parts["word/document.xml"])
    assert "".join(text.text or "" for text in root.iter(qn("w:t"))) == "附件：材料"
