"""DOCX note text and hyperlink ownership through the production Markdown exporter."""

from zipfile import ZipFile

import pytest
from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from docwen_plugin_document.plugin import DocumentPlugin

from ._request_scoped_docx_policy_support import _context, _markdown_from_result

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("zip_fallback", [False, True])
def test_note_export_preserves_mixed_marker_run_and_uses_note_owned_links(tmp_path, kind, zip_fallback):
    document = Document()
    paragraph = document.add_paragraph("Body")
    marker = OxmlElement(f"w:{kind}Reference")
    marker.set(qn("w:id"), "7")
    paragraph.add_run()._r.append(marker)
    # An identically named relationship in the main story points somewhere else.
    document.part.load_rel(RT.HYPERLINK, "https://example.org/body", "rIdShared", is_external=True)
    root = OxmlElement(f"w:{kind}s")
    note = etree.SubElement(root, qn(f"w:{kind}"))
    note.set(qn("w:id"), "7")
    body = etree.SubElement(note, qn("w:p"))
    run = etree.SubElement(body, qn("w:r"))
    etree.SubElement(run, qn(f"w:{kind}Ref"))
    etree.SubElement(run, qn("w:t")).text = "Read"
    space = etree.SubElement(etree.SubElement(body, qn("w:r")), qn("w:t"))
    space.set(qn("xml:space"), "preserve")
    space.text = " "
    for rel_id, label in (("rIdShared", "guide"), ("rIdMissing", "More")):
        hyperlink = etree.SubElement(body, qn("w:hyperlink"))
        hyperlink.set(qn("r:id"), rel_id)
        etree.SubElement(etree.SubElement(hyperlink, qn("w:r")), qn("w:t")).text = label

    filename = f"{kind}s.xml" if zip_fallback else "custom-notes.xml"
    part = Part(
        PackURI("/word/" + filename),
        CT.WML_FOOTNOTES if kind == "footnote" else CT.WML_ENDNOTES,
        etree.tostring(root),
        document.part.package,
    )
    part.load_rel(RT.HYPERLINK, "https://example.org/guide#details", "rIdShared", is_external=True)
    if not zip_fallback:
        document.part.relate_to(part, RT.FOOTNOTES if kind == "footnote" else RT.ENDNOTES)
    source = tmp_path / "notes.docx"
    document.save(str(source))
    if zip_fallback:
        with ZipFile(source, "a") as package:
            package.writestr("word/" + filename, part.blob)
            package.writestr("word/_rels/" + filename + ".rels", part.rels.xml)

    context = _context(
        tmp_path,
        source,
        request_id="note-links",
        config={},
        options={"markdown_extensions": {"output": {"typed_endnotes": True}}},
    )
    markdown = _markdown_from_result(DocumentPlugin().convert(context))
    key = "1" if kind == "footnote" else "endnote:1"
    assert f"Body[^{key}]" in markdown
    assert f"[^{key}]: Read [guide](https://example.org/guide#details)More" in markdown
