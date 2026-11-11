"""Native list and heading writers extend the template's actual numbering part."""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree
from tests.support.numbering import repository_numbering_registry

from docwen_core.text.numbering_word_adapter import translate_scheme
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from docwen_plugin_markdown.to_docx.heading_numbering import (
    capture_heading_numbering_targets,
    write_heading_numbering_to_docx,
)

from .conftest import make_context

pytestmark = pytest.mark.contract

_NUMBERING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"


def _parts(path: Path | BytesIO) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _write(path: Path, parts: dict[str, bytes]) -> None:
    with ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _template(path: Path, name: str, target: str) -> dict[str, bytes]:
    document = Document()
    document.add_paragraph("{{body}}")
    memory = BytesIO()
    document.save(memory)
    parts = _parts(memory)
    parts[name] = parts.pop("word/numbering.xml")
    relationships = etree.fromstring(parts["word/_rels/document.xml.rels"])
    for relationship in relationships:
        if relationship.get("Type") == _NUMBERING_REL:
            relationship.set("Target", target)
    parts["word/_rels/document.xml.rels"] = etree.tostring(relationships)
    manifest = etree.fromstring(parts["[Content_Types].xml"])
    for override in manifest:
        if override.get("PartName") == "/word/numbering.xml":
            override.set("PartName", "/" + name)
    parts["[Content_Types].xml"] = etree.tostring(manifest)
    _write(path, parts)
    return parts


@pytest.mark.parametrize("route", ["heading", "list", "both"])
@pytest.mark.parametrize(
    ("name", "target"),
    [
        ("word/numbering2.xml", "numbering2.xml"),
        ("word/custom/lists.xml", "custom/lists.xml"),
        ("custom/lists.xml", "../custom/lists.xml"),
    ],
)
def test_converter_extends_only_the_relationship_owned_numbering_part(
    tmp_path: Path, route: str, name: str, target: str
) -> None:
    template = tmp_path / "template.docx"
    original_parts = _template(template, name, target)
    original_file = template.read_bytes()
    assert Document(str(template)).part.numbering_part.partname == "/" + name
    source = tmp_path / "source.md"
    heading = "# New heading\n\n" if route != "list" else ""
    items = "1. First\n2. Second\n" if route != "heading" else ""
    source.write_text(heading + items, encoding="utf-8")
    options: dict[str, str | bool] = {"template_name": str(template)}
    if route != "list":
        options.update(
            add_numbering=True,
            heading_numbering_render_mode="word_native",
            numbering_scheme="hierarchical_standard",
        )
    context, _workspace = make_context(str(source), options=options, numbering_registry=repository_numbering_registry())
    result = MdToDocxConverter().convert(context)
    assert result.success, result.error
    output = Path(result.artifacts[0].staging_path)
    parts = _parts(output)
    reopened = Document(str(output))
    assert reopened.part.numbering_part.partname == "/" + name
    assert "word/numbering.xml" not in parts
    relationships = etree.fromstring(parts["word/_rels/document.xml.rels"])
    assert [item.get("Target") for item in relationships if item.get("Type") == _NUMBERING_REL] == [target]
    original_numbering = etree.fromstring(original_parts[name])
    numbering = etree.fromstring(parts[name])
    for original in original_numbering:
        if original.tag not in {qn("w:abstractNum"), qn("w:num")}:
            continue
        identity = qn("w:abstractNumId" if original.tag == qn("w:abstractNum") else "w:numId")
        retained = [
            node for node in numbering if node.tag == original.tag and node.get(identity) == original.get(identity)
        ]
        assert len(retained) == 1
        assert etree.tostring(retained[0], method="c14n", exclusive=True) == etree.tostring(
            original, method="c14n", exclusive=True
        )
    available_ids = {item.get(qn("w:numId")) for item in numbering.findall(qn("w:num"))}
    generated_ids = {
        item.get(qn("w:val"))
        for paragraph in reopened.paragraphs
        for item in paragraph._p.findall(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}")
    }
    assert generated_ids and generated_ids <= available_ids
    if route == "both":
        assert len(generated_ids) == 2
    assert template.read_bytes() == original_file


def test_heading_writer_rejects_ambiguous_part_without_replacing_file(tmp_path: Path) -> None:
    document = Document()
    paragraph = document.add_heading("Generated", level=1)
    targets = capture_heading_numbering_targets(document, [(paragraph, 1)])
    output = tmp_path / "ambiguous.docx"
    document.save(str(output))
    parts = _parts(output)
    relationships = etree.fromstring(parts["word/_rels/document.xml.rels"])
    original = next(item for item in relationships if item.get("Type") == _NUMBERING_REL)
    duplicate = deepcopy(original)
    duplicate.set("Id", "rIdOtherNumbering")
    relationships.append(duplicate)
    parts["word/_rels/document.xml.rels"] = etree.tostring(relationships)
    _write(output, parts)
    before = output.read_bytes()
    with pytest.raises(ValueError):
        write_heading_numbering_to_docx(
            str(output), translate_scheme({"level_1": {"format": "{1.arabic_half} "}}), targets=targets
        )
    assert output.read_bytes() == before
