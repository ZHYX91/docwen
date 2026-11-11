"""DOCX identities follow internal relationships rather than conventional names."""

from __future__ import annotations

import posixpath
from io import BytesIO
from zipfile import ZipFile

import pytest
from docx import Document
from docx.opc.packuri import PackURI
from lxml import etree

from docwen_core.docx_package_parts import (
    CONTENT_TYPES_NS,
    HEADER,
    RELS_NS,
    DocxPackageParts,
    DocxPartError,
    document_story_parts,
    relationship_part_name,
    resolve_part_target,
)

pytestmark = pytest.mark.contract


def _package() -> dict[str, bytes]:
    document = Document()
    document.add_paragraph("Body")
    document.sections[0].header.paragraphs[0].text = "Header"
    document.sections[0].footer.paragraphs[0].text = "Footer"
    document.sections[0].header.part._partname = PackURI("/layout/top.xml")
    document.sections[0].footer.part._partname = PackURI("/archive/bottom.xml")
    stream = BytesIO()
    document.save(stream)
    with ZipFile(stream) as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    members["contents/report.xml"] = members.pop("word/document.xml")
    rels = etree.fromstring(members.pop("word/_rels/document.xml.rels"))
    for relation in rels:
        if relation.get("TargetMode") != "External":
            original = posixpath.normpath(posixpath.join("word", relation.get("Target")))
            relation.set("Target", posixpath.relpath(original, "contents"))
    members["contents/_rels/report.xml.rels"] = etree.tostring(rels)
    root_rels = etree.fromstring(members["_rels/.rels"])
    for relation in root_rels:
        if relation.get("Target") == "word/document.xml":
            relation.set("Target", "contents/report.xml")
    members["_rels/.rels"] = etree.tostring(root_rels)
    types = etree.fromstring(members["[Content_Types].xml"])
    for declaration in types:
        if declaration.get("PartName") == "/word/document.xml":
            declaration.set("PartName", "/contents/report.xml")
    members["[Content_Types].xml"] = etree.tostring(types)
    return members


def test_saved_and_loaded_story_identity_agree_with_owner_relative_targets() -> None:
    members = _package()
    before = dict(members)
    parts = DocxPackageParts(members)
    assert [part.name for part in parts.stories()] == ["contents/report.xml", "layout/top.xml", "archive/bottom.xml"]
    assert resolve_part_target("contents/report.xml", "../layout/top.xml") == "layout/top.xml"
    stream = BytesIO()
    with ZipFile(stream, "w") as archive:
        for name, content in members.items():
            archive.writestr(name, content)
    document = Document(stream)
    assert [str(part.partname) for part in document_story_parts(document)] == [
        "/contents/report.xml",
        "/layout/top.xml",
        "/archive/bottom.xml",
    ]
    assert members == before


@pytest.mark.parametrize(
    "target",
    ["../../outside.xml", "https://example.org/header.xml", "../%2e%2e/header.xml", "a%2fb.xml", "bad\\file.xml"],
)
def test_internal_part_target_rejects_unsafe_or_ambiguous_uris(target: str) -> None:
    with pytest.raises(DocxPartError):
        resolve_part_target("contents/report.xml", target)


@pytest.mark.parametrize("damage", ["external", "missing", "wrong_type", "duplicate_relation_id"])
def test_story_identity_rejects_ambiguous_or_misdeclared_related_parts(damage: str) -> None:
    members = _package()
    rel_name = relationship_part_name("contents/report.xml")
    rels = etree.fromstring(members[rel_name])
    relation = next(item for item in rels if item.get("Type") == HEADER.relationship_type)
    if damage == "external":
        relation.set("TargetMode", "External")
    elif damage == "missing":
        del members["layout/top.xml"]
    elif damage == "duplicate_relation_id":
        duplicate = etree.SubElement(rels, f"{{{RELS_NS}}}Relationship")
        duplicate.attrib.update(relation.attrib)
    else:
        types = etree.fromstring(members["[Content_Types].xml"])
        override = next(item for item in types if item.get("PartName") == "/layout/top.xml")
        assert override.tag == f"{{{CONTENT_TYPES_NS}}}Override"
        override.set("ContentType", "application/xml")
        members["[Content_Types].xml"] = etree.tostring(types)
    members[rel_name] = etree.tostring(rels)
    with pytest.raises(DocxPartError):
        DocxPackageParts(members).stories()


def test_orphan_conventional_header_cannot_gain_story_ownership() -> None:
    members = _package()
    members["word/header1.xml"] = members["layout/top.xml"]
    assert "word/header1.xml" not in {part.name for part in DocxPackageParts(members).stories()}
