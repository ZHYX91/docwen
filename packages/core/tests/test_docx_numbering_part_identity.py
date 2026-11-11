"""Numbering OPC ownership is defined by relationships, not a filename guess."""

from __future__ import annotations

from copy import deepcopy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

from docwen_core.docx_numbering_import import HeadingNumberingProofIndex
from docwen_core.docx_numbering_ooxml import (
    apply_heading_numbering,
    create_heading_numbering_projection,
    existing_numbering_ids,
    write_heading_numbering_projection,
)
from docwen_core.docx_numbering_parts import open_numbering_part
from docwen_core.models.resolved_numbering import (
    HeadingCounterSegment,
    HeadingDefinition,
    HeadingInstance,
    HeadingLevelDefinition,
    HeadingListMaterialization,
    NumberingTarget,
    ResolvedNumberingPlan,
)

pytestmark = pytest.mark.unit

_RELS = "word/_rels/document.xml.rels"
_TYPES = "[Content_Types].xml"
_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"


def _parts(path: Path | BytesIO) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _write(path: Path, parts: dict[str, bytes]) -> None:
    with ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _document_parts() -> dict[str, bytes]:
    memory = BytesIO()
    Document().save(memory)
    return _parts(memory)


def _move(parts: dict[str, bytes], name: str, target: str) -> None:
    parts[name] = parts.pop("word/numbering.xml")
    relationships = etree.fromstring(parts[_RELS])
    next(item for item in relationships if item.get("Type") == _REL).set("Target", target)
    parts[_RELS] = etree.tostring(relationships)
    manifest = etree.fromstring(parts[_TYPES])
    next(item for item in manifest if item.get("PartName") == "/word/numbering.xml").set("PartName", "/" + name)
    parts[_TYPES] = etree.tostring(manifest)


def _projection(document):
    target = NumberingTarget(
        source_start=0,
        source_end=10,
        kind="heading",
        enabled=True,
        target_id=None,
        derived_number="1",
        materialization=HeadingListMaterialization("main", "scope", 1),
    )
    plan = ResolvedNumberingPlan(
        heading_definitions=(
            HeadingDefinition(
                "main",
                (
                    HeadingLevelDefinition(
                        level=1,
                        start=1,
                        number_format="arabic_half",
                        display=(HeadingCounterSegment(1, "arabic_half"),),
                        suffix="space",
                        restart_after_level=None,
                    ),
                ),
            ),
        ),
        heading_instances=(HeadingInstance("scope", "main", ()),),
        targets=(target,),
    )
    abstract_ids, num_ids = existing_numbering_ids(document)
    projection = create_heading_numbering_projection(
        plan,
        heading_style_ids={1: "Heading1"},
        existing_abstract_ids=abstract_ids,
        existing_num_ids=num_ids,
    )
    paragraph = document.add_heading("Authored heading", level=1)
    apply_heading_numbering(paragraph, target, projection, heading_style_ids={1: "Heading1"})
    return projection


@pytest.mark.parametrize(
    ("name", "target"),
    [("word/numbering2.xml", "numbering2.xml"), ("custom/lists.xml", "../custom/lists.xml")],
)
def test_resolved_projection_and_proof_follow_the_same_part(tmp_path: Path, name: str, target: str) -> None:
    path = tmp_path / "moved.docx"
    parts = _document_parts()
    _move(parts, name, target)
    _write(path, parts)
    document = Document(str(path))
    projection = _projection(document)
    document.save(str(path))
    before = _parts(path)
    write_heading_numbering_projection(path, projection)
    after = _parts(path)
    assert "word/numbering.xml" not in after
    assert after[_RELS] == before[_RELS]
    assert after[_TYPES] == before[_TYPES]
    reopened = Document(str(path))
    assert reopened.part.numbering_part.partname == "/" + name
    fact = HeadingNumberingProofIndex.load(path).prove_paragraph(reopened.paragraphs[0])
    assert fact is not None and fact.level == 1 and fact.start == 1
    assert fact.num_id == projection.num_id("scope")


@pytest.mark.parametrize("damage", ["duplicate", "external", "missing", "wrong_type", "escape"])
def test_numbering_support_failures_are_read_only(damage: str) -> None:
    parts = _document_parts()
    relationships = etree.fromstring(parts[_RELS])
    relationship = next(item for item in relationships if item.get("Type") == _REL)
    if damage == "duplicate":
        duplicate = deepcopy(relationship)
        duplicate.set("Id", "rIdDuplicate")
        relationships.append(duplicate)
    elif damage == "external":
        relationship.set("TargetMode", "External")
        relationship.set("Target", "https://example.com/numbering.xml")
    elif damage == "missing":
        relationship.set("Target", "missing.xml")
    elif damage == "escape":
        relationship.set("Target", "../../numbering.xml")
    else:
        manifest = etree.fromstring(parts[_TYPES])
        next(item for item in manifest if item.get("PartName") == "/word/numbering.xml").set(
            "ContentType", "application/xml"
        )
        parts[_TYPES] = etree.tostring(manifest)
    parts[_RELS] = etree.tostring(relationships)
    before = dict(parts)
    with pytest.raises(ValueError):
        open_numbering_part(parts, create=True)
    assert parts == before


def test_creation_does_not_adopt_or_replace_an_unrelated_canonical_part() -> None:
    parts = _document_parts()
    relationships = etree.fromstring(parts[_RELS])
    relationships.remove(next(item for item in relationships if item.get("Type") == _REL))
    parts[_RELS] = etree.tostring(relationships)
    manifest = etree.fromstring(parts[_TYPES])
    next(item for item in manifest if item.get("PartName") == "/word/numbering.xml").set(
        "ContentType", "application/xml"
    )
    parts[_TYPES] = etree.tostring(manifest)
    opaque = b"<unrelated>Keep this opaque member</unrelated>"
    parts["word/numbering.xml"] = opaque
    part = open_numbering_part(parts, create=True)
    assert part is not None and part.name != "word/numbering.xml"
    assert part.element.tag == qn("w:numbering")
    assert parts["word/numbering.xml"] == opaque
    relationships = etree.fromstring(parts[_RELS])
    matching = [item for item in relationships if item.get("Type") == _REL]
    assert len(matching) == 1
    assert "word/" + matching[0].get("Target", "") == part.name


def test_resolved_writer_does_not_replace_package_with_damaged_relationship(tmp_path: Path) -> None:
    document = Document()
    projection = _projection(document)
    path = tmp_path / "broken.docx"
    document.save(str(path))
    parts = _parts(path)
    relationships = etree.fromstring(parts[_RELS])
    next(item for item in relationships if item.get("Type") == _REL).set("Target", "missing.xml")
    parts[_RELS] = etree.tostring(relationships)
    _write(path, parts)
    before = path.read_bytes()
    with pytest.raises(ValueError):
        write_heading_numbering_projection(path, projection)
    assert path.read_bytes() == before
