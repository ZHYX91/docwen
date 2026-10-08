"""Field and bookmark proofs must agree with visible native note payloads."""

from copy import deepcopy
from pathlib import Path

import pytest
from docx import Document
from lxml import etree

from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_plugin_document.shared.note_extraction import NoteExtractor
from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter
from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter
from tests.integration.test_markdown_extension_policy import _context

pytestmark = pytest.mark.integration
Q = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _document(tmp_path, kind):
    label = "n" if kind == "footnote" else "endnote:n"
    source = tmp_path / "source.md"
    source.write_text(f"First[^{label}], again[^{label}].\n\n[^{label}]: Unique body.\n", encoding="utf-8")
    result = MdToDocxConverter().convert(_context(tmp_path, source, "docx", {"input": {"typed_endnotes": True}}))
    assert result.success, result.error
    return Document(result.artifacts[0].staging_path)


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("location", ["rPr", "nested_run"])
def test_fake_native_marker_cannot_authenticate_field(tmp_path, kind, location):
    document = _document(tmp_path, kind)
    reference = next(document.element.iter(Q + kind + "Reference"))
    run = reference.getparent()
    container = run.find(Q + "rPr") if location == "rPr" else etree.SubElement(run, Q + "r")
    assert container is not None
    container.append(reference)
    path = tmp_path / "fake-target.docx"
    document.save(path)
    loaded = Document(path)
    extractor = NoteExtractor(loaded, str(path))
    instruction = next(loaded.element.iter(Q + "instrText")).text
    assert extractor.get_noteref_text(instruction) is None


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("location", ["simple_field", "bookmark"])
@pytest.mark.parametrize("marker_case", ["valid", "unknown_type", "attribute", "text", "child"])
def test_proofing_markers_preserve_only_proven_note_identity(tmp_path, kind, location, marker_case):
    document = _document(tmp_path, kind)
    instruction = next(document.element.iter(Q + "instrText"))
    begin = instruction.getparent().getprevious()
    parts = [begin]
    for _ in range(4):
        parts.append(parts[-1].getnext())
    field = etree.Element(Q + "fldSimple", {Q + "instr": instruction.text})
    field.append(deepcopy(parts[3]))
    begin.addprevious(field)
    for run in parts:
        run.getparent().remove(run)
    marker = etree.Element(Q + "proofErr", {Q + "type": "spellStart"})
    if marker_case == "unknown_type":
        marker.set(Q + "type", "unknown")
    elif marker_case == "attribute":
        marker.set(Q + "extra", "1")
    elif marker_case == "text":
        marker.text = "invalid"
    elif marker_case == "child":
        etree.SubElement(marker, Q + "t").text = "invalid"
    if location == "simple_field":
        field.insert(0, marker)
        field.append(etree.Element(Q + "proofErr", {Q + "type": "spellEnd"}))
    else:
        reference = next(document.element.iter(Q + kind + "Reference"))
        reference.getparent().addprevious(marker)
        reference.getparent().addnext(etree.Element(Q + "proofErr", {Q + "type": "spellEnd"}))
    path = tmp_path / "proofing.docx"
    document.save(path)
    reverse = tmp_path / "reverse"
    reverse.mkdir()
    result = DocxToMarkdownConverter().convert(
        _context(reverse, path, "md", {"output": MarkdownExtensions.obsidian().to_dict()})
    )
    assert result.success, result.error
    text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    label = "[^1]" if kind == "footnote" else "[^endnote:1]"
    # One definition plus one native occurrence; only a proven field adds another.
    assert text.count(label) == (3 if marker_case == "valid" else 2), text
    assert text.count("Unique body.") == 1


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("wrapper_kind", ["sdt", "hyperlink", "smartTag", "customXml", "ins", "moveTo", "fldSimple"])
@pytest.mark.parametrize("malformed", [False, True])
def test_note_target_wrappers_preserve_only_well_formed_payload(tmp_path, kind, wrapper_kind, malformed):
    document = _document(tmp_path, kind)
    reference = next(document.element.iter(Q + kind + "Reference"))
    run = reference.getparent()
    wrapper = etree.Element(Q + wrapper_kind)
    if wrapper_kind in {"ins", "moveTo"}:
        wrapper.set(Q + "id", "7")
        wrapper.set(Q + "author", "Editor")
    elif wrapper_kind in {"smartTag", "customXml"}:
        wrapper.set(Q + "element", "example")
    elif wrapper_kind == "fldSimple":
        wrapper.set(Q + "instr", " QUOTE 1 ")
    run.addprevious(wrapper)
    if wrapper_kind == "sdt":
        properties = etree.SubElement(wrapper, Q + "sdtPr")
        etree.SubElement(properties, Q + "tag", {Q + "val": "Example"})
        content = etree.SubElement(wrapper, Q + "sdtContent")
        content.append(run)
        if malformed:
            # A run belongs in sdtContent, never beside it.
            etree.SubElement(wrapper, Q + "r")
    else:
        wrapper.append(run)
        if malformed:
            etree.SubElement(wrapper, Q + "drawing")
    _assert_full_note_roundtrip(tmp_path, document, kind, valid=not malformed)


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("property_name", ["b", "fldChar", "drawing", "proofErr"])
def test_note_target_run_formatting_is_validated(tmp_path, kind, property_name):
    document = _document(tmp_path, kind)
    reference = next(document.element.iter(Q + kind + "Reference"))
    properties = reference.getparent().find(Q + "rPr")
    assert properties is not None
    etree.SubElement(properties, Q + property_name)
    _assert_full_note_roundtrip(tmp_path, document, kind, valid=property_name == "b")


def _assert_full_note_roundtrip(tmp_path, document, kind, *, valid):
    path = tmp_path / "target-shape.docx"
    document.save(path)
    reverse = tmp_path / "reverse"
    reverse.mkdir()
    result = DocxToMarkdownConverter().convert(
        _context(reverse, path, "md", {"output": MarkdownExtensions.obsidian().to_dict()})
    )
    assert result.success, result.error
    text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    label = "[^1]" if kind == "footnote" else "[^endnote:1]"
    assert text.count(label) == (3 if valid else 2), text
    assert text.count("Unique body.") == 1


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("valid", [True, False])
def test_note_target_preserves_nested_format_revision_metadata(tmp_path, kind, valid):
    document = _document(tmp_path, kind)
    reference = next(document.element.iter(Q + kind + "Reference"))
    properties = reference.getparent().find(Q + "rPr")
    assert properties is not None
    etree.SubElement(properties, Q + "vanish", {Q + "val": " false "})
    change = etree.SubElement(properties, Q + "rPrChange", {Q + "id": "7", Q + "author": "Editor"})
    previous = etree.SubElement(change, Q + "rPr")
    etree.SubElement(previous, Q + ("i" if valid else "fldChar"))
    _assert_full_note_roundtrip(tmp_path, document, kind, valid=valid)


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("valid", [True, False])
def test_note_target_sdt_properties_have_typed_nested_structure(tmp_path, kind, valid):
    document = _document(tmp_path, kind)
    reference = next(document.element.iter(Q + kind + "Reference"))
    run = reference.getparent()
    wrapper = etree.Element(Q + "sdt")
    run.addprevious(wrapper)
    properties = etree.SubElement(wrapper, Q + "sdtPr")
    etree.SubElement(properties, Q + "alias", {Q + "val": "Label"})
    placeholder = etree.SubElement(properties, Q + "placeholder")
    etree.SubElement(placeholder, Q + ("docPart" if valid else "fldChar"), {Q + "val": "Placeholder"})
    etree.SubElement(wrapper, Q + "sdtContent").append(run)
    _assert_full_note_roundtrip(tmp_path, document, kind, valid=valid)


@pytest.mark.parametrize("kind", ["footnote", "endnote"])
@pytest.mark.parametrize("simple", [True, False])
def test_note_field_cache_cannot_hide_payload_in_run_properties(tmp_path, kind, simple):
    document = _document(tmp_path, kind)
    instruction = next(document.element.iter(Q + "instrText"))
    begin = instruction.getparent().getprevious()
    parts = [begin]
    for _ in range(4):
        parts.append(parts[-1].getnext())
    properties = parts[3].find(Q + "rPr")
    assert properties is not None
    etree.SubElement(properties, Q + "fldChar")
    if simple:
        field = etree.Element(Q + "fldSimple", {Q + "instr": instruction.text})
        field.append(deepcopy(parts[3]))
        begin.addprevious(field)
        for run in parts:
            run.getparent().remove(run)
    _assert_full_note_roundtrip(tmp_path, document, kind, valid=False)
