import pytest
from lxml import etree

from docwen_plugin_document.shared.note_extraction import (
    NoteExtractor,
    _extract_note_content,
    build_note_definitions,
)

pytestmark = pytest.mark.unit

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


@pytest.mark.parametrize("ref_tag", ["footnoteRef", "endnoteRef"])
def test_note_content_removes_only_structural_reference_separator(ref_tag: str):
    note = etree.Element(f"{{{WML_NS}}}note")
    first = etree.SubElement(note, f"{{{WML_NS}}}p")
    marker_run = etree.SubElement(first, f"{{{WML_NS}}}r")
    etree.SubElement(marker_run, f"{{{WML_NS}}}{ref_tag}")
    separator_run = etree.SubElement(first, f"{{{WML_NS}}}r")
    separator = etree.SubElement(separator_run, f"{{{WML_NS}}}t")
    separator.set(f"{{{XML_NS}}}space", "preserve")
    separator.text = " "
    authored_run = etree.SubElement(first, f"{{{WML_NS}}}r")
    authored = etree.SubElement(authored_run, f"{{{WML_NS}}}t")
    authored.set(f"{{{XML_NS}}}space", "preserve")
    authored.text = "  Authored leading whitespace"

    second = etree.SubElement(note, f"{{{WML_NS}}}p")
    continuation_run = etree.SubElement(second, f"{{{WML_NS}}}r")
    continuation = etree.SubElement(continuation_run, f"{{{WML_NS}}}t")
    continuation.set(f"{{{XML_NS}}}space", "preserve")
    continuation.text = " Continuation leading whitespace"

    assert _extract_note_content(note, WML_NS, ref_tag) == (
        "  Authored leading whitespace\n Continuation leading whitespace"
    )


def test_build_note_definitions_formats_multiline_content():
    notes = {5: "第一行\n第二行"}
    assert build_note_definitions(notes, {5: "1"}) == "[^1]: 第一行\n    第二行"


def test_endnote_definitions_use_prefix():
    notes = {9: "尾注"}
    assert build_note_definitions(notes, {9: "endnote:1"}) == "[^endnote:1]: 尾注"


def test_note_extractor_footnotes_default_empty():
    extractor = NoteExtractor.__new__(NoteExtractor)
    extractor.footnotes = {}
    extractor.endnotes = {}
    extractor.footnote_id_map = {1: "1"}
    extractor.endnote_id_map = {1: "endnote:1"}

    assert extractor.get_reference_text("footnote", 1) == "[^1]"
    assert extractor.get_reference_text("endnote", 1) == "[^endnote:1]"
    assert extractor.build_definitions_block() == ""


def test_note_extractor_builds_definitions_block():
    extractor = NoteExtractor.__new__(NoteExtractor)
    extractor.footnotes = {7: "脚注内容"}
    extractor.endnotes = {8: "尾注内容"}
    extractor.footnote_id_map = {7: "1"}
    extractor.endnote_id_map = {8: "endnote:1"}

    block = extractor.build_definitions_block()
    assert "[^1]: 脚注内容" in block
    assert "[^endnote:1]: 尾注内容" in block


def test_note_extractor_numbers_each_domain_by_first_reference():
    extractor = NoteExtractor.__new__(NoteExtractor)
    extractor.footnotes = {40: "later Word ID", 3: "earlier Word ID"}
    extractor.endnotes = {90: "later Word ID", 2: "earlier Word ID"}
    extractor.footnote_id_map = {}
    extractor.endnote_id_map = {}

    assert extractor.get_reference_text("footnote", 40) == "[^1]"
    assert extractor.get_reference_text("footnote", 3) == "[^2]"
    assert extractor.get_reference_text("footnote", 40) == "[^1]"
    assert extractor.get_reference_text("endnote", 90) == "[^endnote:1]"
    assert extractor.get_reference_text("endnote", 2) == "[^endnote:2]"

    block = extractor.build_definitions_block()
    assert block.index("[^1]: later Word ID") < block.index("[^2]: earlier Word ID")
    assert block.index("[^endnote:1]: later Word ID") < block.index("[^endnote:2]: earlier Word ID")
