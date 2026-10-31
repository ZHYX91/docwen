import pytest
from docx import Document
from docx.oxml import OxmlElement
from lxml import etree

from docwen_core.docx_parsing.textbox_extraction import extract_textbox_paragraphs

pytestmark = pytest.mark.unit

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_V = "urn:schemas-microsoft-com:vml"


def _box(parent, namespace, tag, texts):
    box = etree.SubElement(parent, f"{{{namespace}}}{tag}")
    content = etree.SubElement(box, f"{{{_W}}}txbxContent")
    for text in texts:
        paragraph = etree.SubElement(content, f"{{{_W}}}p")
        run = etree.SubElement(paragraph, f"{{{_W}}}r")
        etree.SubElement(run, f"{{{_W}}}t").text = text


def test_textbox_ownership_is_xml_occurrence_and_alternate_representation():
    # Equal text at one anchor is authored twice; only Choice/Fallback are
    # alternative serializations of the same shape. Text-key deduplication
    # used to erase the second A and equal text in another real textbox.
    document = Document()
    paragraph = document.add_paragraph("Anchor")
    run = OxmlElement("w:r")
    paragraph._p.append(run)
    alt = etree.SubElement(run, f"{{{_MC}}}AlternateContent")
    choice = etree.SubElement(alt, f"{{{_MC}}}Choice")
    fallback = etree.SubElement(alt, f"{{{_MC}}}Fallback")
    _box(choice, _WPS, "txbx", ["A", "B", "A"])
    _box(fallback, _V, "textbox", ["A", "B", "A"])
    _box(run, _V, "textbox", ["A"])
    extracted = extract_textbox_paragraphs(document)
    assert [(item.text, item.anchor_index, item.source, item.part_name) for item in extracted] == [
        ("A", 0, "textbox", "document"),
        ("B", 0, "textbox", "document"),
        ("A", 0, "textbox", "document"),
        ("A", 0, "textbox", "document"),
    ]
