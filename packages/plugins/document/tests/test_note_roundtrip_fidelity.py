"""Note identity and reverse-fidelity regressions."""

from __future__ import annotations

import pytest
from lxml import etree

from docwen_core.docx_parsing.format_features import DocxMarkdownSyntaxConfig
from docwen_plugin_document.shared.note_extraction import _extract_note_content, build_note_definitions

pytestmark = pytest.mark.contract

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


def test_note_content_preserves_breaks_and_supported_run_formatting() -> None:
    note = etree.Element(f"{{{WML_NS}}}note")
    paragraph = etree.SubElement(note, f"{{{WML_NS}}}p")

    marker_run = etree.SubElement(paragraph, f"{{{WML_NS}}}r")
    etree.SubElement(marker_run, f"{{{WML_NS}}}footnoteRef")
    separator_run = etree.SubElement(paragraph, f"{{{WML_NS}}}r")
    separator = etree.SubElement(separator_run, f"{{{WML_NS}}}t")
    separator.set(f"{{{XML_NS}}}space", "preserve")
    separator.text = " "

    plain_run = etree.SubElement(paragraph, f"{{{WML_NS}}}r")
    etree.SubElement(plain_run, f"{{{WML_NS}}}t").text = "First line"
    etree.SubElement(plain_run, f"{{{WML_NS}}}br")
    etree.SubElement(plain_run, f"{{{WML_NS}}}t").text = "second line"

    bold_run = etree.SubElement(paragraph, f"{{{WML_NS}}}r")
    properties = etree.SubElement(bold_run, f"{{{WML_NS}}}rPr")
    etree.SubElement(properties, f"{{{WML_NS}}}b")
    etree.SubElement(bold_run, f"{{{WML_NS}}}br")
    etree.SubElement(bold_run, f"{{{WML_NS}}}t").text = "third line"

    extracted = _extract_note_content(
        note,
        WML_NS,
        "footnoteRef",
        syntax_config=DocxMarkdownSyntaxConfig(),
    )

    assert extracted == "First line\nsecond line\n**third line**"
    assert build_note_definitions({1: extracted}, {1: "1"}) == (
        "[^1]: First line\n"
        "    second line\n"
        "    **third line**"
    )
