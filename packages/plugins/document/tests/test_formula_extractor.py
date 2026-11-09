"""Tests for DOCX OMML → Markdown formula extraction adapter."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


class TestFormulaExtractor:
    def test_omml_inline_to_markdown(self):
        from docwen_plugin_document.to_markdown.formula_extractor import omml_to_markdown

        r = omml_to_markdown(
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:r><m:t>x</m:t></m:r></m:oMath>",
            block=False,
        )
        assert r == "$x$"

    def test_omml_block_to_markdown(self):
        from docwen_plugin_document.to_markdown.formula_extractor import omml_to_markdown

        r = omml_to_markdown(
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:r><m:t>x</m:t></m:r></m:oMath>",
            block=True,
        )
        assert r == "$$x$$"

    def test_empty_omml_returns_empty(self):
        from docwen_plugin_document.to_markdown.formula_extractor import omml_to_markdown

        assert omml_to_markdown("", block=True) == ""

    def test_unknown_native_operator_retains_visible_placeholder_and_location(self):
        from docx import Document
        from docx.oxml import parse_xml

        from docwen_plugin_document.to_markdown.converter import DocxToMarkdownConverter

        doc = Document()
        paragraph = doc.add_paragraph()
        native = parse_xml(
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:acc><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc></m:oMath>"
        )
        paragraph._p.append(native)
        converter = DocxToMarkdownConverter()
        output = converter._build_paragraph_text_with_formulas(paragraph._p, paragraph)
        assert output == "[Formula could not be converted]"
        assert "<m:" not in output
        [diagnostic] = converter._formula_diagnostics
        assert "unsupported" in diagnostic.reason
        assert diagnostic.source_location == "word/document.xml:/w:document/w:body/w:p/m:oMath"
