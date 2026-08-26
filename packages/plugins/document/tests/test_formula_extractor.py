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
