"""md2docx formula_handler 的单元测试。"""

from __future__ import annotations

import pytest

from docwen.converter.md2docx.handlers import formula_handler as fh

pytestmark = pytest.mark.unit


def test_formula_handler_placeholders_roundtrip() -> None:
    text = "a $E=mc^2$ b\n\n$$\nx\n$$"
    assert fh.has_latex_formulas(text) is True

    replaced, formulas = fh.replace_formulas_with_placeholders(text)
    assert "{{FORMULA_0}}" in replaced
    restored = fh.restore_formulas_from_placeholders(replaced, formulas)
    assert "E=mc^2" in restored


def test_formula_handler_convert_block_formulas_to_paragraphs() -> None:
    lines = ["x", "$$", "a", "$$", "y"]
    out = fh.convert_block_formulas_to_paragraphs(None, lines)
    assert out == ["x", "$$\na\n$$", "y"]
