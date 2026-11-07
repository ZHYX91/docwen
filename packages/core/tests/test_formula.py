"""Tests for docwen_core.formula conversion engines."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


class TestLatexToMathml:
    """Tests for latex_to_mathml."""

    def test_empty_returns_none(self) -> None:
        from docwen_core.formula import latex_to_mathml

        assert latex_to_mathml("") is None
        assert latex_to_mathml("   ") is None

    def test_simple_formula_converts(self) -> None:
        from docwen_core.formula import LATEX2MATHML_AVAILABLE, latex_to_mathml

        if not LATEX2MATHML_AVAILABLE:
            pytest.skip("latex2mathml unavailable")
        result = latex_to_mathml(r"E=mc^2")
        assert result is not None
        assert "<math" in result


class TestMathmlToLatex:
    """Tests for mathml_to_latex."""

    def test_identifier(self) -> None:
        from docwen_core.formula import mathml_to_latex

        r = mathml_to_latex('<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>')
        assert r == "x"

    def test_fraction(self) -> None:
        from docwen_core.formula import mathml_to_latex

        r = mathml_to_latex(
            '<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>'
        )
        assert r == r"\frac{a}{b}"

    def test_superscript(self) -> None:
        from docwen_core.formula import mathml_to_latex

        r = mathml_to_latex('<math xmlns="http://www.w3.org/1998/Math/MathML"><msup><mi>x</mi><mn>2</mn></msup></math>')
        assert r == "x^{2}"

    def test_subscript(self) -> None:
        from docwen_core.formula import mathml_to_latex

        r = mathml_to_latex('<math xmlns="http://www.w3.org/1998/Math/MathML"><msub><mi>a</mi><mi>i</mi></msub></math>')
        assert r == "a_{i}"

    def test_sqrt(self) -> None:
        from docwen_core.formula import mathml_to_latex

        r = mathml_to_latex('<math xmlns="http://www.w3.org/1998/Math/MathML"><msqrt><mi>x</mi></msqrt></math>')
        assert r == r"\sqrt{x}"


class TestMathmlToOmml:
    """Tests for mathml_to_omml."""

    def test_identifier_generates_omath(self) -> None:
        from docwen_core.formula import mathml_to_omml

        r = mathml_to_omml('<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>')
        assert r is not None
        assert "oMath" in r

    def test_fraction_generates_f_node(self) -> None:
        from docwen_core.formula import mathml_to_omml

        r = mathml_to_omml(
            '<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>'
        )
        assert r is not None
        assert "f" in r


class TestOmmlToMathml:
    """Tests for omml_to_mathml."""

    def test_identifier_generates_mathml(self) -> None:
        from docwen_core.formula import omml_to_mathml

        r = omml_to_mathml(
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
            "<m:r><m:t>x</m:t></m:r></m:oMath>"
        )
        assert r is not None
        assert "<math" in r
