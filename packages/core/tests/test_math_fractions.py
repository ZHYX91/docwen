"""Fraction content and layout survive conversion without acquiring a bar."""

from __future__ import annotations

import pytest
from lxml import etree

from docwen_core.formula import latex_to_mathml, mathml_to_latex, mathml_to_omml, omml_to_mathml
from docwen_core.formula.constants import MATHML_NS, OMML_NS

pytestmark = pytest.mark.unit
_NS = {"m": OMML_NS["m"]}
_VAL = f"{{{OMML_NS['m']}}}val"


def _native(latex: str) -> etree._Element:
    mathml = latex_to_mathml(latex)
    assert mathml
    native = mathml_to_omml(mathml)
    assert native
    return etree.fromstring(native.encode())


def _latex(native: etree._Element) -> str:
    mathml = omml_to_mathml(etree.tostring(native, encoding="unicode"))
    assert mathml
    latex = mathml_to_latex(mathml)
    assert latex
    return latex


def _assert_stacked_fraction(native: etree._Element, numerator: str, denominator: str) -> None:
    [fraction] = native.findall(".//m:f", _NS)
    fraction_type = fraction.find("m:fPr/m:type", _NS)
    assert fraction_type is not None and fraction_type.get(_VAL) == "noBar"
    assert "".join(fraction.xpath("m:num//m:t/text()", namespaces=_NS)) == numerator
    assert "".join(fraction.xpath("m:den//m:t/text()", namespaces=_NS)) == denominator


@pytest.mark.parametrize("latex", [r"\binom{n+1}{k-1}", r"{n+1 \choose k-1}"])
def test_binomial_keeps_delimiters_and_operands_across_two_round_trips(latex: str) -> None:
    for _ in range(3):
        native = _native(latex)
        _assert_stacked_fraction(native, "n+1", "k−1")
        delimiter = native.find("m:d", _NS)
        assert delimiter is not None and len(native) == 1
        assert delimiter.find("m:dPr/m:begChr", _NS).get(_VAL) == "("
        assert delimiter.find("m:dPr/m:endChr", _NS).get(_VAL) == ")"
        latex = _latex(native)
        assert r"\frac" not in latex


def test_independent_native_nobar_fraction_does_not_gain_delimiters_or_division() -> None:
    native = etree.fromstring(
        f'<m:oMath xmlns:m="{OMML_NS["m"]}"><m:f><m:fPr><m:type m:val="noBar"/></m:fPr>'
        "<m:num><m:r><m:t>n</m:t></m:r></m:num>"
        "<m:den><m:r><m:t>k</m:t></m:r></m:den></m:f></m:oMath>"
    )
    for _ in range(2):
        latex = _latex(native)
        assert latex == r"{n \atop k}"
        native = _native(latex)
        _assert_stacked_fraction(native, "n", "k")
        assert native.find(".//m:d", _NS) is None


def test_nested_stacked_fractions_do_not_swap_or_merge_arguments() -> None:
    latex = r"{{a \atop b} \atop {c \atop d}}"
    for _ in range(3):
        native = _native(latex)
        outer = native.find("m:f", _NS)
        assert outer is not None
        assert len(native.findall(".//m:f", _NS)) == 3
        assert [node.get(_VAL) for node in native.findall(".//m:fPr/m:type", _NS)] == ["noBar"] * 3
        assert outer.xpath("m:num/m:f/m:num//m:t/text()", namespaces=_NS) == ["a"]
        assert outer.xpath("m:num/m:f/m:den//m:t/text()", namespaces=_NS) == ["b"]
        assert outer.xpath("m:den/m:f/m:num//m:t/text()", namespaces=_NS) == ["c"]
        assert outer.xpath("m:den/m:f/m:den//m:t/text()", namespaces=_NS) == ["d"]
        latex = _latex(native)


@pytest.mark.parametrize("thickness", ["0", "0pt", ".0em", "0.0%"])
def test_mathml_zero_rule_is_supported_by_both_output_directions(thickness: str) -> None:
    source = f'<math xmlns="{MATHML_NS}"><mfrac linethickness="{thickness}"><mi>n</mi><mi>k</mi></mfrac></math>'
    assert mathml_to_latex(source) == r"{n \atop k}"
    native = mathml_to_omml(source)
    assert native
    _assert_stacked_fraction(etree.fromstring(native.encode()), "n", "k")


@pytest.mark.parametrize("native_type", [None, "bar"])
def test_ordinary_native_fraction_still_has_division_semantics(native_type: str | None) -> None:
    properties = f'<m:fPr><m:type m:val="{native_type}"/></m:fPr>' if native_type else ""
    source = (
        f'<m:oMath xmlns:m="{OMML_NS["m"]}"><m:f>{properties}'
        "<m:num><m:r><m:t>1</m:t></m:r></m:num>"
        "<m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>"
    )
    assert _latex(etree.fromstring(source)) == r"\frac{1}{2}"


@pytest.mark.parametrize(
    "attributes",
    ['bevelled="true"', 'numalign="left"', 'denomalign="right"', 'linethickness="2pt"', 'linethickness="thick"'],
)
@pytest.mark.parametrize("converter", [mathml_to_omml, mathml_to_latex])
def test_unmodeled_mathml_fraction_properties_report_failure(converter, attributes: str, caplog) -> None:
    source = f'<math xmlns="{MATHML_NS}"><mfrac {attributes}><mi>a</mi><mi>b</mi></mfrac></math>'
    assert converter(source) is None
    assert any("unsupported fraction property" in message for message in caplog.messages)


@pytest.mark.parametrize("native_type", ["skw", "lin", "unrecognized"])
def test_unmodeled_native_fraction_types_report_failure(native_type: str, caplog) -> None:
    source = (
        f'<m:oMath xmlns:m="{OMML_NS["m"]}"><m:f><m:fPr><m:type m:val="{native_type}"/></m:fPr>'
        "<m:num><m:r><m:t>a</m:t></m:r></m:num>"
        "<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>"
    )
    assert omml_to_mathml(source) is None
    assert any("unsupported native fraction type" in message for message in caplog.messages)
