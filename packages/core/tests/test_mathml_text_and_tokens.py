"""Independent leaf and TeX-token expectations for the shared math engines."""

from __future__ import annotations

import re
from xml.etree import ElementTree as ET

import pytest

from docwen_core.formula import mathml_to_latex, mathml_to_omml, omml_to_mathml

pytestmark = pytest.mark.unit

_MATHML = "http://www.w3.org/1998/Math/MathML"
_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


@pytest.mark.parametrize("text", [" kg", " if ", "−5 kg", " x_1% "])
def test_mathml_text_leaf_keeps_literal_characters_and_spaces(text: str) -> None:
    root = ET.Element(f"{{{_MATHML}}}math")
    ET.SubElement(root, f"{{{_MATHML}}}mtext").text = text
    source = ET.tostring(root, encoding="unicode")
    result = mathml_to_omml(source)
    assert result is not None
    native = ET.fromstring(result)
    assert [n.text for n in native.iter(f"{{{_M}}}t")] == [text]
    assert native.find(f"{{{_M}}}r/{{{_M}}}rPr/{{{_M}}}nor") is not None
    native_text = native.find(f"{{{_M}}}r/{{{_M}}}t")
    assert native_text is not None
    assert native_text.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    expected = text.replace("\\", r"\textbackslash{}").replace("_", r"\_").replace("%", r"\%")
    assert mathml_to_latex(source) == "\\text{" + expected + "}"
    # Native normal-text runs must not lose the whitespace on import either.
    returned = omml_to_mathml(result)
    assert returned is not None
    assert [n.text for n in ET.fromstring(returned).iter() if n.tag.split("}")[-1] == "mtext"] == [text]


@pytest.mark.parametrize(
    ("children", "expected"),
    [
        ("<mi>α</mi><mi>x</mi>", r"\alpha{}x"),
        ("<mrow><mi>α</mi></mrow><mi>x</mi>", r"\alpha{}x"),
        ("<mi>α</mi><mrow><mi>x</mi></mrow>", r"\alpha{}x"),
        ("<mi>lim</mi><mi>x</mi>", r"\lim{}x"),
    ],
)
def test_control_word_boundary_is_inserted_only_at_sibling_join(children: str, expected: str) -> None:
    actual = mathml_to_latex(f'<math xmlns="{_MATHML}">{children}</math>')
    assert actual is not None
    assert actual == expected
    assert re.findall(r"\\[A-Za-z]+|[A-Za-z]", actual) == re.findall(r"\\[A-Za-z]+|[A-Za-z]", expected)


@pytest.mark.parametrize("converter", [mathml_to_omml, mathml_to_latex])
def test_unknown_text_leaf_refuses_partial_formula_with_a_diagnostic(
    converter, caplog: pytest.LogCaptureFixture
) -> None:
    source = f'<math xmlns="{_MATHML}"><mi>x</mi><unsupported>unit</unsupported></math>'
    assert converter(source) is None
    assert any("unsupported" in message.lower() for message in caplog.messages)
