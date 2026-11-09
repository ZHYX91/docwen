"""Independent native-math and operand checks for the shared converters."""

from __future__ import annotations

import pytest
from lxml import etree

from docwen_core.formula import latex_to_mathml, mathml_to_latex, mathml_to_omml, omml_to_mathml
from docwen_core.formula.constants import MATHML_NS, OMML_NS

pytestmark = pytest.mark.unit
_NS = {"m": OMML_NS["m"]}


def _native(latex: str) -> etree._Element:
    mathml = latex_to_mathml(latex)
    assert mathml
    native = mathml_to_omml(mathml)
    assert native
    return etree.fromstring(native.encode())


def _returned(native: etree._Element) -> str:
    mathml = omml_to_mathml(etree.tostring(native, encoding="unicode"))
    assert mathml
    latex = mathml_to_latex(mathml)
    assert latex
    return latex


def _required(element: etree._Element, path: str) -> etree._Element:
    child = element.find(path, _NS)
    assert child is not None, path
    return child


def _text(element: etree._Element, path: str) -> str:
    return "".join(child.text or "" for child in element.findall(path, _NS))


@pytest.mark.parametrize("degree", [None, "3", "n+1"])
def test_radical_degree_visibility_and_compound_operands(degree: str | None) -> None:
    latex = r"\sqrt{a+b}" if degree is None else f"\\sqrt[{degree}]{{a+b}}"
    native = _native(latex)
    radical = native.find("m:rad", _NS)
    assert radical is not None
    assert _required(radical, "m:radPr/m:degHide").get(f"{{{OMML_NS['m']}}}val") == ("1" if degree is None else "0")
    assert _text(radical, "m:deg//m:t") == (degree or "")
    assert _text(radical, "m:e//m:t") == "a+b"
    for _ in range(2):
        returned = _returned(native)
        assert returned.replace(" ", "") == latex
        native = _native(returned)


def test_limit_owns_its_fraction_as_a_function_argument() -> None:
    latex = r"\lim_{h\to0}\frac{f(x+h)-f(x)}{h}"
    for _ in range(2):
        native = _native(latex)
        function = native.find("m:func", _NS)
        assert function is not None
        assert _required(function, "m:fName/m:limLow/m:e/m:r/m:t").text == "lim"
        assert _text(function, "m:fName/m:limLow/m:lim//m:t") == "h→0"
        fraction = function.find("m:e/m:f", _NS)
        assert fraction is not None
        assert _text(fraction, "m:num//m:t") == "fx+h−fx"
        assert _text(fraction, "m:den//m:t") == "h"
        assert len(native) == 1
        latex = _returned(native)


@pytest.mark.parametrize(("latex", "script"), [(r"\sin^2 x", "sSup"), (r"\log_2 x", "sSub")])
def test_named_function_scripts_keep_side_placement(latex: str, script: str) -> None:
    for _ in range(2):
        native = _native(latex)
        assert native.find(f"m:{script}", _NS) is not None
        assert native.find(".//m:func", _NS) is None
        assert native.find(".//m:limLow", _NS) is None
        assert native.find(".//m:limUpp", _NS) is None
        latex = _returned(native)


def test_adjacent_functions_remain_a_product_not_function_composition() -> None:
    latex = r"\sin x \cos y"
    for _ in range(2):
        native = _native(latex)
        assert native.find(".//m:func", _NS) is None
        assert _text(native, "m:r/m:t") == "sinxcosy"
        latex = _returned(native)


@pytest.mark.parametrize(
    ("latex", "operand_text"),
    [(r"\sum_{i=1}^n a_i \sin i", "aisini"), (r"\int x \sin x dx", "xsinxdx")],
)
def test_nary_products_keep_named_function_factors_inside_the_operand(latex: str, operand_text: str) -> None:
    for _ in range(2):
        native = _native(latex)
        assert len(native) == 1
        nary = _required(native, "m:nary")
        assert _text(nary, "m:e//m:t") == operand_text
        assert native.find(".//m:func", _NS) is None
        latex = _returned(native)


def test_parenthesized_function_is_not_given_a_following_factor_as_an_argument() -> None:
    latex = r"\sin (x) y"
    for _ in range(2):
        native = _native(latex)
        assert native.find(".//m:func", _NS) is None
        assert [etree.QName(child).localname for child in native] == ["r", "d", "r"]
        assert _required(native, "m:r/m:t").text == "sin"
        assert _text(native, "m:d//m:t") == "x"
        assert _text(native, "m:r/m:t") == "siny"
        latex = _returned(native)


@pytest.mark.parametrize("latex", [r"\lim_{x\to0}\sin x", r"\lim_{x\to0}\frac{1}{x}y"])
def test_ambiguous_limit_scope_preserves_source_instead_of_inventing_an_argument(latex: str, caplog) -> None:
    mathml = latex_to_mathml(latex)
    assert mathml
    assert mathml_to_omml(mathml) is None
    assert any("ambiguous" in message for message in caplog.messages)


@pytest.mark.parametrize("latex", [r"\vec{x}", r"\hat{x}", r"\overline{x}"])
def test_accents_without_an_explicit_accent_flag_are_never_retyped_as_scripts(latex: str) -> None:
    mathml = latex_to_mathml(latex)
    assert mathml
    assert mathml_to_omml(mathml) is None


def test_explicit_mathml_upright_identifier_is_preserved_and_lossy_tex_parser_is_reported(caplog) -> None:
    source = f'<math xmlns="{MATHML_NS}"><mi mathvariant="normal">d</mi></math>'
    native = mathml_to_omml(source)
    assert native
    returned = omml_to_mathml(native)
    assert returned and 'mathvariant="normal"' in returned
    latex = mathml_to_latex(returned)
    assert latex == r"\mathrm{d}"
    assert latex_to_mathml(latex) is None
    assert any("upright identifier" in message for message in caplog.messages)


@pytest.mark.parametrize("latex", [r"\mathit{lim}_x", r"\text{lim}_x", "l_i"])
def test_text_and_identifiers_are_not_guessed_to_be_limit_operators(latex: str) -> None:
    native = _native(latex)
    assert native.find(".//m:func", _NS) is None
    assert native.find(".//m:limLow", _NS) is None
    assert native.find(".//m:sSub", _NS) is not None


def test_matrix_cells_use_the_same_expression_parser_as_the_body() -> None:
    native = _native(r"\begin{bmatrix}a+b & \frac{1}{x} \\ \sqrt[n+1]{y+z} & c_d\end{bmatrix}")
    matrix = native.find("m:d/m:e/m:m", _NS)
    assert matrix is not None
    rows = matrix.findall("m:mr", _NS)
    assert [len(row.findall("m:e", _NS)) for row in rows] == [2, 2]
    assert rows[0].find("m:e[2]/m:f", _NS) is not None
    assert rows[1].find("m:e/m:rad/m:deg", _NS) is not None
    assert _text(_required(rows[0], "m:e"), ".//m:t") == "a+b"
    assert _returned(_native(_returned(native))) == _returned(native)


def test_compound_group_containing_a_matrix_keeps_all_its_factors() -> None:
    source = (
        f'<math xmlns="{MATHML_NS}"><mfenced open="(" close=")"><mrow>'
        "<mi>x</mi><mo>+</mo><mtable><mtr><mtd><mi>y</mi></mtd></mtr></mtable>"
        "</mrow></mfenced></math>"
    )
    latex = mathml_to_latex(source)
    assert latex == "(x + \\begin{matrix}\ny\n\\end{matrix})"


@pytest.mark.parametrize("latex", [r"\{x\}", r"\$5", r"a\_b", r"a\backslash b"])
def test_visible_special_characters_never_turn_into_tex_control_syntax(latex: str) -> None:
    expected = latex.replace(" ", "{}") if "backslash" in latex else latex
    for _ in range(2):
        native = _native(latex)
        returned = _returned(native)
        assert returned == expected
        latex = returned


@pytest.mark.parametrize("environment", ["matrix", "bmatrix", "pmatrix", "vmatrix", "Vmatrix", "Bmatrix"])
def test_matrix_delimiters_come_from_structure_without_bracket_guessing(environment: str) -> None:
    latex = f"\\begin{{{environment}}}a & b \\\\ c & d\\end{{{environment}}}"
    for _ in range(2):
        returned = _returned(_native(latex))
        assert f"\\begin{{{environment}}}" in returned
        assert "a & b \\\\\nc & d" in returned
        latex = returned


@pytest.mark.parametrize("children", ["<mi>a</mi>", "<mi>a</mi><mi>b</mi><mi>discarded</mi>"])
@pytest.mark.parametrize("converter", [mathml_to_omml, mathml_to_latex])
def test_malformed_fraction_cannot_silently_discard_operands(converter, children: str, caplog) -> None:
    source = f'<math xmlns="{MATHML_NS}"><mi>x</mi><mfrac>{children}</mfrac></math>'
    assert converter(source) is None
    assert any("requires 2 children" in message for message in caplog.messages)


@pytest.mark.parametrize("converter", [mathml_to_omml, mathml_to_latex])
def test_unknown_container_cannot_silently_lose_its_operation(converter, caplog) -> None:
    source = f'<math xmlns="{MATHML_NS}"><menclose notation="box"><mi>x</mi></menclose></math>'
    assert converter(source) is None
    assert any("menclose" in message for message in caplog.messages)


def test_unknown_native_structure_refuses_partial_extraction(caplog) -> None:
    source = (
        f'<m:oMath xmlns:m="{OMML_NS["m"]}"><m:r><m:t>x</m:t></m:r>'
        "<m:acc><m:e><m:r><m:t>y</m:t></m:r></m:e></m:acc></m:oMath>"
    )
    assert omml_to_mathml(source) is None
    assert any("acc" in message for message in caplog.messages)
