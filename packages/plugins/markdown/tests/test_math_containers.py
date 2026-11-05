"""Math boundaries and identities are independent of Markdown containers."""

from __future__ import annotations

import pytest
from docx import Document
from docx.oxml.ns import qn

from docwen_core.links import markdown_heading_ranges
from docwen_core.markdown_math import serialize_formula
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.renderer import MdToDocxRenderer

pytestmark = pytest.mark.contract

_MATRIX = "\\begin{bmatrix}\na & b \\\\\nc & d\n\\end{bmatrix}"


@pytest.mark.parametrize("fenced", [f"$${_MATRIX}$$", f"$$\n{_MATRIX}\n$$"])
@pytest.mark.parametrize("container", ["body", "list", "quote"])
def test_multiline_formula_owns_its_raw_body_before_heading_rules(fenced: str, container: str) -> None:
    source = fenced.replace("\nc & d", "\n=\nc & d")
    if container == "list":
        lines = source.splitlines()
        source = "1. " + lines[0] + "\n" + "\n".join("   " + line for line in lines[1:])
    elif container == "quote":
        source = "\n".join("> " + line for line in source.splitlines())
    assert markdown_heading_ranges(source) == []
    ast = parse_markdown_text(source)

    def math_nodes(nodes):
        for node in nodes:
            if node["type"] == "block_math":
                yield node
            yield from math_nodes(node.get("children", []))

    [formula] = list(math_nodes(ast))
    assert formula["raw"] == _MATRIX.replace("\nc & d", "\n=\nc & d")


@pytest.mark.parametrize("block", [False, True])
@pytest.mark.parametrize("container", ["body", "list", "cell"])
def test_formula_only_paragraph_keeps_its_explicit_identity(block: bool, container: str) -> None:
    source = serialize_formula("x^2", block=block)
    if container == "list":
        source = "1. " + source
    elif container == "cell":
        source = "| Equation |\n| --- |\n| " + source + " |\n"
    document = Document()
    renderer = MdToDocxRenderer(document)
    renderer.render(parse_markdown_text(source))
    assert not renderer.warnings
    paragraph = document.tables[0].cell(1, 0).paragraphs[0] if container == "cell" else document.paragraphs[0]
    [equation] = paragraph._p.xpath(".//m:oMath")
    assert (equation.getparent().tag == qn("m:oMathPara")) is block


def test_inline_and_display_boundaries_in_one_text_container() -> None:
    [paragraph] = parse_markdown_text("Before $$E=mc^2$$ and $x_1$ after")
    formulas = [node for node in paragraph["children"] if node["type"] in {"inline_math", "block_math"}]
    assert [(node["type"], node["raw"]) for node in formulas] == [("block_math", "E=mc^2"), ("inline_math", "x_1")]


def test_list_formulas_share_marker_indents_and_continuations_have_no_marker() -> None:
    document = Document()
    renderer = MdToDocxRenderer(document)
    renderer.render(parse_markdown_text("1. $x$\n2. $$y$$\n\n   $$z$$\n"))
    first, second, continuation = document.paragraphs
    for paragraph in (first, second):
        assert paragraph._p.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is not None
        assert paragraph._p.find(f"{qn('w:pPr')}/{qn('w:ind')}") is None
    assert continuation._p.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is None
    indent = continuation._p.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    assert indent is not None and indent.get(qn("w:firstLine")) == "0"


@pytest.mark.parametrize(
    "source",
    [
        "`$x$` and `$$x$$`",
        "```math\n$$x$$\n```",
        r"\$x\$ and \$\$x\$\$",
        "Prices $5 and $10",
        "$5+$10",
        "$x$2",
        "$ x$",
        "$x $",
    ],
)
def test_code_escaped_dollars_and_amounts_never_become_math(source: str) -> None:
    document = Document()
    renderer = MdToDocxRenderer(document)
    renderer.render(parse_markdown_text(source))
    assert document._element.xpath(".//m:oMath") == []
    assert not renderer.warnings


@pytest.mark.parametrize(
    ("source", "preserved"),
    [
        (r"Before $x\hspace{1em}y$ after", r"$x\hspace{1em}y$"),
        (r"| H |\n| --- |\n| $$\unknown{x}$$ |".replace(r"\n", "\n"), r"$$\unknown{x}$$"),
        (r"1. $$\lim_{x\to0}\sin x$$", r"$$\lim_{x\to0}\sin x$$"),
    ],
)
def test_failed_formula_retains_delimiters_and_reports_its_occurrence(source: str, preserved: str) -> None:
    document = Document()
    renderer = MdToDocxRenderer(document)
    renderer.render(parse_markdown_text(source))
    assert document._element.xpath(".//m:oMath") == []
    [diagnostic] = renderer.warnings
    assert diagnostic[0] == "MD2DOCX-FORMULA-TEXT-FALLBACK"
    assert "/w:document/w:body/" in diagnostic[1]
    assert "preserved source" in diagnostic[1]
    text = "".join(document._element.xpath(".//w:t/text()"))
    assert preserved in text
