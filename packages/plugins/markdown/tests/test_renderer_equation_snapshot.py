"""Equation snapshot cleanup must preserve authored mathematical whitespace."""

import pytest
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docwen_plugin_markdown import renderer as markdown_renderer

pytestmark = pytest.mark.contract


def test_equation_snapshot_cleanup_preserves_whitespace_only_leaf_text() -> None:
    equation = OxmlElement("m:oMath")
    equation.text = "\n  "
    run = OxmlElement("m:r")
    run.text = "\n    "
    leaf = OxmlElement("m:t")
    leaf.set(qn("xml:space"), "preserve")
    leaf.text = " "
    leaf.tail = "\n  "
    run.append(leaf)
    equation.append(run)

    markdown_renderer._strip_serialization_only_whitespace(equation)  # pyright: ignore[reportPrivateUsage]

    assert equation.text is None
    assert run.text is None
    assert leaf.text == " "
    assert leaf.tail is None
