import pytest
from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docwen_core.docx_parsing.format_features import DocxMarkdownSyntaxConfig
from docwen_plugin_document.shared.markdown_runs import (
    render_paragraph_runs,
    render_paragraph_runs_split_on_page_breaks,
)

pytestmark = pytest.mark.unit


def test_simple_hyperlink_preserves_formatted_result_page_breaks_and_literal_mode():
    paragraph = Document().add_paragraph("Before ")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), r' HYPERLINK "https://example.org/qa" \l "part" ')
    paragraph._p.append(field)
    result = paragraph.add_run("label")
    result.bold = True
    result.add_break(WD_BREAK.PAGE)
    result.add_text("continued")
    field.append(result._r)
    paragraph.add_run(" after")
    syntax = DocxMarkdownSyntaxConfig()
    assert render_paragraph_runs_split_on_page_breaks(paragraph, syntax_config=syntax) == [
        "Before [**label**](https://example.org/qa#part)",
        "[**continued**](https://example.org/qa#part) after",
    ]
    assert render_paragraph_runs(paragraph, literal_text=True, syntax_config=syntax) == "Before label\ncontinued after"


@pytest.mark.parametrize("case", ["unsupported-switch", "nested-simple", "nested-complex", "hidden", "owned-reference"])
def test_simple_hyperlink_respects_field_boundaries_and_visible_results(case):
    paragraph = Document().add_paragraph("Before ")
    field = OxmlElement("w:fldSimple")
    instruction = 'HYPERLINK "https://example.org/qa"'
    field.set(qn("w:instr"), instruction + (r" \p" if case == "unsupported-switch" else ""))
    paragraph._p.append(field)
    visible = paragraph.add_run("visible")
    field.append(visible._r)
    spans = {}
    if case == "nested-simple":
        outer = OxmlElement("w:fldSimple")
        outer.set(qn("w:instr"), "IF ...")
        paragraph._p.append(outer)
        outer.append(field)
    elif case == "nested-complex":
        marker = OxmlElement("w:fldChar")
        marker.set(qn("w:fldCharType"), "begin")
        run = paragraph.add_run()
        run._r.append(marker)
        field.addprevious(run._r)
    elif case == "hidden":
        hidden = paragraph.add_run("secret")
        hidden.font.hidden = True
        field.append(hidden._r)
    elif case == "owned-reference":
        text = visible._r.find(qn("w:t"))
        spans[text] = (text, "@[[#^owned|Alias]]")
    paragraph.add_run(" after")
    expected = "@[[#^owned|Alias]]" if case == "owned-reference" else "visible"
    assert render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig(), reference_spans=spans) == (
        f"Before {expected} after"
    )
