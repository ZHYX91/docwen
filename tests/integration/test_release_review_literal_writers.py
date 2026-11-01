"""Actual enabled Markdown dialect consumes literal writer output as prose."""

import pytest
from docx import Document

from docwen_core.docx_parsing.format_features import DocxMarkdownSyntaxConfig
from docwen_plugin_document.shared.markdown_runs import render_paragraph_runs
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.number_suite_direct_semantics import analyze_markdown_semantics_v3
from docwen_plugin_optimizer_invoice_cn.invoice_cn.markdown_writer import _render_markdown_table

pytestmark = pytest.mark.integration


@pytest.mark.parametrize("writer", ["docx", "invoice"])
@pytest.mark.parametrize(
    "source", ["$x^2$", "Price $5 + $6", "%%\nhidden literal\n%%", "@cite-one", "Figure: ordinary prose"]
)
def test_release_review_literal_output_does_not_activate_docwen_dialect(writer: str, source: str) -> None:
    if writer == "docx":
        paragraph = Document().add_paragraph(source)
        markdown = render_paragraph_runs(paragraph, syntax_config=DocxMarkdownSyntaxConfig())
    else:
        source = source.replace("\n", " ")
        markdown = _render_markdown_table(headers=["Name"], rows=[{"Name": source}])
    tokens = parse_markdown_text(markdown)

    def flatten(nodes):
        for node in nodes:
            yield node
            yield from flatten(node.get("children", []))

    all_nodes = list(flatten(tokens))
    assert not {"inline_math", "block_math", "comment"}.intersection(node.get("type") for node in all_nodes)
    visible = "".join(node.get("raw", "") for node in all_nodes if node.get("type") == "text")
    assert source.replace("\n", "") in visible.replace("\n", "")
    semantic = analyze_markdown_semantics_v3(markdown, input_id="literal.md", consumer_profile="number_suite_direct")
    assert not semantic.projection["citations"]
    assert not semantic.projection["targets"]
