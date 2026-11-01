"""Invoice independent Markdown writer and public source-name regressions."""

from pathlib import Path

import pytest
import yaml
from markdown_it import MarkdownIt

from docwen_core.models.document_node import ConversionIdentity
from docwen_core.models.file_ref import SOURCE_PRESENTATION_NAME_METADATA_KEY
from docwen_plugin_optimizer_invoice_cn.invoice_cn.converter import InvoiceCnConverter
from docwen_plugin_optimizer_invoice_cn.invoice_cn.markdown_writer import _render_markdown_table

from ._invoice_conversions_support import _build_fake_context

pytestmark = pytest.mark.contract


def test_release_review_invoice_table_retains_literal_visible_text() -> None:
    source = r"*服务* _fee_ [name](x) &amp; <b>raw</b> `code` \ |"
    markdown = _render_markdown_table(headers=["Name"], rows=[{"Name": source + "\nnext"}])
    tokens = MarkdownIt().enable("table").parse(markdown)
    cells = [token for token in tokens if token.type == "inline"]
    assert len(cells) == 2
    children = cells[1].children or []
    assert {token.type for token in children} == {"text"}
    assert "".join(token.content for token in children) == source + " next"


@pytest.mark.parametrize(
    "public_name", ["Invoice [计划] #1.ofd", "logical-source.ofd", "Invoice" + "长" * 120 + ".ofd"]
)
def test_release_review_invoice_public_source_title(
    tmp_path: Path, sample_invoice_ofd_path: Path, public_name: str
) -> None:
    staged = tmp_path / "input-0000.ofd"
    staged.write_bytes(sample_invoice_ofd_path.read_bytes())
    context = _build_fake_context(str(staged), str(tmp_path), "md", action_name="invoice_cn", source_format="ofd")
    ref = context.request.input_refs[0]
    if public_name.startswith("Invoice"):
        ref.metadata[SOURCE_PRESENTATION_NAME_METADATA_KEY] = public_name
        ref.logical_path = "wrong/logical.ofd"
    else:
        ref.logical_path = "declared/" + public_name
    context.request.conversion_identity = ConversionIdentity.create(
        task_id="review", source_stem="truncated-output", source_format="ofd", source_name=public_name
    )
    result = InvoiceCnConverter().convert(context)
    assert result.success, result.error
    text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    metadata = yaml.safe_load(text.split("---", 2)[1])
    assert metadata["标题"] == Path(public_name).stem
    assert metadata["aliases"] == [Path(public_name).stem]
    assert "input-0000" not in text


@pytest.mark.parametrize(
    "source", ["$x^2$", "Price $5 + $6", "%%\nhidden literal\n%%", "@cite-one", "Figure: ordinary prose"]
)
def test_release_review_literal_output_does_not_activate_docwen_dialect(source: str) -> None:
    from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
    from docwen_plugin_markdown.number_suite_direct_semantics import analyze_markdown_semantics_v3

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
