"""Source ownership follows the Structural Tables switch in both analyzers."""

import pytest

from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_plugin_markdown import document_semantics_v3, number_suite_direct_semantics

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("direct", [False, True])
@pytest.mark.parametrize("enabled", [False, True])
@pytest.mark.parametrize("quoted", [False, True])
def test_structural_table_anchor_ownership_is_dialect_scoped(direct: bool, enabled: bool, quoted: bool) -> None:
    source = "| A | B |\n| C | D |\n| --- || --- |\n| E | F |\n\n^table-id\n"
    if quoted:
        source = "\n".join("> " + line for line in source.splitlines()) + "\n"
    extensions = MarkdownExtensions(structural_tables=enabled)
    analysis = (
        number_suite_direct_semantics.analyze_markdown_semantics_v3(
            source, input_id="table.md", extensions=extensions, consumer_profile="number_suite_direct"
        )
        if direct
        else document_semantics_v3.analyze_markdown_semantics_v3(source, input_id="table.md", extensions=extensions)
    )
    assert analysis.has_errors is not enabled
    if enabled:
        assert [(item["id"], item["block_kind"]) for item in analysis.projection["anchors"]] == [("table-id", "table")]
