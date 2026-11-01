"""Escaped literal citation punctuation follows CommonMark backslash parity."""

import pytest

from docwen_plugin_markdown.mistune_extensions import parse_markdown_text
from docwen_plugin_markdown.number_suite_direct_semantics import analyze_markdown_semantics_v3

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("count", [1, 2, 3, 4])
def test_release_review_citation_backslash_parity(count: int) -> None:
    source = "\\" * count + "@cite-one"
    analysis = analyze_markdown_semantics_v3(source, input_id="literal.md", consumer_profile="number_suite_direct")
    citations = analysis.projection["citations"]
    assert len(citations) == (0 if count % 2 else 1)
    if citations:
        assert citations[0]["raw"] == "@cite-one"
        assert citations[0]["range"] == {"start": count, "end": len(source)}
    tokens = parse_markdown_text(source)
    assert tokens[0]["children"][0]["type"] == "text"
