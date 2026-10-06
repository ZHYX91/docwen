"""Direct Number Suite source-consumer contracts."""

from __future__ import annotations

import pytest

from docwen_plugin_markdown.number_suite_direct_semantics import analyze_markdown_semantics_v3

pytestmark = pytest.mark.contract


def _analyze(source: str):
    return analyze_markdown_semantics_v3(
        source,
        input_id="fixture.md",
        consumer_profile="number_suite_direct",
    )


def test_direct_number_suite_profile_keeps_standalone_captions_and_normalizes_targets() -> None:
    source = """Figure: Standalone   Caption ^Figure-ID


ordinary body

Figure: Bound caption

![bound](bound.png)

figure: lowercase near miss ^raw-id

Figure:no-space ^raw-no-space

See @[[#^figure-id|  Figure alias  ]] and @[[#figure: standalone caption]].
"""
    analysis = _analyze(source)

    assert not analysis.has_errors
    assert analysis.projection["schema"] == "docwen.number_suite_direct.v1"
    captions = [item for item in analysis.projection["targets"] if item["kind"] == "figure"]
    assert [(item["title"], item["number"]) for item in captions] == [
        ("Standalone   Caption", "1"),
        ("Bound caption", "2"),
    ]
    assert "object_range" not in captions[0]
    assert source[captions[1]["object_range"]["start"] : captions[1]["object_range"]["end"]].startswith("![bound]")
    assert [(item["id"], item["block_kind"]) for item in analysis.projection["anchors"]] == [
        ("raw-id", "paragraph"),
        ("raw-no-space", "paragraph"),
    ]
    references = analysis.projection["references"]
    assert [(item["resolution_status"], item["cached_number"]) for item in references] == [
        ("resolved", "1"),
        ("resolved", "1"),
    ]
    assert references[0]["target_id"] == "Figure-ID"
    assert references[0]["alias"] == "Figure alias"
    assert references[1]["resolved_target_id"] == "Figure-ID"


def test_direct_number_suite_profile_does_not_promote_nested_caption_lines() -> None:
    source = """> Figure: Quoted ^quoted
>
> body

- Figure: Listed ^listed

Figure: Top ^top
"""

    analysis = _analyze(source)

    assert not analysis.has_errors
    captions = [item for item in analysis.projection["targets"] if item["kind"] == "figure"]
    assert [(item["title"], item.get("id")) for item in captions] == [("Top", "top")]
    assert all(item["title"] not in {"Quoted", "Listed"} for item in captions)


def test_direct_number_suite_profile_does_not_invent_hierarchical_title_references() -> None:
    source = "# Parent\n## Child\n\n@[[#Parent#Child]]\n"

    analysis = _analyze(source)

    assert analysis.has_errors
    [reference] = analysis.projection["references"]
    assert reference["heading_path"] == ["Parent#Child"]
    assert reference["resolution_status"] == "missing"
    assert [item["code"] for item in analysis.diagnostics] == ["docwen.markdown.cross_reference.missing"]


def test_direct_number_suite_profile_treats_case_only_block_ids_as_duplicate() -> None:
    source = "# One ^Same\n\nParagraph ^same\n"

    analysis = _analyze(source)

    assert analysis.has_errors
    assert [item["code"] for item in analysis.diagnostics] == ["docwen.markdown.anchor.duplicate"]


def test_direct_number_suite_reference_scanner_respects_literal_regions() -> None:
    source = """# Target ^target

\\@[[#^missing]]
<!-- @[[#^missing]] @hidden-html -->
%% @[[#^missing]] @hidden-obsidian %%
[Link](https://example.test/@[[#^missing]])
<span data-ref="@[[#^missing]]">@hidden-attribute</span>
`@[[#^missing]] @hidden-code`
https://example.test/path@[[#^target]]
Real @[[#^target]] @real-cite.
"""

    analysis = _analyze(source)

    assert not analysis.has_errors
    references = analysis.projection["references"]
    assert [item["raw"] for item in references] == ["@[[#^target]]", "@[[#^target]]"]
    assert all(item["resolution_status"] == "resolved" for item in references)
    assert [item["raw"] for item in analysis.projection["citations"]] == ["@real-cite"]


def test_direct_number_suite_multiline_comments_hide_reference_like_tokens() -> None:
    source = """# Target ^target

<!--
@[[#^missing]]
@hidden-html
-->
%%
@[[#^missing]]
@hidden-obsidian
%%

@[[#^target]]
"""

    analysis = _analyze(source)

    assert not analysis.has_errors
    assert [item["raw"] for item in analysis.projection["references"]] == ["@[[#^target]]"]
    assert analysis.projection["citations"] == []
