"""Regression guards for the classic Markdown heading/body merge contract."""

from __future__ import annotations

import pytest

from docwen_core.markdown_extensions import MarkdownExtensions
from docwen_core.text.heading_merge import (
    DEFAULT_HEADING_MERGE_PUNCTUATION,
    normalize_heading_merge_punctuation,
)
from docwen_plugin_markdown.ast_transforms import annotate_ast_with_merges
from docwen_plugin_markdown.mistune_extensions import parse_markdown_text

pytestmark = pytest.mark.unit


def _merged_headings(source, *, mode="punct_required", punctuation=None, extensions=None):
    ast = parse_markdown_text(source, extensions=extensions)
    annotate_ast_with_merges(ast, mode=mode, punctuation=punctuation)
    headings = [node for node in ast if node.get("type") == "heading"]
    for index, node in enumerate(ast):
        if node.get("_merge"):
            assert ast[index + 1]["type"] == "paragraph"
            assert ast[index + 1]["_merged_into_heading"] is True
    return {index for index, node in enumerate(headings) if node.get("_merge")}


def test_punctuation_mode_requires_an_immediately_adjacent_body_line() -> None:
    assert _merged_headings("## 言之有谋，强化顶层设计：\n坚持全局眼光。") == {0}
    assert _merged_headings("## 言之有谋，强化顶层设计：\n\n坚持全局眼光。") == set()


def test_always_mode_removes_only_the_punctuation_requirement() -> None:
    assert _merged_headings("## 工作要求\n坚持全局眼光。", mode="always") == {0}
    assert _merged_headings("## 工作要求\n- 坚持全局眼光。", mode="always") == set()


def test_never_mode_disables_even_a_punctuation_ending_adjacent_pair() -> None:
    assert _merged_headings("## 工作要求：\n坚持全局眼光。", mode="never") == set()


@pytest.mark.parametrize(
    "special_line",
    [
        "### 后续标题",
        "$$E=mc^2$$",
        "| 表头 |",
        "> 引用",
        "```python",
        "~~~text",
        "- 无序列表",
        "1. 有序列表",
        "---",
        "    indented code",
    ],
)
def test_markdown_block_constructs_never_merge_into_a_heading(special_line: str) -> None:
    assert _merged_headings(f"## 标题：\n{special_line}", mode="always") == set()


def test_custom_punctuation_and_invalid_atx_text_use_exact_heading_indices() -> None:
    source = "#not-a-heading\nordinary\n## Custom§\nbody"
    assert _merged_headings(source, punctuation=frozenset("§")) == {0}
    assert _merged_headings(source, punctuation=frozenset("：")) == set()


def test_configurable_punctuation_normalization_uses_strong_semantic_default() -> None:
    normalized = normalize_heading_merge_punctuation(None)
    assert normalized == frozenset(DEFAULT_HEADING_MERGE_PUNCTUATION)
    assert DEFAULT_HEADING_MERGE_PUNCTUATION == "。：！？.:!?"
    assert normalized == frozenset({"。", "：", "！", "？", ".", ":", "!", "?"})
    assert not normalized.intersection({"，", "；", "、", "—", "-", "～", "…", ",", ";"})
    assert normalize_heading_merge_punctuation(" ：： § \n") == frozenset({"：", "§"})
    assert normalize_heading_merge_punctuation("") == frozenset()


@pytest.mark.parametrize("ending", ["，", "；", "、", "—", "-", "～", "…", ",", ";"])
def test_weak_punctuation_does_not_merge_by_default(ending: str) -> None:
    assert _merged_headings(f"## 工作要求{ending}\n坚持全局眼光。") == set()
    assert _merged_headings(
        f"## 工作要求{ending}\n坚持全局眼光。",
        punctuation=frozenset({ending}),
    ) == {0}


@pytest.mark.parametrize("extended", [False, True])
def test_disabled_extended_heading_cannot_shift_a_later_merge(extended: bool) -> None:
    source = "####### Seven\nplain\n\n## Real：\nbody"
    assert _merged_headings(source, extensions=MarkdownExtensions(extended_headings=extended)) == (
        {1} if extended else {0}
    )


def test_nested_headings_and_fenced_examples_cannot_shift_a_top_level_merge() -> None:
    source = "> # Quoted\n> body\n\n```md\n# Sample\n```\n## Real：\nbody"
    assert _merged_headings(source) == {0}


@pytest.mark.parametrize("heading", ["First\nsecond：\n---", "First **bold：**", "First [link：](https://example.com)"])
def test_actual_heading_content_controls_merge(heading: str) -> None:
    source = heading if "\n" in heading else "## " + heading
    assert _merged_headings(source + "\nbody") == {0}


@pytest.mark.parametrize("tail", ["$x$", "![alt](image.png)", "![alt：](image.png)", "<span>"])
def test_opaque_atom_after_punctuation_is_not_dropped_for_merge(tail: str) -> None:
    assert _merged_headings("## Title： " + tail + "\nbody") == set()
