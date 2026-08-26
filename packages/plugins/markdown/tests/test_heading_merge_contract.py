"""Regression guards for the classic Markdown heading/body merge contract."""

from __future__ import annotations

import pytest

from docwen_core.text.heading_merge import (
    DEFAULT_HEADING_MERGE_PUNCTUATION,
    normalize_heading_merge_punctuation,
)
from docwen_plugin_markdown.preprocessor import detect_heading_merges

pytestmark = pytest.mark.unit


def test_punctuation_mode_requires_an_immediately_adjacent_body_line() -> None:
    assert detect_heading_merges("## 言之有谋，强化顶层设计：\n坚持全局眼光。") == {0}
    assert detect_heading_merges("## 言之有谋，强化顶层设计：\n\n坚持全局眼光。") == set()


def test_always_mode_removes_only_the_punctuation_requirement() -> None:
    assert detect_heading_merges("## 工作要求\n坚持全局眼光。", mode="always") == {0}
    assert detect_heading_merges("## 工作要求\n- 坚持全局眼光。", mode="always") == set()


def test_never_mode_disables_even_a_punctuation_ending_adjacent_pair() -> None:
    assert detect_heading_merges("## 工作要求：\n坚持全局眼光。", mode="never") == set()


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
    assert detect_heading_merges(f"## 标题：\n{special_line}", mode="always") == set()


def test_custom_punctuation_and_invalid_atx_text_use_exact_heading_indices() -> None:
    source = "#not-a-heading\nordinary\n## Custom§\nbody"
    assert detect_heading_merges(source, punctuation=frozenset("§")) == {0}
    assert detect_heading_merges(source, punctuation=frozenset("：")) == set()


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
    assert detect_heading_merges(f"## 工作要求{ending}\n坚持全局眼光。") == set()
    assert detect_heading_merges(
        f"## 工作要求{ending}\n坚持全局眼光。",
        punctuation=frozenset({ending}),
    ) == {0}
