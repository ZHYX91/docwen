"""Tests for business-neutral text splitting helpers."""

from __future__ import annotations

import pytest

from docwen_core.text.splitting import split_once

pytestmark = pytest.mark.unit


def test_split_once_uses_explicit_boundary_without_interpreting_punctuation() -> None:
    text = "小标题：第一句。第二句；第三句。"

    assert split_once(text, len("小标题：")) == (
        "小标题：",
        "第一句。第二句；第三句。",
    )


@pytest.mark.parametrize("boundary", [-1, 4])
def test_split_once_rejects_out_of_range_boundary(boundary: int) -> None:
    with pytest.raises(ValueError, match="boundary"):
        split_once("abc", boundary)


def test_split_once_accepts_text_edges() -> None:
    assert split_once("abc", 0) == ("", "abc")
    assert split_once("abc", 3) == ("abc", "")
