"""CLI 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

import docwen.cli.utils as cli_utils

pytestmark = pytest.mark.unit


def test_prompt_files_accepts_single_existing_path_with_spaces(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    f = tmp_path / "a b.docx"
    f.write_text("x", encoding="utf-8")

    monkeypatch.setattr("builtins.input", lambda *_args, **_kwargs: str(f), raising=True)

    files = cli_utils.prompt_files("x")
    assert files == [str(f)]


def test_categorize_files_maps_text_to_markdown(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        "docwen.utils.file_type_utils.get_strategy_file_category", lambda *_a, **_k: "markdown", raising=True
    )
    assert cli_utils.categorize_files(["a.bin"]) == {"markdown": ["a.bin"]}


def test_expand_paths_preserves_order_and_deduplicates(tmp_path: Path) -> None:
    b = tmp_path / "b.docx"
    a = tmp_path / "a.docx"
    b.write_text("x", encoding="utf-8")
    a.write_text("x", encoding="utf-8")

    expanded = cli_utils.expand_paths([str(b), str(a), str(b)])
    assert expanded == [str(b), str(a)]
