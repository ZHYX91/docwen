"""config 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.config.toml_operations import (
    extract_inline_comments,
    get_toml_value,
    read_toml_file,
    save_mapping_with_comments,
    update_toml_value,
    validate_toml_syntax,
    write_toml_file,
)

pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_read_write_update_get(tmp_path: Path) -> None:
    toml_path = tmp_path / "config.toml"
    data = {
        "gui_config": {"window": {"width": 800, "height": 600}, "theme": {"default_theme": "dark"}},
        "logging": {"level": "info"},
    }

    assert write_toml_file(toml_path, data) is True
    assert read_toml_file(toml_path)["gui_config"]["window"]["width"] == 800

    assert update_toml_value(toml_path, "gui_config.window", "width", 1024) is True
    assert get_toml_value(toml_path, "gui_config.window", "width") == 1024
    assert get_toml_value(toml_path, "missing.section", "k", default="d") == "d"


@pytest.mark.unit
def test_validate_toml_syntax(tmp_path: Path) -> None:
    missing = tmp_path / "missing.toml"
    assert validate_toml_syntax(missing) is False

    good = tmp_path / "good.toml"
    good.write_text('a = 1\nb = "x"\n', encoding="utf-8")
    assert validate_toml_syntax(good) is True

    bad = tmp_path / "bad.toml"
    bad.write_text("a = \n", encoding="utf-8")
    assert validate_toml_syntax(bad) is False


@pytest.mark.unit
def test_extract_inline_comments(tmp_path: Path) -> None:
    toml_path = tmp_path / "c.toml"
    toml_path.write_text(
        '[section]\na = "x" # comment a\nb = 2\n',
        encoding="utf-8",
    )

    comments = extract_inline_comments(toml_path, "section")
    assert comments["a"] == "comment a"
    assert "b" not in comments


@pytest.mark.unit
def test_save_mapping_with_comments(tmp_path: Path) -> None:
    toml_path = tmp_path / "mapping.toml"

    assert (
        save_mapping_with_comments(
            toml_path,
            "mapping",
            mapping_data={"k": ["v1", "v2"]},
            comments_data={"k": "hello"},
        )
        is True
    )

    data = read_toml_file(toml_path)
    assert data["mapping"]["k"] == ["v1", "v2"]
    comments = extract_inline_comments(toml_path, "mapping")
    assert comments["k"] == "hello"
