"""Unit tests for docwen_core.toml_tools primitives."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


class TestParseTomlText:
    def test_parses_string_to_dict(self):
        from docwen_core.toml_tools import parse_toml_text

        data = parse_toml_text("[section]\nvalue = 1\n")
        assert data["section"]["value"] == 1

    def test_invalid_text_raises(self):
        from tomllib import TOMLDecodeError

        from docwen_core.toml_tools import parse_toml_text

        with pytest.raises(TOMLDecodeError):
            parse_toml_text("not = valid = toml = !!!\n")


class TestReadTomlText:
    def test_returns_mutable_document_with_comments(self):
        import tomlkit as tomlkit_module
        from tomlkit import TOMLDocument

        from docwen_core.toml_tools import read_toml_text

        doc = read_toml_text('# keep me\n[settings]\norder = ["a"]\n')
        assert isinstance(doc, TOMLDocument)
        settings = doc["settings"]
        assert isinstance(settings, dict)
        settings["new"] = True  # type: ignore[index]
        text = tomlkit_module.dumps(doc)
        assert "# keep me" in text
        assert "new = true" in text

    def test_read_then_parse_compatible(self):
        from docwen_core.toml_tools import parse_toml_text, read_toml_text

        via_read = read_toml_text("[s]\nv = 1\n")
        via_parse = parse_toml_text("[s]\nv = 1\n")
        assert via_read["s"]["v"] == via_parse["s"]["v"] == 1


class TestTomlValue:
    def test_value_without_comment(self):
        from docwen_core.toml_tools import toml_value

        item = toml_value(["a", "b"])
        assert item == ["a", "b"]

    def test_value_with_comment(self):
        import tomlkit

        from docwen_core.toml_tools import toml_value

        item = toml_value(["纠正"], comment="用户备注")
        tbl = tomlkit.table()
        tbl["错字"] = item
        text = tomlkit.dumps(tbl)
        assert "用户备注" in text
        assert "错字" in text

    def test_empty_comment_is_noop(self):
        from docwen_core.toml_tools import toml_value

        item = toml_value(42, comment="")
        assert item == 42


class TestTomlTable:
    def test_empty_table(self):
        import tomlkit

        from docwen_core.toml_tools import toml_table

        tbl = toml_table()
        tbl["key"] = "value"
        doc = tomlkit.document()
        doc["section"] = tbl
        text = tomlkit.dumps(doc)
        assert "[section]" in text
        assert 'key = "value"' in text
