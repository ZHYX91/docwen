from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

from docwen_runtime.toml_io import (
    load_toml_document,
    new_toml_document,
    read_toml_file,
    save_toml_document,
    write_toml_file,
)


def test_atomic_writer_is_exported_from_config_public_surface() -> None:
    from docwen_runtime.config import atomic_write_text as public_atomic_write_text
    from docwen_runtime.toml_io import atomic_write_text as implementation_atomic_write_text

    assert public_atomic_write_text is implementation_atomic_write_text


def test_read_and_write_toml_file_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "sample.toml"

    write_toml_file(path, {"section": {"enabled": True, "order": ["a", "b"]}})

    data = read_toml_file(path)
    assert data["section"]["enabled"] is True
    assert list(data["section"]["order"]) == ["a", "b"]


def test_load_and_save_toml_document_preserves_comments(tmp_path: Path) -> None:
    path = tmp_path / "sample.toml"
    path.write_text("[section]\nvalue = 1  # keep me\n", encoding="utf-8", newline="\n")

    doc = load_toml_document(path)
    doc["section"]["value"] = 2
    save_toml_document(path, doc)

    assert "# keep me" in path.read_text(encoding="utf-8")
    assert read_toml_file(path)["section"]["value"] == 2


def test_new_toml_document_can_be_saved(tmp_path: Path) -> None:
    path = tmp_path / "created.toml"

    doc = new_toml_document()
    doc["root"] = {"name": "docwen"}
    save_toml_document(path, doc)

    assert read_toml_file(path)["root"]["name"] == "docwen"


# ── update_toml_document_sections ───────────────────────────────────────────


class TestUpdateTomlDocumentSections:
    """update_toml_document_sections 把 dict 的指定 section 合并进已有
    tomlkit document，保留未覆盖部分的注释与键序。"""

    def test_preserves_comments_in_untouched_sections(self, tmp_path):
        from docwen_runtime.toml_io import (
            load_toml_document,
            save_toml_document,
            update_toml_document_sections,
        )

        path = tmp_path / "cfg.toml"
        path.write_text(
            '# 顶部注释\n[settings]\norder = ["a"]\n\n[rules]\n# 规则注释\nid = "x"\n',
            encoding="utf-8",
        )
        doc = load_toml_document(path)
        update_toml_document_sections(doc, {"settings": {"order": ["b", "a"]}})
        save_toml_document(path, doc)

        text = path.read_text(encoding="utf-8")
        assert "# 顶部注释" in text
        assert "# 规则注释" in text  # inline comment inside [rules]
        assert 'order = ["b", "a"]' in text

    def test_replaces_section_value_when_key_present(self, tmp_path):
        from tomlkit import table

        from docwen_runtime.toml_io import (
            new_toml_document,
            save_toml_document,
            update_toml_document_sections,
        )

        doc = new_toml_document()
        doc["settings"] = table()
        doc["settings"]["order"] = ["old"]
        update_toml_document_sections(doc, {"settings": {"order": ["new"]}})
        path = tmp_path / "cfg.toml"
        save_toml_document(path, doc)
        assert 'order = ["new"]' in path.read_text(encoding="utf-8")

    def test_adds_section_when_absent(self, tmp_path):
        from docwen_runtime.toml_io import (
            new_toml_document,
            save_toml_document,
            update_toml_document_sections,
        )

        doc = new_toml_document()
        update_toml_document_sections(doc, {"schemes": {"gongwen": {"enabled": True}, "custom": {"enabled": False}}})
        path = tmp_path / "cfg.toml"
        save_toml_document(path, doc)
        text = path.read_text(encoding="utf-8")
        assert "[schemes" in text
        assert "gongwen" in text
        assert "custom" in text

    def test_empty_updates_dict_is_noop(self, tmp_path):
        from docwen_runtime.toml_io import (
            load_toml_document,
            save_toml_document,
            update_toml_document_sections,
        )

        path = tmp_path / "cfg.toml"
        original = '# keep\n[settings]\norder = ["a"]\n'
        path.write_text(original, encoding="utf-8")
        doc = load_toml_document(path)
        update_toml_document_sections(doc, {})
        save_toml_document(path, doc)
        assert path.read_text(encoding="utf-8") == original

    def test_does_not_touch_sections_not_in_updates(self, tmp_path):
        from docwen_runtime.toml_io import (
            load_toml_document,
            save_toml_document,
            update_toml_document_sections,
        )

        path = tmp_path / "cfg.toml"
        path.write_text(
            '# keep me\n[rules]\nid = "x"\n[settings]\norder = ["a"]\n',
            encoding="utf-8",
        )
        doc = load_toml_document(path)
        update_toml_document_sections(doc, {"settings": {"order": ["b"]}})
        save_toml_document(path, doc)
        text = path.read_text(encoding="utf-8")
        assert "# keep me" in text
        assert 'id = "x"' in text  # rules section preserved as-is


# ── read / load type contract ───────────────────────────────────────────────


def test_read_and_edit_toml_apis_have_distinct_types(tmp_path: Path) -> None:
    """Read-only parsing stays plain; editing explicitly preserves formatting."""
    path = tmp_path / "cfg.toml"
    path.write_text("[section]\nvalue = 1\n", encoding="utf-8")

    via_read = read_toml_file(path)
    via_load = load_toml_document(path)

    assert via_read["section"]["value"] == 1
    assert via_load["section"]["value"] == 1
    from tomlkit import TOMLDocument

    assert type(via_read) is dict
    assert isinstance(via_load, TOMLDocument)
