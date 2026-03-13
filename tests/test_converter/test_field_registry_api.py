from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.md2docx.field_registry import (
    get_available_processors,
    get_merged_placeholder_rules,
    reset_field_registry_cache,
    set_processor_enabled,
)

pytestmark = pytest.mark.unit


def _write_field_processors_toml(path: Path, *, content: str) -> None:
    path.write_text(content, encoding="utf-8")


def test_set_processor_enabled_toggles_effective_rules(monkeypatch, tmp_path: Path) -> None:
    config_path = tmp_path / "field_processors.toml"
    _write_field_processors_toml(
        config_path,
        content=(
            "[settings]\n"
            'order = ["gongwen"]\n'
            "\n"
            "[processors.gongwen]\n"
            'module = "docwen.converter.md2docx.fields.gongwen"\n'
            'locales = ["zh_CN"]\n'
            "enabled = false\n"
        ),
    )

    monkeypatch.setenv("DOCWEN_FIELD_PROCESSORS_CONFIG", str(config_path))
    reset_field_registry_cache()

    rules_disabled = get_merged_placeholder_rules("zh_CN")
    assert len(rules_disabled.get("delete_paragraph_if_empty", [])) == 0

    assert set_processor_enabled("gongwen", True) is True

    rules_enabled = get_merged_placeholder_rules("zh_CN")
    assert len(rules_enabled.get("delete_paragraph_if_empty", [])) > 0
    rules_en = get_merged_placeholder_rules("en_US")
    assert len(rules_en.get("delete_paragraph_if_empty", [])) == 0


def test_get_available_processors_filters_by_locale(monkeypatch, tmp_path: Path) -> None:
    config_path = tmp_path / "field_processors.toml"
    _write_field_processors_toml(
        config_path,
        content=(
            "[settings]\n"
            'order = ["gongwen", "en_only"]\n'
            "\n"
            "[processors.gongwen]\n"
            'module = "docwen.converter.md2docx.fields.gongwen"\n'
            'name = "公文字段优化"\n'
            'locales = ["zh_CN"]\n'
            "enabled = true\n"
            "\n"
            "[processors.en_only]\n"
            'module = "docwen.converter.md2docx.fields.gongwen"\n'
            'name = "EN Only"\n'
            'locales = ["en_US"]\n'
            "enabled = true\n"
        ),
    )

    monkeypatch.setenv("DOCWEN_FIELD_PROCESSORS_CONFIG", str(config_path))
    reset_field_registry_cache()

    zh_list = get_available_processors("zh_CN")
    assert [p["id"] for p in zh_list] == ["gongwen"]

    en_list = get_available_processors("en_US")
    assert [p["id"] for p in en_list] == ["en_only"]


def test_module_import_failure_is_downgraded(monkeypatch, tmp_path: Path) -> None:
    config_path = tmp_path / "field_processors.toml"
    _write_field_processors_toml(
        config_path,
        content=(
            "[settings]\n"
            'order = ["broken"]\n'
            "\n"
            "[processors.broken]\n"
            'module = "docwen.__definitely_not_exists__"\n'
            'locales = ["*"]\n'
            "enabled = true\n"
        ),
    )

    monkeypatch.setenv("DOCWEN_FIELD_PROCESSORS_CONFIG", str(config_path))
    reset_field_registry_cache()

    rules = get_merged_placeholder_rules("zh_CN")
    assert len(rules.get("delete_paragraph_if_empty", [])) == 0

    processors = get_available_processors("zh_CN")
    assert processors and processors[0]["id"] == "broken"
    assert "load_error" in processors[0]


def test_missing_module_is_not_activated(monkeypatch, tmp_path: Path) -> None:
    config_path = tmp_path / "field_processors.toml"
    _write_field_processors_toml(
        config_path,
        content=('[settings]\norder = ["no_module"]\n\n[processors.no_module]\nlocales = ["*"]\nenabled = true\n'),
    )

    monkeypatch.setenv("DOCWEN_FIELD_PROCESSORS_CONFIG", str(config_path))
    reset_field_registry_cache()

    rules = get_merged_placeholder_rules("zh_CN")
    assert len(rules.get("delete_paragraph_if_empty", [])) == 0

    processors = get_available_processors("zh_CN")
    assert processors and processors[0]["id"] == "no_module"
    assert processors[0]["load_error"] == "missing module"


def test_missing_config_file_is_safe(monkeypatch, tmp_path: Path) -> None:
    missing = tmp_path / "missing.toml"
    monkeypatch.setenv("DOCWEN_FIELD_PROCESSORS_CONFIG", str(missing))
    reset_field_registry_cache()

    rules = get_merged_placeholder_rules("zh_CN")
    assert len(rules.get("delete_paragraph_if_empty", [])) == 0
    assert get_available_processors("zh_CN") == []
    assert set_processor_enabled("gongwen", True) is False
