"""One-time renderer migration preserves effective output and later choices."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

from docwen_runtime.config.heading_numbering_migration import migrate_heading_numbering_preferences
from docwen_runtime.config.loader import ConfigLoader
from docwen_runtime.config.validation import ConfigSemanticError, validate_config_file

from ._config_loader_support import write_minimal_base_config_tree

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("legacy", [True, False, None])
@pytest.mark.parametrize("mode", ["text", "word_native"])
def test_legacy_renderer_migrates_once(tmp_path: Path, mode: str, legacy: bool | None):
    base, user = tmp_path / "base", tmp_path / "user"
    write_minimal_base_config_tree(base)
    (base / "text.toml").write_text('heading_numbering_render_mode = "text"\n', encoding="utf-8")
    user.mkdir()
    (user / "text.toml").write_text(f'heading_numbering_render_mode = "{mode}"\n', encoding="utf-8")
    if legacy is not None:
        (user / "template_fill.toml").write_text(
            f"gongwen_numbering_text = {str(legacy).lower()}\ngongwen_numbering_comments = true\n",
            encoding="utf-8",
        )
    loader = ConfigLoader(base_dir=base, user_dir=user)
    expected = mode if legacy is False else "text"
    assert loader.config.text.heading_numbering_render_mode == expected
    persisted = tomllib.loads((user / "text.toml").read_text(encoding="utf-8"))
    assert persisted["heading_numbering_migration_version"] == 1
    assert persisted["heading_numbering_render_mode"] == expected
    assert loader.set_value("text.heading_numbering_render_mode", "word_native")
    restarted = ConfigLoader(base_dir=base, user_dir=user)
    assert restarted.config.text.heading_numbering_render_mode == "word_native"


def test_fresh_profile_defaults_to_text_and_runtime_override_wins(tmp_path: Path):
    base, user = tmp_path / "base", tmp_path / "user"
    write_minimal_base_config_tree(base)
    (base / "text.toml").write_text('heading_numbering_render_mode = "text"\n', encoding="utf-8")
    loader = ConfigLoader(base_dir=base, user_dir=user)
    assert loader.config.text.heading_numbering_render_mode == "text"
    assert not user.exists()
    overridden = ConfigLoader(
        base_dir=base,
        user_dir=user,
        runtime_overrides={"text": {"heading_numbering_render_mode": "word_native"}},
    )
    assert overridden.config.text.heading_numbering_render_mode == "word_native"


def test_migration_is_nonmutating_and_does_not_override_new_native_choice():
    source = {"heading_numbering_render_mode": "word_native", "heading_numbering_migration_version": 1}
    assert migrate_heading_numbering_preferences(source, {"gongwen_numbering_text": True}) == source
    old = {"heading_numbering_render_mode": "word_native", "numbering_scheme": "custom"}
    migrated = migrate_heading_numbering_preferences(old, {})
    assert old["heading_numbering_render_mode"] == "word_native"
    assert migrated["numbering_scheme"] == "custom"


def test_legacy_text_flag_no_longer_clears_numbering_comments():
    validated = validate_config_file(
        "template_fill.toml",
        {"gongwen_numbering_text": False, "gongwen_numbering_comments": True},
        {},
    )
    assert validated["gongwen_numbering_comments"] is True


@pytest.mark.parametrize("value", ["true", 1, None, []])
def test_document_marker_option_requires_boolean(value):
    with pytest.raises(ConfigSemanticError):
        validate_config_file("document.toml", {"export_numbering_markers": value}, {})


@pytest.mark.parametrize("value", [True, False])
def test_document_marker_option_accepts_boolean(value):
    assert validate_config_file("document.toml", {"export_numbering_markers": value}, {}) == {
        "export_numbering_markers": value
    }


@pytest.mark.parametrize(
    "invalid", ['heading_numbering_render_mode = "bad"', "heading_numbering_migration_version = true"]
)
def test_invalid_text_override_quarantines_then_accepts_new_native_choice(tmp_path: Path, invalid: str):
    base, user = tmp_path / "base", tmp_path / "user"
    write_minimal_base_config_tree(base)
    (base / "text.toml").write_text('heading_numbering_render_mode = "text"\n', encoding="utf-8")
    user.mkdir()
    (user / "text.toml").write_text(invalid, encoding="utf-8")
    loader = ConfigLoader(base_dir=base, user_dir=user)
    assert loader.config.text.heading_numbering_render_mode == "text"
    assert list(user.glob("text.toml.bak_*"))
    assert loader.set_value("text.heading_numbering_render_mode", "word_native")
    assert ConfigLoader(base_dir=base, user_dir=user).config.text.heading_numbering_render_mode == "word_native"


@pytest.mark.parametrize(
    "method",
    ["set_value", "set_values", "write_file_content", "save_file_text", "update_file_sections", "update_file_document"],
)
def test_first_native_save_survives_restart_without_read_only_init_writes(tmp_path: Path, method: str):
    base, user = tmp_path / "base", tmp_path / "user"
    write_minimal_base_config_tree(base)
    (base / "text.toml").write_text('heading_numbering_render_mode = "text"\n', encoding="utf-8")
    loader = ConfigLoader(base_dir=base, user_dir=user)
    assert not user.exists()
    value = {"heading_numbering_render_mode": "word_native"}
    if method == "set_value":
        saved = loader.set_value("text.heading_numbering_render_mode", "word_native")
    elif method == "set_values":
        saved = loader.set_values({"text.heading_numbering_render_mode": "word_native"})
    elif method == "save_file_text":
        saved = loader.save_file_text(
            "text.toml", '# Keep this comment\nheading_numbering_render_mode = "word_native"\n'
        )
        assert "# Keep this comment" in (user / "text.toml").read_text(encoding="utf-8")
    elif method == "update_file_document":
        saved = loader.update_file_document("text.toml", lambda doc: doc.update(value))
    else:
        saved = getattr(loader, method)("text.toml", value)
    assert saved
    assert ConfigLoader(base_dir=base, user_dir=user).config.text.heading_numbering_render_mode == "word_native"
    assert tomllib.loads((user / "text.toml").read_text(encoding="utf-8"))["heading_numbering_migration_version"] == 1
