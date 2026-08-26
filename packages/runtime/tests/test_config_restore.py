"""Section-level config restore tests for ConfigLoader (registry-driven user override).

Validates that ``reset_section`` resets a config section to base defaults
by deleting the user override file for the owning spec.  Each file has
exactly one namespace, so section-level reset == file-level reset.
"""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def write_minimal_base_config_tree(base_dir: Path) -> None:
    """Create an empty TOML file for every spec in the registry under *base_dir*."""
    from docwen_runtime.config.registry import CONFIG_FILES

    for spec in CONFIG_FILES:
        path = base_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")


# ---------------------------------------------------------------------------
# reset_section (routing convenience)
# ---------------------------------------------------------------------------


class TestResetSectionRouting:
    """Contract tests for ConfigLoader.reset_section()."""

    def test_resets_root_section(self, tmp_path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")
        (user_dir / "gui.toml").parent.mkdir(parents=True, exist_ok=True)
        (user_dir / "gui.toml").write_text('[window]\ndefault_mode = "batch"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.config.gui.window.default_mode == "batch"

        ok = loader.reset_section("gui")
        assert ok is True
        assert loader.config.gui.window.default_mode == "single"

    def test_resets_nested_section(self, tmp_path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "link.toml").write_text('[format]\nimage_link_style = "wiki_embed"\n', encoding="utf-8")
        (user_dir / "link.toml").parent.mkdir(parents=True, exist_ok=True)
        (user_dir / "link.toml").write_text('[format]\nimage_link_style = "markdown_link"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.config.link.format.image_link_style == "markdown_link"

        ok = loader.reset_section("link.format")
        assert ok is True
        assert loader.config.link.format.image_link_style == "wiki_embed"

    def test_unknown_root_returns_false(self, tmp_path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        ok = loader.reset_section("no_such_root.something")
        assert ok is False

    def test_empty_dotted_section_returns_false(self, tmp_path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.reset_section("") is False

    def test_reset_section_deletes_user_file(self, tmp_path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")
        (user_dir / "gui.toml").parent.mkdir(parents=True, exist_ok=True)
        (user_dir / "gui.toml").write_text('[window]\ndefault_mode = "batch"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        loader.reset_section("gui")
        assert not (user_dir / "gui.toml").exists()

    def test_reset_subdirectory_section(self, tmp_path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "proofread" / "engine.toml").write_text("enable_typos_rule = true\n", encoding="utf-8")
        (user_dir / "proofread").mkdir(parents=True, exist_ok=True)
        (user_dir / "proofread" / "engine.toml").write_text("enable_typos_rule = false\n", encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.config.proofread.engine.enable_typos_rule is False

        ok = loader.reset_section("proofread.engine")
        assert ok is True
        assert loader.config.proofread.engine.enable_typos_rule is True
        assert not (user_dir / "proofread" / "engine.toml").exists()


# ---------------------------------------------------------------------------
# Isolation: resetting one file does not affect other files
# ---------------------------------------------------------------------------


class TestResetIsolation:
    """Resetting one spec's user file must leave other files untouched."""

    def test_reset_gui_does_not_affect_output(self, tmp_path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")
        (base_dir / "output.toml").write_text('[directory]\nmode = "source"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.set_value("gui.window.default_mode", "batch") is True
        assert loader.set_value("output.directory.mode", "custom") is True

        ok = loader.reset_section("gui")
        assert ok is True

        assert loader.config.gui.window.default_mode == "single"
        assert loader.config.output.directory.mode == "custom"

    def test_reset_proofread_does_not_affect_link(self, tmp_path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "link.toml").write_text('[format]\nimage_link_style = "wiki_embed"\n', encoding="utf-8")
        (base_dir / "proofread" / "engine.toml").write_text("enable_typos_rule = true\n", encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.set_value("link.format.image_link_style", "markdown_link") is True
        assert loader.set_value("proofread.engine.enable_typos_rule", False) is True

        ok = loader.reset_section("proofread.engine")
        assert ok is True

        assert loader.config.proofread.engine.enable_typos_rule is True
        assert loader.config.link.format.image_link_style == "markdown_link"
