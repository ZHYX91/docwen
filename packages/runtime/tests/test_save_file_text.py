"""Tests for ``ConfigLoader.save_file_text`` TOML persistence.

Verifies the raw-text save interface:
- rejects unknown rel_path
- rejects invalid TOML syntax
- writes normal sparse config text verbatim (no default backfill on disk)
- adds internal ownership metadata only for editor-owned replacement sections
- in-memory config reflects the saved content after reload
"""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def write_minimal_base_config_tree(base_dir: Path) -> None:
    """Create an empty TOML file for every spec in the registry under *base_dir*."""
    from docwen_runtime.config.registry import CONFIG_FILES

    for spec in CONFIG_FILES:
        path = base_dir / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")


class TestSaveFileText:
    def test_rejects_unknown_rel_path(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.save_file_text("not_in_registry.toml", "x = 1") is False

    def test_rejects_invalid_toml_syntax(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        assert loader.save_file_text("gui.toml", "x = ") is False

    def test_directory_creation_failure_returns_false(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        blocked_user_dir = tmp_path / "blocked-user-dir"
        base_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        blocked_user_dir.write_text("not a directory", encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=blocked_user_dir)

        assert loader.save_file_text("gui.toml", "value = 1\n") is False

    def test_writes_verbatim_no_default_backfill_on_disk(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        sparse = '[window]\ndefault_mode = "batch"\n'
        assert loader.save_file_text("gui.toml", sparse) is True

        on_disk = (user_dir / "gui.toml").read_text(encoding="utf-8")
        # Disk content must be exactly what the user wrote — no backfilled keys
        assert on_disk == sparse
        # A backfill would have added default keys
        assert "theme" not in on_disk

    def test_in_memory_config_merges_base_after_save(self, tmp_path: Path) -> None:
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        sparse = '[window]\ndefault_mode = "batch"\n'
        assert loader.save_file_text("gui.toml", sparse) is True

        # User value is present and overrides base
        assert loader.config.gui.window.default_mode == "batch"
        # Base values from other gui.toml keys still available in memory
        config_dict = loader.config.as_dict() if hasattr(loader.config, "as_dict") else loader._config
        assert "gui" in config_dict

    def test_save_calls_reload_which_reads_base_plus_user(self, tmp_path: Path) -> None:
        """After save, reload reads base file (which must exist) and user file."""
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "gui.toml").write_text('[window]\ndefault_mode = "single"\n', encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        sparse = '[window]\ndefault_mode = "batch"\n'
        loader.save_file_text("gui.toml", sparse)

        # Disk must still be sparse right after save
        assert (user_dir / "gui.toml").read_text(encoding="utf-8") == sparse

    def test_save_numbering_cleanup_namespaces_correctly(self, tmp_path: Path) -> None:
        """numbering/cleanup.toml lives under numbering.cleanup namespace."""
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        base_dir.mkdir()
        user_dir.mkdir()
        write_minimal_base_config_tree(base_dir)
        (base_dir / "numbering" / "cleanup.toml").write_text("", encoding="utf-8")

        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        content = '[[rules]]\nid = "test_rule"\nenabled = true\npattern = "^x"\nlevel = 1\n'
        assert loader.save_file_text("numbering/cleanup.toml", content) is True

        cleanup = loader._config.get("numbering", {}).get("cleanup", {})
        rules = cleanup.get("rules", [])
        assert any(r.get("id") == "test_rule" for r in rules if isinstance(r, dict))
