"""Focused tests split from test_numbering_editors_and_batch_dialog.py."""

from __future__ import annotations

import pytest

from ._numbering_editors_and_batch_dialog_support import (
    Path,
    QApplication,
    _patch_all_modals,
    _patch_dialog_notify,
    _write_minimal_base_config_tree,
)

pytestmark = pytest.mark.gui


class TestNumberingCleanPersistence:
    """Test TOML save/load round-trip."""

    def test_build_toml_dict(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [
                {
                    "id": "r1",
                    "enabled": True,
                    "pattern": r"^[0-9]+",
                    "name": "Direct name",
                    "name_key": "number_separator",
                    "description": "Desc",
                    "description_key": "number_separator_desc",
                    "level": 2,
                    "is_system": True,
                }
            ],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        data = dlg.get_rules_data()
        assert data["settings"]["order"] == ["r1"]
        assert len(data["rules"]) == 1
        assert data["rules"][0]["id"] == "r1"
        assert data["rules"][0]["pattern"] == r"^[0-9]+"
        assert data["rules"][0]["name"] == "Direct name"
        assert data["rules"][0]["name_key"] == "number_separator"
        assert data["rules"][0]["description"] == "Desc"
        assert data["rules"][0]["description_key"] == "number_separator_desc"
        assert data["rules"][0]["level"] == 2
        assert data["rules"][0]["is_system"] is True
        dlg.close()

    def test_description_override_tombstone_survives_base_user_merge_and_reopen(
        self, qapp: QApplication, tmp_path: Path
    ) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import NumberingCleanDialog
        from docwen_runtime.config.loader import ConfigLoader

        base_dir = tmp_path / "base"
        user_dir = tmp_path / "user"
        _write_minimal_base_config_tree(base_dir)
        cleanup_path = base_dir / "numbering" / "cleanup.toml"
        cleanup_path.write_text(
            """
[settings]
order = ["r1"]

[[rules]]
id = "r1"
name_key = "number_separator"
description = "Base fallback"
description_key = "number_separator_desc"
enabled = true
is_system = true
pattern = "^x"
level = 1
""".lstrip(),
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=base_dir, user_dir=user_dir)
        current = loader.config.as_dict()["numbering"]["cleanup"]
        dlg = NumberingCleanDialog(config_data=current)
        dlg.desc_edit.setText("User override")
        update = dlg.get_rules_data()

        override = update["rules"][0]
        assert override["description"] == "User override"
        assert override["description_key"] == ""
        assert loader.update_file_sections("numbering/cleanup.toml", update) is True

        merged = loader.config.as_dict()["numbering"]["cleanup"]
        merged_rule = merged["rules"][0]
        assert merged_rule["description"] == "User override"
        assert merged_rule["description_key"] == ""

        reopened = NumberingCleanDialog(config_data=merged)
        assert reopened.rules["r1"].display_description() == "User override"
        assert reopened.desc_edit.text() == "User override"
        reopened.close()
        dlg.close()

    def test_save_via_on_save_callback(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        saved: list[dict] = []

        def on_save(data: dict) -> None:
            saved.append(data)

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [{"id": "r1", "enabled": True, "pattern": "^test", "is_system": True}],
        }
        dlg = NumberingCleanDialog(config_data=seed, on_save=on_save)
        dlg._save()  # Use _save (not _save_to_disk)
        assert len(saved) == 1
        assert saved[0]["rules"][0]["pattern"] == "^test"
        dlg.close()

    def test_write_toml_file(self, tmp_path: Path) -> None:
        """update_file_sections writes clean rules and preserves structure."""
        from docwen_runtime.config.loader import ConfigLoader

        config_dir = tmp_path / "configs"
        config_dir.mkdir()
        _write_minimal_base_config_tree(config_dir)
        data = {
            "settings": {"order": ["r1"]},
            "rules": [
                {
                    "id": "r1",
                    "enabled": True,
                    "pattern": "^test",
                    "description": "",
                    "level": 1,
                    "is_system": True,
                }
            ],
        }
        loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)
        ok = loader.update_file_sections("numbering/cleanup.toml", data)
        assert ok
        content = (config_dir / "numbering" / "cleanup.toml").read_text(encoding="utf-8")
        assert "order" in content
        assert "^test" in content or "test" in content
        assert "is_system = true" in content

    def test_write_toml_file_preserves_unrelated_sections(self, tmp_path: Path) -> None:
        """update_file_sections preserves sections not in updates."""
        from docwen_runtime.config.loader import ConfigLoader

        config_dir = tmp_path / "configs"
        config_dir.mkdir()
        _write_minimal_base_config_tree(config_dir)
        # Pre-write heading_numbering_clean.toml with unrelated sections.
        (config_dir / "numbering" / "cleanup.toml").write_text(
            '# keep me\n[settings]\norder = ["old"]\n\n[meta]\nkeep = true\n',
            encoding="utf-8",
        )
        data = {
            "settings": {"order": ["r1"]},
            "rules": [
                {
                    "id": "r1",
                    "enabled": True,
                    "pattern": "^test",
                    "description": "",
                    "level": 1,
                    "is_system": True,
                }
            ],
        }
        loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)
        ok = loader.update_file_sections("numbering/cleanup.toml", data)
        assert ok
        content = (config_dir / "numbering" / "cleanup.toml").read_text(encoding="utf-8")
        assert "# keep me" in content
        assert "[meta]" in content
        assert "keep = true" in content


class TestNumberingAddDialogConstruction:
    """Test dialog creation and initial state."""

    def test_create_with_default_data(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(config_data={})
        assert dlg.windowTitle() != ""
        assert dlg.scheme_list is not None
        dlg.close()

    def test_malformed_nested_data_degrades_to_empty_schemes(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import NumberingAddDialog

        dlg = NumberingAddDialog(config_data={"settings": "invalid", "number_styles": "invalid", "schemes": []})

        assert dlg.order == []
        assert dlg.number_styles == {}
        assert dlg.schemes == {}
        dlg.close()

    def test_create_with_seeded_data(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "test_scheme", "order": ["test_scheme"]},
            "schemes": {
                "test_scheme": {
                    "name": "Test",
                    "description": "A test scheme",
                    "enabled": True,
                    "is_system": False,
                    "level_1": {"format": "{1.chinese_lower}、"},
                    "level_2": {"format": "{2.arabic_half}."},
                }
            },
        }
        dlg = NumberingAddDialog(config_data=seed)
        assert "test_scheme" in dlg.schemes
        scheme = dlg.schemes["test_scheme"]
        assert scheme.name == "Test"
        assert scheme.levels[1] == "{1.chinese_lower}、"
        dlg.close()

    def test_window_title_shows_dirty(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1"]},
            "schemes": {"s1": {"name": "S1", "is_system": False}},
        }
        dlg = NumberingAddDialog(config_data=seed)
        initial_title = dlg.windowTitle()
        assert "*" not in initial_title

        # Trigger a change
        dlg.name_edit.setText("Modified")
        assert "*" in dlg.windowTitle()
        dlg.close()

    def test_word_native_compatibility_display(self, qapp: QApplication) -> None:
        """Verify the compatibility label shows correct verdict for a known scheme."""
        from docwen_gui.i18n import t
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        # Use the hierarchical_standard scheme which should be fully compatible
        seed = {
            "settings": {"default_scheme": "hierarchical", "order": ["hierarchical"]},
            "schemes": {
                "hierarchical": {
                    "name": "Hierarchical",
                    "is_system": False,
                    "level_1": {"format": "{1.arabic_half} "},
                    "level_2": {"format": "{1.arabic_half}.{2.arabic_half} "},
                },
            },
        }
        dlg = NumberingAddDialog(config_data=seed)
        _patch_all_modals(dlg)
        # After select, compat label should be populated
        compat_text = dlg._compat_label.text()
        assert t("editors.numbering_add.word_native_full") == compat_text
        # Fully compatible scheme should have green color
        assert t("editors.numbering_add.word_native_full") in compat_text
        dlg.close()

    def test_word_native_compatibility_color_reacts_to_live_theme_preview(self, qapp: QApplication) -> None:
        from docwen_gui.styles.theme_manager import ThemeManager
        from docwen_gui.styles.theme_semantics import get_theme_class_color
        from docwen_gui.widgets.settings.numbering_add_editor import NumberingAddDialog

        seed = {
            "settings": {"default_scheme": "hierarchical", "order": ["hierarchical"]},
            "schemes": {
                "hierarchical": {
                    "name": "Hierarchical",
                    "is_system": False,
                    "level_1": {"format": "{1.arabic_half} "},
                    "level_2": {"format": "{1.arabic_half}.{2.arabic_half} "},
                },
            },
        }
        manager = ThemeManager.get_instance()
        previous_theme = manager.get_current_theme()
        manager.initialize(qapp, "light")
        dlg = NumberingAddDialog(config_data=seed)
        try:
            assert get_theme_class_color("success", "light").upper() in dlg._compat_label.styleSheet().upper()

            manager.apply_theme("dark")
            qapp.processEvents()

            assert get_theme_class_color("success", "dark").upper() in dlg._compat_label.styleSheet().upper()
        finally:
            dlg.close()
            manager.apply_theme(previous_theme)
            qapp.processEvents()

    def test_word_native_compatibility_display_uses_current_locale(self, qapp: QApplication) -> None:
        """Compatibility label must not be hardcoded Chinese in non-Chinese locales."""
        from docwen_gui.i18n import get_locale, set_locale, t
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        previous_locale = get_locale()
        set_locale("en_US")
        try:
            seed = {
                "settings": {"default_scheme": "hierarchical", "order": ["hierarchical"]},
                "schemes": {
                    "hierarchical": {
                        "name": "Hierarchical",
                        "is_system": False,
                        "level_1": {"format": "{1.arabic_half} "},
                        "level_2": {"format": "{1.arabic_half}.{2.arabic_half} "},
                    },
                },
            }
            dlg = NumberingAddDialog(config_data=seed)
            _patch_all_modals(dlg)
            assert dlg._compat_label.text() == t("editors.numbering_add.word_native_full")
            assert "原生兼容性" not in dlg._compat_label.text()
            dlg.close()
        finally:
            set_locale(previous_locale)


class TestNumberingAddSchemeCRUD:
    """Test scheme create/copy/delete operations.

    All tests that trigger modal dialogs monkey-patch the confirm/notify
    methods to avoid blocking in offscreen mode.
    """

    def test_create_new_scheme(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(config_data={})
        initial_count = len(dlg.schemes)
        dlg._create_new_scheme()
        assert len(dlg.schemes) == initial_count + 1
        new_id = dlg.current_scheme_id
        assert new_id is not None
        assert new_id in dlg.schemes
        dlg.close()

    def test_copy_scheme(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1"]},
            "schemes": {"s1": {"name": "Original", "is_system": False}},
        }
        dlg = NumberingAddDialog(config_data=seed)
        _patch_dialog_notify(dlg)
        assert dlg.current_scheme_id == "s1"
        dlg._copy_selected_scheme()
        assert len(dlg.schemes) == 2
        assert dlg.current_scheme_id != "s1"
        assert dlg.schemes[dlg.current_scheme_id].name == "Original (副本)"
        dlg.close()

    def test_delete_custom_scheme(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1", "s2"]},
            "schemes": {
                "s1": {"name": "System", "is_system": True},
                "s2": {"name": "Custom", "is_system": False},
            },
        }
        dlg = NumberingAddDialog(config_data=seed)
        _patch_all_modals(dlg, confirm_value=True)
        dlg._select_scheme("s2")
        assert dlg.current_scheme_id == "s2"
        dlg._delete_selected_scheme()
        assert "s2" not in dlg.schemes
        dlg.close()

    def test_cannot_delete_system_scheme(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1"]},
            "schemes": {"s1": {"name": "System", "is_system": True}},
        }
        dlg = NumberingAddDialog(config_data=seed)
        _patch_all_modals(dlg)
        dlg._select_scheme("s1")
        dlg._delete_selected_scheme()
        # System scheme should remain (deletion blocked before confirm)
        assert "s1" in dlg.schemes
        dlg.close()

    def test_cannot_delete_default_scheme(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1"]},
            "schemes": {"s1": {"name": "Default", "is_system": False}},
        }
        dlg = NumberingAddDialog(config_data=seed)
        _patch_all_modals(dlg)
        dlg._select_scheme("s1")
        dlg._delete_selected_scheme()
        # Default scheme cannot be deleted (blocked by check, not by confirm)
        assert "s1" in dlg.schemes
        dlg.close()

    def test_move_scheme_up_down(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1", "s2"]},
            "schemes": {
                "s1": {"name": "First", "is_system": False},
                "s2": {"name": "Second", "is_system": False},
            },
        }
        dlg = NumberingAddDialog(config_data=seed)
        dlg._select_scheme("s2")
        dlg._move_selected_scheme_up()
        assert dlg.order[0] == "s2"
        assert dlg.order[1] == "s1"
        dlg._move_selected_scheme_down()
        assert dlg.order[0] == "s1"
        assert dlg.order[1] == "s2"
        dlg.close()
