"""Focused tests split from test_numbering_editors_and_batch_dialog.py."""

from __future__ import annotations

import pytest

from ._numbering_editors_and_batch_dialog_support import (
    Path,
    QApplication,
    _patch_all_modals,
    _write_minimal_base_config_tree,
)

pytestmark = pytest.mark.gui


class TestNumberingAddLevelEditing:
    """Test per-level format editing, validation, and preview."""

    def test_level_edit_updates_preview(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1"]},
            "schemes": {"s1": {"name": "Test", "is_system": False}},
        }
        dlg = NumberingAddDialog(config_data=seed)
        _patch_all_modals(dlg)
        preview_text = dlg.preview_text.toPlainText()
        assert "no_format" in preview_text.lower() or "暂无格式定义" in preview_text

        dlg.level_edits[1].setText("{1.chinese_lower}、")
        preview_text = dlg.preview_text.toPlainText()
        assert "一、" in preview_text
        dlg.close()

    def test_placeholder_insertion(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(config_data={})
        dlg._create_new_scheme()
        edit = dlg.level_edits[1]
        edit.clear()
        dlg._insert_placeholder(1, "{1.chinese_lower}")
        assert "{1.chinese_lower}" in edit.text()
        dlg.close()

    def test_validation_valid_format(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(config_data={})
        dlg._create_new_scheme()
        dlg._validate_level(1, "{1.chinese_lower}、{2.arabic_half}")
        status = dlg.level_status_labels[1]
        assert status.text() == "✓"  # ✓
        dlg.close()

    def test_validation_fixed_text(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(config_data={})
        dlg._create_new_scheme()
        dlg._validate_level(1, "No placeholders here")
        status = dlg.level_status_labels[1]
        assert status.text() == "!"
        dlg.close()

    def test_validation_unknown_style(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(config_data={})
        dlg._create_new_scheme()
        dlg._validate_level(1, "{1.unknown_style}")
        status = dlg.level_status_labels[1]
        assert status.text() == "x"
        dlg.close()

    def test_preview_renders_all_levels(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1"]},
            "schemes": {
                "s1": {
                    "name": "Full",
                    "is_system": False,
                    "level_1": {"format": "{1.chinese_lower}、"},
                    "level_2": {"format": "{2.arabic_half}."},
                    "level_3": {"format": "{3.arabic_half}."},
                }
            },
        }
        dlg = NumberingAddDialog(config_data=seed)
        preview = dlg.preview_text.toPlainText()
        assert "一、" in preview
        assert "2." in preview
        dlg.close()


class TestNumberingAddStyleSamples:
    """Test the style sample renderer (static methods, no dialogs needed)."""

    def test_chinese_lower(self) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        assert NumberingAddDialog._style_sample("chinese_lower", 1) == "一"  # 一
        assert NumberingAddDialog._style_sample("chinese_lower", 3) == "三"  # 三

    def test_arabic_half(self) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        assert NumberingAddDialog._style_sample("arabic_half", 5) == "5"

    def test_roman_upper(self) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        assert NumberingAddDialog._style_sample("roman_upper", 1) == "I"
        assert NumberingAddDialog._style_sample("roman_upper", 4) == "IV"
        assert NumberingAddDialog._style_sample("roman_upper", 9) == "IX"

    def test_render_format(self) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(config_data={})
        result = dlg._render_format("{1.chinese_lower}、{2.arabic_half}")
        assert "一" in result  # 一
        assert "2" in result
        dlg.close()


class TestNumberingAddPersistence:
    """Test TOML save/load round-trip."""

    def test_build_toml_dict(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1"]},
            "number_styles": {
                "custom_style": {
                    "name": "Custom Style",
                    "description": "Custom style description",
                }
            },
            "schemes": {
                "s1": {
                    "name": "Test",
                    "description": "Desc",
                    "enabled": True,
                    "is_system": False,
                    "level_1": {"format": "{1.chinese_lower}、"},
                }
            },
        }
        dlg = NumberingAddDialog(config_data=seed)
        data = dlg.get_schemes_data()
        assert data["settings"]["default_scheme"] == "s1"
        assert data["number_styles"]["custom_style"]["name"] == "Custom Style"
        assert data["schemes"]["s1"]["name"] == "Test"
        assert data["schemes"]["s1"]["level_1"]["format"] == "{1.chinese_lower}、"
        dlg.close()

    def test_save_via_on_save_callback(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        saved: list[dict] = []

        def on_save(data: dict) -> None:
            saved.append(data)

        seed = {
            "settings": {"default_scheme": "s1", "order": ["s1"]},
            "schemes": {"s1": {"name": "CB", "is_system": False}},
        }
        dlg = NumberingAddDialog(config_data=seed, on_save=on_save)
        dlg._save_to_disk()
        assert len(saved) == 1
        assert saved[0]["schemes"]["s1"]["name"] == "CB"
        dlg.close()

    def test_write_toml_file(self, tmp_path: Path) -> None:
        """update_file_sections writes numbering rules and preserves comments."""
        from docwen_runtime.config.loader import ConfigLoader

        config_dir = tmp_path / "configs"
        config_dir.mkdir()
        _write_minimal_base_config_tree(config_dir)
        (config_dir / "numbering").mkdir(parents=True, exist_ok=True)
        (config_dir / "numbering" / "add.toml").write_text(
            '[settings]\norder = ["s1"]\ndefault_scheme = "s1"\n[schemes.s1]\nname = "F"\nis_system = false\n'
            '[schemes.s1.level_1]\nformat = "{1.arabic_half}."\n',
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)
        ok = loader.update_file_sections(
            "numbering/add.toml",
            {
                "settings": {"default_scheme": "s1", "order": ["s1"]},
                "schemes": {
                    "s1": {
                        "name": "F",
                        "is_system": False,
                        "level_1": {"format": "{1.arabic_half}."},
                    }
                },
            },
        )
        assert ok
        content = (config_dir / "numbering" / "add.toml").read_text(encoding="utf-8")
        assert "default_scheme" in content
        assert '"F"' in content or "F" in content

    def test_write_toml_file_preserves_unrelated_sections(self, tmp_path: Path) -> None:
        """update_file_sections preserves sections not listed in updates."""
        from docwen_runtime.config.loader import ConfigLoader

        config_dir = tmp_path / "configs"
        config_dir.mkdir()
        _write_minimal_base_config_tree(config_dir)
        (config_dir / "numbering").mkdir(parents=True, exist_ok=True)
        # Pre-write with unrelated sections.
        (config_dir / "numbering" / "add.toml").write_text(
            "[defaults.text]\nadd_numbering = true\n\n# keep me\n[extra]\nkeep = true\n",
            encoding="utf-8",
        )
        loader = ConfigLoader(base_dir=config_dir, user_dir=config_dir)
        ok = loader.update_file_sections(
            "numbering/add.toml",
            {
                "settings": {"default_scheme": "s1", "order": ["s1"]},
                "schemes": {
                    "s1": {
                        "name": "F",
                        "is_system": False,
                        "level_1": {"format": "{1.arabic_half}."},
                    }
                },
            },
        )
        assert ok
        content = (config_dir / "numbering" / "add.toml").read_text(encoding="utf-8")
        assert "# keep me" in content
        assert "[extra]" in content
        assert "keep = true" in content
