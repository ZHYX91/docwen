"""Focused tests split from test_numbering_editors_and_batch_dialog.py."""

from __future__ import annotations

from ._numbering_editors_and_batch_dialog_support import (
    QApplication,
    _patch_all_modals,
    _patch_dialog_notify,
    pytest,
)

pytestmark = pytest.mark.gui


class TestNumberingSchemeDataContract:
    """Permanent localization and fallback contract for numbering schemes."""

    def test_display_name_translation_and_fallback(self, monkeypatch: pytest.MonkeyPatch) -> None:
        from docwen_gui.widgets.settings import numbering_add_editor

        translations = {"editors.numbering_add.names.localized": "Localized name"}

        def fake_t(key: str) -> str:
            return translations.get(key, f"[{key}]")

        monkeypatch.setattr(numbering_add_editor, "t", fake_t)

        assert (
            numbering_add_editor.NumberingScheme(
                "translated",
                name="Explicit name",
                name_key="localized",
            ).display_name()
            == "Localized name"
        )
        assert (
            numbering_add_editor.NumberingScheme(
                "explicit",
                name="Explicit name",
                name_key="missing",
            ).display_name()
            == "Explicit name"
        )
        assert (
            numbering_add_editor.NumberingScheme(
                "stable-id",
                name_key="missing",
            ).display_name()
            == "stable-id"
        )

    def test_display_description_translation_and_fallback(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_gui.widgets.settings import numbering_add_editor

        translations = {"editors.numbering_add.descriptions.localized": "Localized description"}

        def fake_t(key: str) -> str:
            return translations.get(key, f"[{key}]")

        monkeypatch.setattr(numbering_add_editor, "t", fake_t)

        assert (
            numbering_add_editor.NumberingScheme(
                "translated",
                description="Explicit description",
                description_key="localized",
            ).display_description()
            == "Localized description"
        )
        assert (
            numbering_add_editor.NumberingScheme(
                "explicit",
                description="Explicit description",
                description_key="missing",
            ).display_description()
            == "Explicit description"
        )
        assert (
            numbering_add_editor.NumberingScheme(
                "empty",
                description_key="missing",
            ).display_description()
            == ""
        )


class TestNumberingCleanDialogConstruction:
    """Test clean dialog creation and initial state."""

    def test_create_with_default_data(self, qapp: QApplication) -> None:
        from docwen_gui.i18n import t
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        dlg = NumberingCleanDialog(config_data={})
        assert dlg.windowTitle() == t(
            "editors.numbering_clean.window_title",
            "Numbering Removal Rule Editor",
        )
        assert dlg.rule_list is not None
        dlg.close()

    def test_malformed_nested_data_degrades_to_empty_rules(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import NumberingCleanDialog

        dlg = NumberingCleanDialog(config_data={"settings": "invalid", "rules": 42})

        assert dlg.order == []
        assert dlg.rules == {}
        dlg.close()

    def test_legacy_mapping_schema_is_not_accepted(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import NumberingCleanDialog

        dlg = NumberingCleanDialog(
            config_data={
                "settings": {"order": ["r1"]},
                "rules": {"r1": {"regex": "^legacy", "name": "Legacy"}},
            }
        )

        assert dlg.order == []
        assert dlg.rules == {}
        dlg.close()

    def test_create_with_seeded_data(self, qapp: QApplication) -> None:
        from docwen_gui.i18n import t
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [
                {
                    "id": "r1",
                    "pattern": r"^\s*[0-9]+",
                    "description": "A test rule",
                    "enabled": True,
                    "level": 1,
                }
            ],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        assert dlg.windowTitle() == t(
            "editors.numbering_clean.window_title",
            "Numbering Removal Rule Editor",
        )
        assert "r1" in dlg.rules
        rule = dlg.rules["r1"]
        assert rule.description == "A test rule"
        assert rule.pattern == r"^\s*[0-9]+"
        dlg.close()

    def test_rule_list_uses_cleanup_rule_localization_and_localized_tooltip(
        self, qapp: QApplication, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        from docwen_gui.widgets.settings import numbering_clean_editor

        translations = {
            "editors.numbering_clean.names.number_separator": "Localized rule",
            "editors.numbering_clean.descriptions.number_separator_desc": "Localized description",
        }

        def fake_t(key: str, default: str = "", **_kwargs: object) -> str:
            return translations.get(key, default or key)

        monkeypatch.setattr(numbering_clean_editor, "t", fake_t)
        dlg = numbering_clean_editor.NumberingCleanDialog(
            config_data={
                "settings": {"order": ["arabic_separator"]},
                "rules": [
                    {
                        "id": "arabic_separator",
                        "name_key": "number_separator",
                        "description_key": "number_separator_desc",
                        "description": "非当前界面的回退文案",
                        "pattern": r"^[0-9]+[.]",
                        "enabled": False,
                    }
                ],
            }
        )

        item = dlg.rule_list.item(0)
        assert item.text() == "Localized rule [off]"
        assert item.toolTip() == "Localized description\n^[0-9]+[.]"
        assert dlg.desc_edit.text() == "Localized description"
        dlg.close()

    def test_rule_list_falls_back_to_explicit_name_then_stable_id(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import NumberingCleanDialog

        dlg = NumberingCleanDialog(
            config_data={
                "settings": {"order": ["named", "id_only"]},
                "rules": [
                    {"id": "named", "name": "Readable name", "pattern": "^a"},
                    {"id": "id_only", "pattern": "^b"},
                ],
            }
        )

        assert [dlg.rule_list.item(index).text() for index in range(2)] == ["Readable name", "id_only"]
        dlg.close()

    def test_editing_description_replaces_localized_description_key(self, qapp: QApplication) -> None:
        from PySide6.QtTest import QTest

        from docwen_gui.widgets.settings.numbering_clean_editor import NumberingCleanDialog

        dlg = NumberingCleanDialog(
            config_data={
                "settings": {"order": ["r1"]},
                "rules": [
                    {
                        "id": "r1",
                        "description": "Original",
                        "description_key": "number_separator_desc",
                        "pattern": "^x",
                    }
                ],
            }
        )

        dlg.desc_edit.clear()
        QTest.keyClicks(dlg.desc_edit, "Alpha Beta")
        assert dlg.desc_edit.text() == "Alpha Beta"
        assert dlg.rules["r1"].description == "Alpha Beta"
        assert dlg.rules["r1"].description_key == ""
        assert dlg.rule_list.item(0).toolTip() == "Alpha Beta\n^x"
        dlg.close()

    def test_window_title_shows_dirty(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [{"id": "r1", "enabled": True, "pattern": "^test"}],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        initial_title = dlg.windowTitle()
        assert "*" not in initial_title

        dlg.pattern_edit.setText("^modified")
        assert "*" in dlg.windowTitle()
        dlg.close()


class TestNumberingCleanRuleCRUD:
    """Test rule create/copy/delete operations (modals patched)."""

    def test_create_new_rule(self, qapp: QApplication) -> None:
        from docwen_gui.i18n import t
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        dlg = NumberingCleanDialog(config_data={})
        initial_count = len(dlg.rules)
        dlg._create_new_rule()
        assert len(dlg.rules) == initial_count + 1
        new_id = dlg.current_rule_id
        assert new_id is not None
        assert new_id in dlg.rules
        assert dlg.rules[new_id].name == t("editors.numbering_clean.new_rule", "New Rule")
        assert dlg.rule_list.currentItem().text() == t("editors.numbering_clean.new_rule", "New Rule")
        dlg.close()

    def test_copy_rule_freezes_localized_text_and_clears_translation_keys(
        self, qapp: QApplication, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        from docwen_gui.widgets.settings import numbering_clean_editor

        translations = {
            "editors.numbering_clean.names.number_separator": "Localized rule",
            "editors.numbering_clean.descriptions.number_separator_desc": "Localized description",
            "editors.numbering_clean.copy_suffix": "Localized copy",
        }

        def fake_t(key: str, default: str = "", **_kwargs: object) -> str:
            return translations.get(key, default or key)

        monkeypatch.setattr(numbering_clean_editor, "t", fake_t)

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [
                {
                    "id": "r1",
                    "name_key": "number_separator",
                    "description_key": "number_separator_desc",
                    "description": "Fallback",
                    "pattern": "^test",
                    "enabled": True,
                    "is_system": True,
                }
            ],
        }
        dlg = numbering_clean_editor.NumberingCleanDialog(config_data=seed)
        _patch_dialog_notify(dlg)
        assert dlg.current_rule_id == "r1"
        dlg._copy_selected_rule()
        assert len(dlg.rules) == 2
        copied = dlg.rules["r1_copy"]
        assert copied.name == "Localized rule Localized copy"
        assert copied.name_key == ""
        assert copied.description == "Localized description"
        assert copied.description_key == ""
        assert copied.is_system is False
        assert dlg.rule_list.item(1).text() == "Localized rule Localized copy"
        serialized = dlg.get_rules_data()["rules"][1]
        assert serialized["name"] == "Localized rule Localized copy"
        assert serialized["name_key"] == ""
        assert serialized["description_key"] == ""
        dlg.close()

    def test_delete_custom_rule(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1", "r2"]},
            "rules": [
                {"id": "r1", "pattern": "^a", "enabled": True},
                {"id": "r2", "pattern": "^b", "enabled": True},
            ],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        _patch_all_modals(dlg, confirm_value=True)
        dlg._select_rule("r2")
        assert dlg.current_rule_id == "r2"
        dlg._delete_selected_rule()
        assert "r2" not in dlg.rules
        dlg.close()

    def test_cannot_delete_system_rule(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [{"id": "r1", "pattern": "^a", "enabled": True, "is_system": True}],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        _patch_all_modals(dlg, confirm_value=True)
        dlg._select_rule("r1")
        assert dlg.delete_btn.isEnabled() is False
        dlg._delete_selected_rule()
        assert "r1" in dlg.rules
        dlg.close()

    def test_move_rule_up_down(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1", "r2"]},
            "rules": [
                {"id": "r1", "pattern": "^a", "enabled": True},
                {"id": "r2", "pattern": "^b", "enabled": True},
            ],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        dlg._select_rule("r2")
        dlg._move_selected_rule_up()
        assert dlg.order[0] == "r2"
        assert dlg.order[1] == "r1"
        dlg._move_selected_rule_down()
        assert dlg.order[0] == "r1"
        assert dlg.order[1] == "r2"
        dlg.close()

    def test_enable_disable_toggle(self, qapp: QApplication) -> None:
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [{"id": "r1", "pattern": "^test", "enabled": True}],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        dlg._select_rule("r1")
        assert dlg.rules["r1"].enabled is True
        assert dlg.enabled_check.isChecked() is True

        dlg.enabled_check.setChecked(False)
        assert dlg.rules["r1"].enabled is False
        assert dlg._dirty

        dlg.enabled_check.setChecked(True)
        assert dlg.rules["r1"].enabled is True
        dlg.close()


class TestNumberingCleanRegexTest:
    """Test the live regex test playground (no modals)."""

    def test_regex_match_success(self, qapp: QApplication) -> None:
        from docwen_gui.i18n import t
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [
                {
                    "id": "r1",
                    "pattern": r"^[0-9]+",
                    "enabled": True,
                }
            ],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        dlg.test_input.setText("123abc")
        dlg._run_regex_test()
        status = dlg.status_label.text()
        assert status == t("editors.numbering_clean.match_success", "Match Success")
        result = dlg.test_result.toPlainText()
        assert result == "abc"  # digits stripped
        dlg.close()

    def test_regex_no_match(self, qapp: QApplication) -> None:
        from docwen_gui.i18n import t
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [
                {
                    "id": "r1",
                    "pattern": r"^[0-9]+",
                    "enabled": True,
                }
            ],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        dlg.test_input.setText("abc123")
        dlg._run_regex_test()
        status = dlg.status_label.text()
        assert status == t("editors.numbering_clean.no_match", "No Match")
        dlg.close()

    def test_regex_invalid_syntax(self, qapp: QApplication) -> None:
        from docwen_gui.i18n import t
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["r1"]},
            "rules": [
                {
                    "id": "r1",
                    "pattern": r"[unclosed",
                    "enabled": True,
                }
            ],
        }
        dlg = NumberingCleanDialog(config_data=seed)
        dlg.test_input.setText("test")
        dlg._run_regex_test()
        status = dlg.status_label.text()
        assert status.startswith(f"{t('editors.numbering_clean.regex_error', 'Regex Error')}:")
        dlg.close()
