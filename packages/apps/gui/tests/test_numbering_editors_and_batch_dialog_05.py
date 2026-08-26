"""Focused tests split from test_numbering_editors_and_batch_dialog.py."""

from __future__ import annotations

from ._numbering_editors_and_batch_dialog_support import (
    QApplication,
    _patch_all_modals,
    pytest,
)

pytestmark = pytest.mark.gui


class TestNumberingEditorsUserPath:
    """End-to-end user path: VM → dialog → edit → save → VM → reopen.

    These tests verify that the dialog→ViewModel→dialog round-trip works
    without monkey-patching the save path.  Modal confirmations for
    close/cancel are still suppressed to allow automated testing.
    """

    def test_text_tab_editor_projection_degrades_malformed_dynamic_drafts(
        self, qapp: QApplication, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        from docwen_gui.models.settings_config import SettingsConfig
        from docwen_gui.view_models.settings_vm import SECTION_TEXT, SettingsViewModel
        from docwen_gui.widgets.settings import numbering_add_editor, numbering_clean_editor
        from docwen_gui.widgets.settings.text_tab import TextTab

        vm = SettingsViewModel(config=SettingsConfig())
        vm.set_field(SECTION_TEXT, "numbering_schemes", ["invalid dynamic value"])
        vm.set_field(SECTION_TEXT, "numbering_clean_rules", "invalid dynamic value")
        captured: dict[str, dict] = {}

        class _EditorStub:
            kind = ""

            def __init__(self, _parent, *, config_data: dict, on_save) -> None:
                captured[self.kind] = config_data

            def exec(self) -> int:
                return 0

        class _AddEditorStub(_EditorStub):
            kind = "add"

        class _CleanEditorStub(_EditorStub):
            kind = "clean"

        monkeypatch.setattr(numbering_add_editor, "NumberingAddDialog", _AddEditorStub)
        monkeypatch.setattr(numbering_clean_editor, "NumberingCleanDialog", _CleanEditorStub)

        tab = TextTab(vm)
        tab._open_numbering_scheme_editor()
        tab._open_numbering_clean_editor()

        assert captured["add"] == {
            "number_styles": {},
            "schemes": {},
            "settings": {"default_scheme": "hierarchical_standard", "order": []},
        }
        assert captured["clean"] == {"settings": {"order": []}, "rules": []}

        tab._on_scheme_changed(0)
        assert tab._scheme_combo.count() == 0
        assert vm.config.text.numbering_schemes == ["invalid dynamic value"]
        tab.close()

    def test_add_editor_round_trip_preserves_data(self, qapp: QApplication) -> None:
        """Save data via dialog, verify VM holds it, reopen and see it."""
        from docwen_gui.models.settings_config import SettingsConfig, TextConfig
        from docwen_gui.view_models.settings_vm import SECTION_TEXT, SettingsViewModel
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        # 1. Create VM with seed data
        config = SettingsConfig()
        config.text = TextConfig(
            numbering_schemes={
                "settings": {"default_scheme": "my_scheme", "order": ["my_scheme"]},
                "schemes": {
                    "my_scheme": {
                        "name": "My Scheme",
                        "description": "Original desc",
                        "enabled": True,
                        "is_system": False,
                        "level_1": {"format": "{1.chinese_lower}、"},
                    }
                },
            }
        )
        vm = SettingsViewModel(config=config)

        # 2. Build current_data from VM state (same as text_tab does)
        cfg = vm.config
        ns = cfg.text.numbering_schemes
        current_data = {
            "number_styles": ns.get("number_styles", {}),
            "schemes": ns.get("schemes", {}),
            "settings": {
                "default_scheme": cfg.text.default_scheme,
                "order": list(ns.get("settings", {}).get("order", [])),
            },
        }

        # 3. Open dialog, edit, save via dialog
        dlg = NumberingAddDialog(config_data=current_data)
        _patch_all_modals(dlg, confirm_value=True)
        dlg._select_scheme("my_scheme")
        dlg.desc_edit.setText("Updated via user path")
        dlg._save_to_disk()

        # 4. Apply dialog data to VM (same as text_tab's if-accepted block)
        schemes_data = dlg.get_schemes_data()
        vm.set_field(SECTION_TEXT, "numbering_schemes", schemes_data)

        # 5. Verify VM has updated data
        cfg2 = vm.config
        saved_schemes = cfg2.text.numbering_schemes
        assert saved_schemes["schemes"]["my_scheme"]["description"] == "Updated via user path"
        assert saved_schemes["number_styles"] == {}

        # 6. Reopen dialog with VM data — must see saved changes
        ns2 = cfg2.text.numbering_schemes
        current_data2 = {
            "number_styles": ns2.get("number_styles", {}),
            "schemes": ns2.get("schemes", {}),
            "settings": {
                "default_scheme": cfg2.text.default_scheme,
                "order": list(ns2.get("settings", {}).get("order", [])),
            },
        }
        dlg2 = NumberingAddDialog(config_data=current_data2)
        assert "my_scheme" in dlg2.schemes
        assert dlg2.schemes["my_scheme"].description == "Updated via user path"
        dlg2.close()

    def test_clean_editor_round_trip_preserves_data(self, qapp: QApplication) -> None:
        """Save clean rules via dialog, verify VM holds them, reopen and see them."""
        from docwen_gui.models.settings_config import SettingsConfig, TextConfig
        from docwen_gui.view_models.settings_vm import SECTION_TEXT, SettingsViewModel
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        config = SettingsConfig()
        config.text = TextConfig(
            numbering_clean_rules={
                "settings": {"order": ["r1"]},
                "rules": [
                    {
                        "id": "r1",
                        "pattern": r"^[0-9]+",
                        "enabled": True,
                        "description": "My Rule",
                        "level": 1,
                    }
                ],
            }
        )
        vm = SettingsViewModel(config=config)

        cfg = vm.config
        cr = cfg.text.numbering_clean_rules
        current_data = {
            "rules": cr.get("rules", []),
            "settings": {
                "order": list(cr.get("settings", {}).get("order", [])),
            },
        }

        dlg = NumberingCleanDialog(config_data=current_data)
        _patch_all_modals(dlg, confirm_value=True)
        dlg._select_rule("r1")
        dlg.pattern_edit.setText(r"^[0-9]+\.")
        dlg._save()

        rules_data = dlg.get_rules_data()
        vm.set_field(SECTION_TEXT, "numbering_clean_rules", rules_data)

        cfg2 = vm.config
        saved_rules = cfg2.text.numbering_clean_rules
        assert saved_rules["rules"][0]["pattern"] == r"^[0-9]+\."

        cr2 = cfg2.text.numbering_clean_rules
        current_data2 = {
            "rules": cr2.get("rules", []),
            "settings": {
                "order": list(cr2.get("settings", {}).get("order", [])),
            },
        }
        dlg2 = NumberingCleanDialog(config_data=current_data2)
        assert "r1" in dlg2.rules
        assert dlg2.rules["r1"].pattern == r"^[0-9]+\."
        dlg2.close()

    def test_both_editors_use_separate_keys_no_interference(self, qapp: QApplication) -> None:
        """Verify numbering_schemes and numbering_clean_rules are independent."""
        from docwen_gui.models.settings_config import SettingsConfig, TextConfig
        from docwen_gui.view_models.settings_vm import SECTION_TEXT, SettingsViewModel
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        config = SettingsConfig()
        config.text = TextConfig(
            numbering_schemes={
                "settings": {"default_scheme": "s1", "order": ["s1"]},
                "schemes": {"s1": {"name": "AddScheme", "is_system": False}},
            },
            numbering_clean_rules={
                "settings": {"order": ["r1"]},
                "rules": [
                    {
                        "id": "r1",
                        "pattern": r"^x",
                        "enabled": True,
                        "description": "CleanRule",
                        "level": 1,
                    }
                ],
            },
        )
        vm = SettingsViewModel(config=config)

        # Save from add editor
        cfg = vm.config
        ns = cfg.text.numbering_schemes
        dlg_add = NumberingAddDialog(
            config_data={
                "number_styles": ns.get("number_styles", {}),
                "schemes": ns.get("schemes", {}),
                "settings": {
                    "default_scheme": cfg.text.default_scheme,
                    "order": list(ns.get("settings", {}).get("order", [])),
                },
            }
        )
        _patch_all_modals(dlg_add, confirm_value=True)
        dlg_add._select_scheme("s1")
        dlg_add.name_edit.setText("AddScheme Modified")
        dlg_add._save_to_disk()
        vm.set_field(SECTION_TEXT, "numbering_schemes", dlg_add.get_schemes_data())

        # Save from clean editor
        cfg = vm.config
        cr = cfg.text.numbering_clean_rules
        dlg_clean = NumberingCleanDialog(
            config_data={
                "rules": cr.get("rules", []),
                "settings": {
                    "order": list(cr.get("settings", {}).get("order", [])),
                },
            }
        )
        _patch_all_modals(dlg_clean, confirm_value=True)
        dlg_clean._select_rule("r1")
        dlg_clean.desc_edit.setText("CleanRule Modified")
        dlg_clean._save()

        vm.set_field(SECTION_TEXT, "numbering_clean_rules", dlg_clean.get_rules_data())

        # Verify both keys are independent
        final = vm.config
        add_data = final.text.numbering_schemes
        clean_data = final.text.numbering_clean_rules

        # Add editor data preserved
        assert add_data["schemes"]["s1"]["name"] == "AddScheme Modified"
        assert "rules" not in add_data  # no cross-contamination

        # Clean editor data preserved
        assert clean_data["rules"][0]["description"] == "CleanRule Modified"
        assert "schemes" not in clean_data  # no cross-contamination


class TestBatchAddFailedDialogUserPath:
    """Test that the batch add failed dialog is connected to the batch UI pipeline."""

    def test_batch_list_shows_dialog_on_failed_files(self, qapp: QApplication, monkeypatch: pytest.MonkeyPatch) -> None:
        """When files_failed comes through the signal, the dialog is shown."""
        from docwen_gui.view_models.batch_list_vm import BatchListViewModel
        from docwen_gui.widgets.batch_list import BatchList

        vm = BatchListViewModel()
        widget = BatchList(vm)

        # Capture the dialog call
        dialog_shown: list[list[tuple[str, str]]] = []

        def fake_show(parent, failed_files):
            dialog_shown.append(failed_files)

        monkeypatch.setattr(
            "docwen_gui.widgets.batch_dialogs.show_batch_add_failed_dialog",
            fake_show,
        )

        # Simulate files_added signal with failed entries
        added = ["/path/ok.md"]
        failed = [("/path/bad.xyz", "Unsupported file type")]
        vm.files_added.emit(added, failed)

        qapp.processEvents()
        assert len(dialog_shown) == 1
        assert len(dialog_shown[0]) == 1
        assert dialog_shown[0][0][0] == "/path/bad.xyz"

        widget.close()

    def test_batch_list_no_dialog_when_no_failures(self, qapp: QApplication, monkeypatch: pytest.MonkeyPatch) -> None:
        """No dialog when all files are added successfully."""
        from docwen_gui.view_models.batch_list_vm import BatchListViewModel
        from docwen_gui.widgets.batch_list import BatchList

        vm = BatchListViewModel()
        widget = BatchList(vm)

        dialog_shown: list[list[tuple[str, str]]] = []

        def fake_show(parent, failed_files):
            dialog_shown.append(failed_files)

        monkeypatch.setattr(
            "docwen_gui.widgets.batch_dialogs.show_batch_add_failed_dialog",
            fake_show,
        )

        added = ["/path/ok.md"]
        vm.files_added.emit(added, [])

        qapp.processEvents()
        assert len(dialog_shown) == 0

        widget.close()


class TestTextTabEditorSaveFailure:
    """When persistence returns False, memory must NOT be updated and error must be shown."""

    def test_schemes_save_failure_does_not_update_memory(self, qapp: QApplication) -> None:
        """Verify memory is NOT updated when persist_numbering_schemes_source fails."""
        from docwen_gui.i18n import t
        from docwen_gui.models.settings_config import SettingsConfig
        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.text_tab import TextTab

        vm = SettingsViewModel(config=SettingsConfig())
        tab = TextTab(vm)

        # Capture original state
        original_schemes = vm.config.text.numbering_schemes

        schemes_data = {
            "settings": {"default_scheme": "new_scheme", "order": ["new_scheme"]},
            "schemes": {"new_scheme": {"name": "New", "is_system": False}},
        }

        # Mock persist to return False
        persist_called_with: list[dict[str, object]] = []

        def failing_persist(data: dict[str, object]) -> bool:
            persist_called_with.append(data)
            return False

        vm.persist_numbering_schemes_source = failing_persist  # type: ignore[method-assign]

        # Mock QMessageBox.warning to avoid actual dialog
        warning_shown = []

        def fake_warning(parent, title, message):
            warning_shown.append((title, message))
            return None  # QMessageBox.warning returns ButtonRole, not needed

        from PySide6.QtWidgets import QMessageBox

        original_warning = QMessageBox.warning
        QMessageBox.warning = fake_warning  # type: ignore[method-assign]

        try:
            tab._on_numbering_schemes_saved(schemes_data)

            # Verify persist was called
            assert len(persist_called_with) == 1

            # Verify warning was shown
            assert len(warning_shown) == 1
            assert warning_shown[0] == (
                t("common.save_failed", "Save Failed"),
                t(
                    "settings.text.save_numbering_schemes_failed",
                    "Failed to save numbering schemes to disk. Changes were not persisted.",
                ),
            )

            # Verify memory was NOT updated
            assert vm.config.text.numbering_schemes == original_schemes
        finally:
            QMessageBox.warning = original_warning  # type: ignore[method-assign]

    def test_clean_rules_save_failure_does_not_update_memory(self, qapp: QApplication) -> None:
        """Verify memory is NOT updated when persist_numbering_clean_rules_source fails."""
        from docwen_gui.i18n import t
        from docwen_gui.models.settings_config import SettingsConfig
        from docwen_gui.view_models.settings_vm import SettingsViewModel
        from docwen_gui.widgets.settings.text_tab import TextTab

        vm = SettingsViewModel(config=SettingsConfig())
        tab = TextTab(vm)

        # Capture original state
        original_rules = vm.config.text.numbering_clean_rules

        rules_data = {
            "settings": {"order": ["r1"]},
            "rules": [{"id": "r1", "pattern": "^test", "enabled": True}],
        }

        persist_called_with: list[dict[str, object]] = []

        def failing_persist(data: dict[str, object]) -> bool:
            persist_called_with.append(data)
            return False

        vm.persist_numbering_clean_rules_source = failing_persist  # type: ignore[method-assign]

        warning_shown = []

        def fake_warning(parent, title, message):
            warning_shown.append((title, message))

        from PySide6.QtWidgets import QMessageBox

        original_warning = QMessageBox.warning
        QMessageBox.warning = fake_warning  # type: ignore[method-assign]

        try:
            tab._on_numbering_clean_rules_saved(rules_data)

            # Verify persist was called
            assert len(persist_called_with) == 1

            # Verify warning was shown
            assert len(warning_shown) == 1
            assert warning_shown[0] == (
                t("common.save_failed", "Save Failed"),
                t(
                    "settings.text.save_numbering_clean_rules_failed",
                    "Failed to save numbering clean rules to disk. Changes were not persisted.",
                ),
            )

            # Verify memory was NOT updated
            assert vm.config.text.numbering_clean_rules == original_rules
        finally:
            QMessageBox.warning = original_warning  # type: ignore[method-assign]
