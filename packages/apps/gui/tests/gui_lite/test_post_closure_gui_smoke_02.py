"""Focused tests split from test_post_closure_gui_smoke.py."""

from __future__ import annotations

import pytest

from ._post_closure_gui_smoke_support import (
    Path,
)

pytestmark = pytest.mark.gui
from ._post_closure_gui_smoke_support import (
    window as window,
)


class TestTemplateSelectorLite:
    """Verify TemplateSelector widget creation and basic interaction."""

    def test_template_selector_creates_with_object_name(self, qapp):
        from docwen_gui.widgets.template_selector import TemplateSelector

        selector = TemplateSelector(template_type="docx")
        assert selector.objectName() == "templateSelectorRoot"

    def test_template_selector_empty_state_visible(self, qapp):
        from PySide6.QtWidgets import QWidget

        from docwen_gui.widgets.template_selector import TemplateSelector

        selector = TemplateSelector(template_type="docx")
        empty = selector.findChild(QWidget, "templateSelectorEmptyState")
        assert empty is not None
        # QWidget.isVisible() requires the entire parent chain to be visible;
        # verify the empty state is not hidden at the widget level instead.
        assert not empty.isHidden()

    def test_template_selector_populate_and_select(self, qapp, qtbot):
        from docwen_gui.widgets.template_selector import TemplateSelector

        selector = TemplateSelector(template_type="docx")
        qtbot.addWidget(selector)
        names = ["Standard Report", "Academic Paper", "Business Letter"]
        selector.add_templates(names, auto_select_first=True)

        selected = selector.get_selected()
        assert selected is not None
        assert selected in names

        # List widget should have 3 items
        assert selector._list.count() == 3

    def test_template_selector_signal_emission(self, qapp, qtbot):
        from docwen_gui.widgets.template_selector import TemplateSelector

        selector = TemplateSelector(template_type="docx")
        qtbot.addWidget(selector)

        emitted: list[str] = []
        selector.template_selected.connect(lambda name: emitted.append(name))

        names = ["Custom Template"]
        selector.add_templates(names, auto_select_first=True)

        assert len(emitted) >= 1
        assert emitted[0] == "Custom Template"

    def test_tabbed_template_selector_creates(self, qapp):
        from docwen_gui.widgets.template_selector_tabbed import TabbedTemplateSelector

        tabbed = TabbedTemplateSelector()
        assert tabbed.objectName() == "tabbedTemplateSelector"

    def test_tabbed_template_selector_has_two_tabs(self, qapp):
        from docwen_gui.widgets.template_selector_tabbed import TabbedTemplateSelector

        tabbed = TabbedTemplateSelector()
        assert "docx" in tabbed._selectors
        assert "xlsx" in tabbed._selectors
        assert tabbed._stack.count() == 2


class TestNumberingEditorsLite:
    """Verify NumberingAddDialog and NumberingCleanDialog creation."""

    def test_numbering_add_dialog_creates_with_object_name(self, qapp, window):
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(parent=window, config_data={})
        assert dlg.objectName() == "numberingAddDialog"
        dlg.close()

    def test_numbering_add_dialog_has_scheme_list(self, qapp, window):
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        dlg = NumberingAddDialog(parent=window, config_data={})
        assert dlg.scheme_list is not None
        dlg.close()

    def test_numbering_add_dialog_with_seeded_data(self, qapp, window):
        from docwen_gui.widgets.settings.numbering_add_editor import (
            NumberingAddDialog,
        )

        seed = {
            "settings": {"default_scheme": "test", "order": ["test"]},
            "schemes": {
                "test": {
                    "name": "Test Scheme",
                    "enabled": True,
                    "is_system": False,
                    "level_1": {"format": "{1.chinese_lower}、"},
                }
            },
        }
        dlg = NumberingAddDialog(parent=window, config_data=seed)
        assert "test" in dlg.schemes
        assert dlg.scheme_list.count() >= 1
        dlg.close()

    def test_numbering_clean_dialog_creates_with_object_name(self, qapp, window):
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        dlg = NumberingCleanDialog(parent=window, config_data={})
        assert dlg.objectName() == "numberingCleanDialog"
        dlg.close()

    def test_numbering_clean_dialog_has_rule_list(self, qapp, window):
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        dlg = NumberingCleanDialog(parent=window, config_data={})
        assert dlg.rule_list is not None
        dlg.close()

    def test_numbering_clean_dialog_with_seeded_data(self, qapp, window):
        from docwen_gui.widgets.settings.numbering_clean_editor import (
            NumberingCleanDialog,
        )

        seed = {
            "settings": {"order": ["clean_test"]},
            "rules": [
                {
                    "id": "clean_test",
                    "enabled": True,
                    "pattern": r"^\s*",
                    "description": "Test Rule",
                    "level": 1,
                }
            ],
        }
        dlg = NumberingCleanDialog(parent=window, config_data=seed)
        assert "clean_test" in dlg.rules
        assert dlg.rule_list.count() >= 1
        dlg.close()


class TestProofreadEditorDialogsLite:
    """Verify proofread editor dialogs have objectNames."""

    def test_symbol_mapping_editor_has_object_name(self, qapp, window, tmp_path):
        from docwen_gui.widgets.settings.proofread_tab import (
            _SymbolMappingEditor,
        )

        dlg = _SymbolMappingEditor(str(tmp_path / "proofread_pairing.toml"), parent=window)
        assert dlg.objectName() == "symbolMappingEditor"
        dlg.close()

    def test_typos_dictionary_editor_has_object_name(self, qapp, window, tmp_path):
        from docwen_gui.widgets.settings.proofread_tab import (
            _TyposDictionaryEditor,
        )

        dlg = _TyposDictionaryEditor(str(tmp_path / "proofread_typos.toml"), parent=window)
        assert dlg.objectName() == "typosDictionaryEditor"
        dlg.close()

    def test_sensitive_word_editor_has_object_name(self, qapp, window, tmp_path):
        from docwen_gui.widgets.settings.proofread_tab import (
            _SensitiveWordEditor,
        )

        dlg = _SensitiveWordEditor(str(tmp_path / "proofread_sensitive.toml"), parent=window)
        assert dlg.objectName() == "sensitiveWordEditor"
        dlg.close()


class TestTomlEditorLite:
    """Verify TOML editor widgets have objectNames."""

    def test_toml_editor_widget_has_object_name(self, qapp, window):
        from pathlib import Path

        from docwen_gui.widgets.settings.toml_editor import TomlEditorWidget

        def _resolver(name: str) -> Path:
            return Path("/tmp/test.toml")

        editor = TomlEditorWidget(
            window,
            config_name="test",
            path_resolver=_resolver,
        )
        assert editor.objectName() == "tomlEditorWidget"

    def test_toml_editor_dialog_has_object_name(self, qapp, window):
        from pathlib import Path

        from docwen_gui.widgets.settings.toml_editor import (
            TomlEditorDialog,
            TomlEditorWidget,
        )

        def _resolver(name: str) -> Path:
            return Path("/tmp/test.toml")

        widget = TomlEditorWidget(
            window,
            config_name="test",
            path_resolver=_resolver,
        )
        dlg = TomlEditorDialog(window, editor=widget)
        assert dlg.objectName() == "tomlEditorDialog"
