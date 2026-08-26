"""Settings integration for Runtime-owned optimization resources."""

from __future__ import annotations

from copy import deepcopy
from typing import Any, cast

import pytest
from PySide6.QtWidgets import QComboBox
from tests.support.gui_vm_fakes import FakeController, optimization_capability_projection

from docwen_gui.view_models.settings_vm import SettingsViewModel

pytestmark = pytest.mark.gui


def _vm(controller: object | None) -> SettingsViewModel:
    vm = SettingsViewModel(controller=None)
    vm._controller = cast(Any, controller)  # pyright: ignore[reportPrivateUsage]
    return vm


def _empty_projection() -> dict[str, object]:
    projection = deepcopy(optimization_capability_projection())
    optimizations = projection["optimizations"]
    assert isinstance(optimizations, dict)
    optimizations["resources"] = []
    optimizations["counts"] = {
        "resources": 0,
        "available_resources": 0,
        "unavailable_resources": 0,
        "bindings": 0,
        "available_bindings": 0,
        "unavailable_bindings": 0,
    }
    return projection


class TestOptimizationDiscovery:
    def test_missing_controller_is_failed_not_ready_empty(self) -> None:
        result = _vm(None).get_optimization_choices_result(source_category="document")

        assert result.status == "failed"
        assert result.choices == ()
        assert result.error is not None

    def test_valid_empty_catalog_is_ready(self) -> None:
        controller = FakeController()
        controller.describe_runtime_capabilities = _empty_projection  # type: ignore[method-assign]

        result = _vm(controller).get_optimization_choices_result(source_category="document")

        assert result.status == "ready"
        assert result.choices == ()
        assert result.error is None

    def test_malformed_catalog_fails_closed(self) -> None:
        controller = FakeController()
        controller.describe_runtime_capabilities = lambda: {"broken": True}  # type: ignore[method-assign]

        result = _vm(controller).get_optimization_choices_result(source_category="document")

        assert result.status == "failed"
        assert result.choices == ()

    def test_category_binding_selects_document_image_and_layout_resources(self) -> None:
        vm = _vm(FakeController())

        document = vm.get_optimization_choices_result(source_category="document").items
        image = vm.get_optimization_choices_result(source_category="image").items
        layout = vm.get_optimization_choices_result(source_category="layout").items

        assert list(document) == ["gongwen"]
        assert list(image) == ["invoice_cn"]
        assert list(layout) == ["invoice_cn"]

    def test_config_only_orders_and_disables_known_runtime_ids(self) -> None:
        controller = FakeController(
            {
                "optimize": {
                    "settings": {"order": ["ghost", "invoice_cn", "gongwen"]},
                    "types": {"gongwen": {"enabled": False}, "ghost": {"enabled": True}},
                }
            }
        )

        choices = _vm(controller).get_optimization_choices_result().items

        assert list(choices) == ["invoice_cn"]
        assert "ghost" not in choices

    def test_locale_does_not_hide_runtime_capabilities(self) -> None:
        vm = _vm(FakeController())

        assert list(vm.get_optimization_choices_result(locale="en_US").items) == ["gongwen", "invoice_cn"]
        assert list(vm.get_optimization_choices_result(locale="zh_CN").items) == ["gongwen", "invoice_cn"]


def _combo(dialog: Any, tab_key: str) -> QComboBox:
    tab = dialog._tabs[tab_key]
    combo = tab.widgets["to_md_optimization_type"]  # type: ignore[attr-defined]
    assert isinstance(combo, QComboBox)
    return combo


class TestDialogOptimizationPopulation:
    def test_three_category_combos_have_only_runtime_resources(self, qapp) -> None:
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        dialog = SettingsDialog(view_model=_vm(FakeController()))
        try:
            assert [_combo(dialog, "document").itemData(i) for i in range(_combo(dialog, "document").count())] == [
                "gongwen"
            ]
            assert [_combo(dialog, "image").itemData(i) for i in range(_combo(dialog, "image").count())] == [
                "invoice_cn"
            ]
            assert [_combo(dialog, "layout").itemData(i) for i in range(_combo(dialog, "layout").count())] == [
                "invoice_cn"
            ]
            assert "contract" not in {
                _combo(dialog, "document").itemData(i) for i in range(_combo(dialog, "document").count())
            }
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_ready_empty_is_blank_without_error_placeholder(self, qapp) -> None:
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        controller = FakeController()
        controller.describe_runtime_capabilities = _empty_projection  # type: ignore[method-assign]
        dialog = SettingsDialog(view_model=_vm(controller))
        try:
            assert _combo(dialog, "document").count() == 0
            assert not _combo(dialog, "document").isEnabled()
        finally:
            dialog.close()
            dialog.deleteLater()

    def test_failure_is_contained_inside_combo_and_page_still_loads(self, qapp) -> None:
        from docwen_gui.widgets.settings.dialog import SettingsDialog

        controller = FakeController()
        controller.describe_runtime_capabilities = lambda: {"broken": True}  # type: ignore[method-assign]
        dialog = SettingsDialog(view_model=_vm(controller))
        try:
            combo = _combo(dialog, "document")
            assert combo.count() == 1
            assert combo.itemData(0) is None
            assert not combo.isEnabled()
            assert dialog._tabs["document"].objectName() != "settingsTabLoadErrorPage"
        finally:
            dialog.close()
            dialog.deleteLater()


class TestInitialTabKey:
    def test_default_and_updates(self) -> None:
        vm = SettingsViewModel()
        assert vm.initial_tab_key is None
        vm.set_initial_tab("layout")
        assert vm.initial_tab_key == "layout"

    def test_dialog_activates_requested_tab(self, qapp) -> None:
        from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

        vm = SettingsViewModel()
        vm.set_initial_tab("document")
        dialog = SettingsDialog(view_model=vm)
        try:
            assert TAB_KEYS[dialog._tab_widget.currentIndex()] == "document"
        finally:
            dialog.close()
            dialog.deleteLater()


class TestDetermineOptimalSettingsTab:
    @staticmethod
    def _window(file_paths: list[str], contexts: dict[str, tuple[str, str]]) -> Any:
        from unittest.mock import MagicMock

        from docwen_gui.main_window import MainWindow

        window = MagicMock()
        batch_vm = MagicMock()
        batch_vm.get_files.return_value = file_paths
        batch_vm.get_file_display_category.side_effect = {
            path: "text" if category == "markdown" else category for path, (_format, category) in contexts.items()
        }.get
        window._batch_list_vm = batch_vm
        window._CATEGORY_TO_SETTINGS_TAB = {
            "text": "text",
            "document": "document",
            "spreadsheet": "spreadsheet",
            "image": "image",
            "layout": "layout",
        }
        window.determine = lambda: MainWindow._determine_optimal_settings_tab(window)
        return window

    @pytest.mark.parametrize(
        ("path", "context", "expected"),
        [
            ("a.docx", ("docx", "document"), "document"),
            ("b.png", ("png", "image"), "image"),
            ("c.md", ("markdown", "markdown"), "text"),
            ("d.xyz", ("xyz", "other"), None),
        ],
    )
    def test_single_category(self, path: str, context: tuple[str, str], expected: str | None) -> None:
        assert self._window([path], {path: context}).determine() == expected

    def test_majority_and_tie_breaking(self) -> None:
        contexts = {
            "a.docx": ("docx", "document"),
            "b.docx": ("docx", "document"),
            "c.png": ("png", "image"),
        }
        assert self._window(list(contexts), contexts).determine() == "document"
        assert self._window(["a.docx", "c.png"], contexts).determine() == "document"
