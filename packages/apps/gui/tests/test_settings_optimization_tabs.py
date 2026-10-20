"""Settings integration for Runtime-owned optimization resources."""

from __future__ import annotations

from copy import deepcopy
from typing import Any, cast

import pytest
from PySide6.QtWidgets import QLabel, QPushButton, QWidget
from tests.support.gui_vm_fakes import FakeController, optimization_capability_projection

from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.dialog import SettingsDialog
from docwen_gui.widgets.settings.gongwen_tab import GongwenTab

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


def _default_button(dialog: SettingsDialog, category: str) -> QPushButton:
    button = dialog.findChild(QPushButton, f"{category}OptimizationDefaultButton")
    assert button is not None
    return button


@pytest.fixture
def optimization_dialog(qapp):
    dialog = SettingsDialog(view_model=_vm(FakeController()))
    yield dialog
    dialog.view_model.cancel()
    dialog.close()
    dialog.deleteLater()


class TestSpecializedOptimizationPages:
    def test_pages_own_defaults_and_ordinary_pages_have_no_optimization_fields(self, optimization_dialog) -> None:
        dialog = optimization_dialog
        assert dialog.activate_section("gongwen")
        assert dialog.findChild(QWidget, "gongwenToMarkdownCard") is not None
        assert dialog.findChild(QWidget, "gongwenFromMarkdownCard") is not None
        assert dialog.activate_section("invoice")
        assert dialog.findChild(QWidget, "invoiceLayoutToMarkdownCard") is not None
        assert dialog.findChild(QWidget, "invoiceImageToMarkdownCard") is not None
        for key in ("document", "image", "layout"):
            assert not any("optimization" in field for field in dialog._tabs[key].widgets)
        assert dialog._tabs["text"].findChild(QWidget, "textGongwenFieldsCard") is None

    def test_document_default_can_replace_another_choice_and_clear_only_this_choice(self, optimization_dialog) -> None:
        from docwen_gui.i18n import t

        dialog = optimization_dialog
        vm = dialog.view_model
        button = _default_button(dialog, "document")
        vm.set_conversion_default("document", "to_md_optimization", "another_optimizer")
        current = dialog.findChild(QLabel, "documentOptimizationCurrent")
        assert current is not None and "another_optimizer" in current.text()
        assert button.text() == t("settings.optimization.set_default")
        button.click()
        assert vm.get_default_optimization("document") == "gongwen"
        assert button.text() == t("settings.optimization.clear_default")
        assert "another_optimizer" not in current.text()
        button.click()
        assert vm.get_default_optimization("document") == "none"
        assert button.text() == t("settings.optimization.set_default")

    def test_invoice_scopes_and_config_reload_stay_in_sync(self, optimization_dialog) -> None:
        from docwen_gui.i18n import t

        dialog = optimization_dialog
        vm = dialog.view_model
        layout_button = _default_button(dialog, "layout")
        image_button = _default_button(dialog, "image")
        layout_button.click()
        assert vm.get_default_optimization("layout") == "invoice_cn"
        assert vm.get_default_optimization("image") == "none"
        image_button.click()
        assert vm.get_default_optimization("image") == "invoice_cn"
        assert vm.get_default_optimization("document") == "none"
        vm.set_conversion_default("image", "to_md_optimization", "none")
        assert image_button.text() == t("settings.optimization.set_default")
        assert layout_button.text() == t("settings.optimization.clear_default")
        vm.cancel()
        assert vm.get_default_optimization("layout") == "none"
        assert layout_button.text() == t("settings.optimization.set_default")
        config = vm.config
        config.conversion_defaults.image["to_md_optimization"] = "invoice_cn"
        vm.load_full_config(config)
        assert image_button.text() == t("settings.optimization.clear_default")

    def test_ready_empty_has_localized_reason_without_discovery_failure(self, qapp) -> None:
        from docwen_gui.i18n import t

        controller = FakeController()
        controller.describe_runtime_capabilities = _empty_projection  # type: ignore[method-assign]
        dialog = SettingsDialog(view_model=_vm(controller))
        try:
            reason = dialog.findChild(QLabel, "documentOptimizationAvailability")
            assert reason is not None
            assert reason.text() == t("settings.optimization.unavailable", name=t("settings.gongwen.title"))
            assert reason.text() != t("settings.optimization.discovery_failed")
            assert not _default_button(dialog, "document").isEnabled()
        finally:
            dialog.close()
            dialog.deleteLater()

    @pytest.mark.parametrize("failed", [False, True])
    def test_saved_unavailable_default_remains_clearable(self, qapp, failed) -> None:
        from docwen_gui.i18n import t

        controller = FakeController()
        controller.describe_runtime_capabilities = (  # type: ignore[method-assign]
            (lambda: {"broken": True}) if failed else _empty_projection
        )
        vm = _vm(controller)
        vm.set_conversion_default("layout", "to_md_optimization", "invoice_cn")
        dialog = SettingsDialog(view_model=vm)
        try:
            button = _default_button(dialog, "layout")
            reason = dialog.findChild(QLabel, "layoutOptimizationAvailability")
            assert reason is not None and reason.text()
            if failed:
                assert reason.text() == t("settings.optimization.discovery_failed")
            assert button.text() == t("settings.optimization.clear_default")
            assert button.isEnabled()
            button.click()
            assert vm.get_default_optimization("layout") == "none"
            assert not button.isEnabled()
            assert dialog._tabs["invoice"].objectName() != "settingsTabLoadErrorPage"
            gongwen_tab = dialog._tabs["gongwen"]
            assert isinstance(gongwen_tab, GongwenTab)
            assert gongwen_tab._gongwen_fields.isEnabled()
        finally:
            vm.cancel()
            dialog.close()
            dialog.deleteLater()

    @pytest.mark.parametrize(
        "page,button_name,destination",
        [
            ("gongwen", "gongwenDocumentSettingsButton", "document"),
            ("invoice", "invoiceLayoutSettingsButton", "layout"),
            ("invoice", "invoiceOcrSettingsButton", "export"),
        ],
    )
    def test_shared_settings_links_navigate_without_editing(self, optimization_dialog, page, button_name, destination):
        dialog = optimization_dialog
        assert dialog.activate_section(page)
        button = dialog.findChild(QPushButton, button_name)
        assert button is not None
        button.click()
        assert dialog.current_section() == destination
        assert not dialog.view_model.is_dirty


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
