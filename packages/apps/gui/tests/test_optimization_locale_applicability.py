"""Interface language, resource ownership, navigation and saved-default boundaries."""

from __future__ import annotations

from pathlib import Path
from typing import Any, cast

import pytest
from PySide6.QtWidgets import QPushButton
from tests.support.gui_vm_fakes import FakeController, FakeMainWindowViewModel, optimization_capability_projection

from docwen_application.controller import ApplicationController
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_gui.i18n import get_locale, set_locale
from docwen_gui.view_models.action_area_vm import ActionAreaViewModel
from docwen_gui.view_models.settings_vm import SettingsViewModel
from docwen_gui.widgets.settings.dialog import SettingsDialog
from docwen_runtime.config import ConfigLoader

pytestmark = pytest.mark.gui
LOCALES = ("zh_CN", "zh_TW", "en_US", "de_DE", "fr_FR", "es_ES", "pt_BR", "ru_RU", "ja_JP", "ko_KR", "vi_VN", "ar_SA")


@pytest.fixture(autouse=True)
def restore_locale():
    original = get_locale()
    yield
    set_locale(original)


def _vm(controller: object) -> SettingsViewModel:
    view_model = SettingsViewModel(controller=None)
    view_model._controller = cast(Any, controller)  # pyright: ignore[reportPrivateUsage]
    return view_model


@pytest.mark.parametrize("locale", LOCALES)
def test_only_declared_locale_gets_specialized_pages_without_an_open_file(qapp, locale: str) -> None:
    set_locale(locale)
    dialog = SettingsDialog(view_model=_vm(FakeController()))
    try:
        expected = locale == "zh_CN"
        assert dialog.activate_section("gongwen") is expected
        assert dialog.activate_section("invoice") is expected
        assert ("gongwen" in dialog.available_page_keys) is expected
        assert ("invoice" in dialog.available_page_keys) is expected
        assert dialog.activate_section("text")
        reset = dialog.findChild(QPushButton, "settingsResetTabButton")
        assert reset is not None and reset.isEnabled()
    finally:
        dialog.close()
        dialog.deleteLater()


@pytest.mark.parametrize(
    "category,setup_method,file,resource",
    [
        ("document", "setup_for_document_file", "/sample.docx", "gongwen"),
        ("image", "setup_for_image_file", "/sample.png", "invoice_cn"),
        ("layout", "setup_for_layout_file", "/sample.pdf", "invoice_cn"),
    ],
)
def test_language_change_clears_open_input_selection_without_erasing_saved_default(
    qapp, category: str, setup_method: str, file: str, resource: str
) -> None:
    set_locale("zh_CN")
    key = f"{category}.to_md_optimization"
    main_vm = FakeMainWindowViewModel({key: resource})
    vm = ActionAreaViewModel(cast(Any, main_vm))
    setup = getattr(vm, setup_method)
    setup(file)
    assert vm.optimize_for_type == resource
    set_locale("en_US")
    assert vm.action_name == ""
    assert vm.optimize_for_type is None
    assert not vm.show_optimize
    assert vm.optimization_route_options == ()
    assert main_vm.controller.config_port.get(key) == resource
    vm.optimize_for_type = resource
    assert vm.action_name == ""
    set_locale("zh_CN")
    vm.refresh_optimization_choices()
    assert vm.optimize_for_type is None
    setup(file)
    assert vm.optimize_for_type == resource


def _future_projection() -> dict[str, object]:
    projection = optimization_capability_projection()
    resources = cast(dict[str, Any], projection["optimizations"])["resources"]
    resources[0]["id"] = "english-documents"
    resources[0]["name"] = "English documents"
    projection["optimization_applicable_locales"] = {"english-documents": ["en_US", "ja_JP"], "invoice_cn": ["zh_CN"]}
    return projection


def test_future_resource_has_generic_defaults_page_with_real_apply_and_cancel(qapp, tmp_path, monkeypatch) -> None:
    set_locale("en_US")
    loader = ConfigLoader(base_dir=Path(__file__).resolve().parents[4] / "configs", user_dir=tmp_path / "config")
    port = ConfigPortAdapter(loader)
    assert port.set("document.to_md_optimization", "gongwen")
    controller = ApplicationController(config_port=port)
    monkeypatch.setattr(controller, "describe_gui_runtime_capabilities", _future_projection)
    vm = SettingsViewModel(controller=controller)
    dialog = SettingsDialog(view_model=vm)
    try:
        assert not dialog.activate_section("gongwen")
        assert dialog.activate_section("optimization:english-documents")
        button = dialog.findChild(QPushButton, "documentOptimizationDefaultButton")
        assert button is not None
        reset = dialog.findChild(QPushButton, "settingsResetTabButton")
        assert reset is not None and not reset.isEnabled()
        button.click()
        assert vm.get_default_optimization("document") == "english-documents"
        assert port.get("document.to_md_optimization") == "gongwen"
        vm.cancel()
        assert vm.get_default_optimization("document") == "gongwen"
        button.click()
        assert vm.apply_changes()
        assert port.get("document.to_md_optimization") == "english-documents"
        button.click()
        assert vm.get_default_optimization("document") == "none"
        vm.cancel()
        assert vm.get_default_optimization("document") == "english-documents"
        assert port.get("document.to_md_optimization") == "english-documents"
    finally:
        vm.cancel()
        dialog.close()
        dialog.deleteLater()


def test_future_resource_keeps_category_filter_and_unavailable_routes(qapp) -> None:
    set_locale("en_US")
    controller = FakeController()
    controller.describe_gui_runtime_capabilities = _future_projection  # type: ignore[method-assign]
    vm = _vm(controller)
    assert list(vm.get_optimization_choices_result(source_category="document").items) == ["english-documents"]
    assert vm.get_optimization_choices_result(source_category="image").choices == ()
    assert vm.get_optimization_choices_result(source_category="layout").choices == ()
    assert vm.get_optimization_choices_result(locale="zh_TW").choices == ()
