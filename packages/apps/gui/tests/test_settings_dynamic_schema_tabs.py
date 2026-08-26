from __future__ import annotations

from collections.abc import Mapping
from typing import Any, cast

import pytest
from PySide6.QtWidgets import QCheckBox, QComboBox, QSpinBox, QWidget

pytestmark = pytest.mark.gui


def _control[ControlT: QWidget](widgets: Mapping[str, QWidget], key: str, expected_type: type[ControlT]) -> ControlT:
    control = widgets[key]
    assert isinstance(control, expected_type), f"{key} must create {expected_type.__name__}"
    return control


def _combo_values(combo: QComboBox) -> list[object]:
    return [combo.itemData(i) for i in range(combo.count())]


def test_image_and_other_dynamic_schema_tabs_create_with_expected_values(qapp) -> None:
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.image_tab import ImageTab
    from docwen_gui.widgets.settings.other_tab import OtherTab

    vm = SettingsViewModel(config=SettingsConfig())
    image_tab = ImageTab(vm)
    other_tab = OtherTab(vm)

    image_widgets = image_tab.widgets
    other_widgets = other_tab.widgets
    image_keep_images = _control(image_widgets, "to_md_keep_images", QCheckBox)
    image_enable_ocr = _control(image_widgets, "to_md_enable_ocr", QCheckBox)
    ocr_language = _control(image_widgets, "ocr_language", QComboBox)
    compress_mode = _control(image_widgets, "compress_mode", QComboBox)
    size_limit = _control(image_widgets, "size_limit", QSpinBox)
    size_unit = _control(image_widgets, "size_unit", QComboBox)
    pdf_quality = _control(image_widgets, "pdf_quality", QComboBox)
    tiff_mode = _control(image_widgets, "tiff_mode", QComboBox)
    other_keep_images = _control(other_widgets, "to_md_keep_images", QCheckBox)
    other_enable_ocr = _control(other_widgets, "to_md_enable_ocr", QCheckBox)

    assert _combo_values(ocr_language) == [
        "auto",
        "chinese",
        "chinese_cht",
        "english",
        "japanese",
        "korean",
        "latin",
        "cyrillic",
    ]
    assert _combo_values(compress_mode) == ["lossless", "limit_size"]
    assert _combo_values(pdf_quality) == ["original", "fit_a4", "fit_a3"]
    assert _combo_values(tiff_mode) == ["smart", "rgb"]
    assert sorted(other_widgets) == ["to_md_enable_ocr", "to_md_keep_images"]

    image_keep_images.setChecked(True)
    image_enable_ocr.setChecked(True)
    image_tab.set_combo_data(ocr_language, "japanese")
    image_tab.set_combo_data(compress_mode, "limit_size")
    size_limit.setValue(1024)
    image_tab.set_combo_data(size_unit, "MB")
    image_tab.set_combo_data(pdf_quality, "fit_a4")
    image_tab.set_combo_data(tiff_mode, "rgb")
    other_keep_images.setChecked(True)
    other_enable_ocr.setChecked(True)

    config = vm.config
    assert config.conversion_defaults.image["to_md_keep_images"] is True
    assert config.conversion_defaults.image["to_md_enable_ocr"] is True
    assert config.conversion_defaults.image["ocr_language"] == "japanese"
    assert config.conversion_defaults.image["compress_mode"] == "limit_size"
    assert config.conversion_defaults.image["size_limit"] == 1024
    assert config.conversion_defaults.image["size_unit"] == "MB"
    assert config.conversion_defaults.image["pdf_quality"] == "fit_a4"
    assert config.conversion_defaults.image["tiff_mode"] == "rgb"
    assert config.conversion_defaults.other["to_md_keep_images"] is True
    assert config.conversion_defaults.other["to_md_enable_ocr"] is True


def test_document_layout_spreadsheet_dynamic_schema_values(qapp, request: pytest.FixtureRequest) -> None:
    from tests.support.gui_vm_fakes import FakeController

    from docwen_core.models.manifest import OptimizationResourceSpec, PluginManifest, RouteSpec
    from docwen_gui.models.settings_config import SettingsConfig
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import SettingsDialog
    from docwen_gui.widgets.settings.document_tab import DocumentTab
    from docwen_gui.widgets.settings.layout_tab import LayoutTab
    from docwen_gui.widgets.settings.spreadsheet_tab import SpreadsheetTab
    from docwen_runtime.capabilities import build_runtime_capability_projection

    config = SettingsConfig()
    config.text.numbering_schemes = {
        "settings": {"order": ["hierarchical_standard", "legal_standard"]},
        "schemes": {
            "hierarchical_standard": {"name": "Hierarchical"},
            "legal_standard": {"name": "Legal"},
        },
    }
    config.conversion_defaults.document = {
        "to_md_remove_numbering": True,
        "to_md_add_numbering": False,
        "to_md_default_scheme": "hierarchical_standard",
        "to_md_enable_optimization": False,
        "to_md_optimization_type": "gongwen",
    }
    projection = build_runtime_capability_projection(
        [
            PluginManifest(
                plugin_id="settings_schema_optimizers",
                name="Settings schema optimizers",
                version="1",
                routes=[
                    RouteSpec("docx", "md", action_name="gongwen"),
                    RouteSpec("pdf", "md", action_name="invoice_cn"),
                ],
                optimization_resources=[
                    OptimizationResourceSpec("gongwen", "Gongwen", "gongwen"),
                    OptimizationResourceSpec("invoice_cn", "Invoice CN", "invoice_cn"),
                ],
            )
        ],
        platform_id="windows",
        egress_guard_status={},
    )
    controller = FakeController()
    controller.describe_runtime_capabilities = lambda: projection  # type: ignore[method-assign]
    vm = SettingsViewModel(config=config)
    vm._controller = cast(Any, controller)  # pyright: ignore[reportPrivateUsage]
    dialog = SettingsDialog(view_model=vm)
    request.addfinalizer(dialog.deleteLater)
    request.addfinalizer(dialog.close)
    document_tab = dialog._tabs["document"]  # pyright: ignore[reportPrivateUsage]
    layout_tab = dialog._tabs["layout"]  # pyright: ignore[reportPrivateUsage]
    spreadsheet_tab = dialog._tabs["spreadsheet"]  # pyright: ignore[reportPrivateUsage]
    assert isinstance(document_tab, DocumentTab)
    assert isinstance(layout_tab, LayoutTab)
    assert isinstance(spreadsheet_tab, SpreadsheetTab)

    document_widgets = document_tab.widgets
    layout_widgets = layout_tab.widgets
    spreadsheet_widgets = spreadsheet_tab.widgets
    document_keep_images = _control(document_widgets, "to_md_keep_images", QCheckBox)
    document_enable_ocr = _control(document_widgets, "to_md_enable_ocr", QCheckBox)
    document_remove_numbering = _control(document_widgets, "to_md_remove_numbering", QCheckBox)
    document_add_numbering = _control(document_widgets, "to_md_add_numbering", QCheckBox)
    document_default_scheme = _control(document_widgets, "to_md_default_scheme", QComboBox)
    document_enable_optimization = _control(document_widgets, "to_md_enable_optimization", QCheckBox)
    document_optimization_type = _control(document_widgets, "to_md_optimization_type", QComboBox)
    document_table_merge = _control(document_widgets, "to_md_table_merge_export_strategy", QComboBox)
    layout_keep_images = _control(layout_widgets, "to_md_keep_images", QCheckBox)
    layout_enable_ocr = _control(layout_widgets, "to_md_enable_ocr", QCheckBox)
    layout_enable_optimization = _control(layout_widgets, "to_md_enable_optimization", QCheckBox)
    layout_optimization_type = _control(layout_widgets, "to_md_optimization_type", QComboBox)
    layout_render_dpi = _control(layout_widgets, "render_dpi", QComboBox)
    spreadsheet_keep_images = _control(spreadsheet_widgets, "to_md_keep_images", QCheckBox)
    spreadsheet_enable_ocr = _control(spreadsheet_widgets, "to_md_enable_ocr", QCheckBox)
    spreadsheet_table_merge = _control(spreadsheet_widgets, "to_md_table_merge_export_strategy", QComboBox)
    spreadsheet_merge_mode = _control(spreadsheet_widgets, "merge_mode", QComboBox)

    assert _combo_values(document_optimization_type) == ["gongwen"]
    assert _combo_values(document_table_merge) == ["fill", "empty", "marker"]
    assert _combo_values(document_default_scheme) == ["hierarchical_standard", "legal_standard"]
    assert document_remove_numbering.isChecked() is True
    assert document_add_numbering.isChecked() is False
    assert document_default_scheme.currentData() == "hierarchical_standard"
    assert document_default_scheme.isEnabled() is False
    assert _combo_values(layout_optimization_type) == ["invoice_cn"]
    assert _combo_values(layout_render_dpi) == [150, 300, 600]
    assert _combo_values(spreadsheet_table_merge) == ["fill", "empty", "marker"]
    assert _combo_values(spreadsheet_merge_mode) == [1, 2, 3]

    document_keep_images.setChecked(True)
    document_enable_ocr.setChecked(True)
    document_remove_numbering.setChecked(False)
    document_add_numbering.setChecked(True)
    document_tab.set_combo_data(document_default_scheme, "legal_standard")
    document_enable_optimization.setChecked(True)
    document_tab.set_combo_data(document_optimization_type, "gongwen")
    document_tab.set_combo_data(document_table_merge, "marker")
    layout_keep_images.setChecked(True)
    layout_enable_ocr.setChecked(True)
    layout_enable_optimization.setChecked(True)
    layout_tab.set_combo_data(layout_render_dpi, 600)
    spreadsheet_keep_images.setChecked(True)
    spreadsheet_enable_ocr.setChecked(True)
    spreadsheet_tab.set_combo_data(spreadsheet_table_merge, "marker")
    spreadsheet_tab.set_combo_data(spreadsheet_merge_mode, 2)

    config = vm.config
    assert config.conversion_defaults.document["to_md_keep_images"] is True
    assert config.conversion_defaults.document["to_md_enable_ocr"] is True
    assert config.conversion_defaults.document["to_md_remove_numbering"] is False
    assert config.conversion_defaults.document["to_md_add_numbering"] is True
    assert config.conversion_defaults.document["to_md_default_scheme"] == "legal_standard"
    assert document_default_scheme.isEnabled() is True
    assert config.conversion_defaults.document["to_md_enable_optimization"] is True
    assert config.conversion_defaults.document["to_md_optimization_type"] == "gongwen"
    assert config.conversion_defaults.document["to_md_table_merge_export_strategy"] == "marker"
    assert config.conversion_defaults.layout["to_md_keep_images"] is True
    assert config.conversion_defaults.layout["to_md_enable_ocr"] is True
    assert config.conversion_defaults.layout["to_md_enable_optimization"] is True
    assert config.conversion_defaults.layout["render_dpi"] == 600
    assert config.conversion_defaults.spreadsheet["to_md_keep_images"] is True
    assert config.conversion_defaults.spreadsheet["to_md_enable_ocr"] is True
    assert config.conversion_defaults.spreadsheet["to_md_table_merge_export_strategy"] == "marker"
    assert config.conversion_defaults.spreadsheet["merge_mode"] == 2

    collected = document_tab.collect_values()
    assert collected["to_md_remove_numbering"] is False
    assert collected["to_md_add_numbering"] is True
    assert collected["to_md_default_scheme"] == "legal_standard"
