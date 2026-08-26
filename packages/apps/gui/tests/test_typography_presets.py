"""Visible semantic typography preset contracts."""

from __future__ import annotations

import re

import pytest
from tests.support.gui_vm_fakes import FakeMainWindowViewModel

pytestmark = pytest.mark.gui


def _point_size(widget) -> int:
    size = widget.font().pointSize()
    assert size > 0
    return size


def test_semantic_typography_preserves_role_hierarchy() -> None:
    from docwen_gui.styles.design_tokens import Typography

    expected = {
        "small": (10, 11, 12, 13, 14, 15, 17, 19),
        "default": (11, 12, 13, 14, 15, 16, 18, 20),
        "large": (12, 13, 14, 15, 16, 17, 19, 21),
        "xlarge": (14, 15, 16, 17, 18, 19, 21, 23),
    }
    bases = (
        Typography.CAPTION_SIZE,
        Typography.BODY_SIZE,
        Typography.CARD_TITLE_SIZE,
        Typography.SECTION_TITLE_SIZE,
        Typography.EMPHASIS_TITLE_SIZE,
        Typography.HERO_SIZE,
        Typography.PAGE_TITLE_SIZE,
        Typography.DIALOG_TITLE_SIZE,
    )

    for preset, sizes in expected.items():
        assert tuple(Typography.resolve(base, preset) for base in bases) == sizes


def test_global_stylesheet_has_one_dynamic_typography_owner() -> None:
    from docwen_gui.styles.global_aggregate import build_global_stylesheet

    default_css = build_global_stylesheet("light", "default")
    xlarge_css = build_global_stylesheet("light", "xlarge")

    assert "/* docwen-application-typography */" in default_css
    assert "font-size: 12pt;" in default_css
    assert "font-size: 16pt;" in default_css
    assert "font-size: 15pt;" in xlarge_css
    assert "font-size: 19pt;" in xlarge_css
    assert re.search(r"font-size\s*:\s*\d+px", default_css) is None


def test_dark_settings_secondary_buttons_use_application_theme() -> None:
    from docwen_gui.styles.global_aggregate import build_global_stylesheet

    dark_css = build_global_stylesheet("dark", "xlarge")

    assert "QWidget#settingsTabRoot QPushButton" in dark_css
    assert "QDialog#numberingAddDialog QPushButton" in dark_css
    assert "QDialog#numberingAddDialog QToolButton" in dark_css
    assert "QWidget#tomlEditorDialog QPushButton" in dark_css


def test_existing_main_window_roles_follow_runtime_preset(qapp) -> None:
    from PySide6.QtWidgets import QLabel

    from docwen_gui.main_window import MainWindow
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    window = MainWindow(view_model=MainWindowViewModel(controller=None))
    window.setup_ui()
    window.show()
    qapp.processEvents()
    try:
        prompt = window.findChild(QLabel, "fileDropPromptLabel")
        conversion_description = window.findChild(QLabel, "conversionSectionDescription")
        version = window.findChild(QLabel, "versionLabel")
        assert prompt is not None
        assert conversion_description is not None
        assert version is not None

        window._apply_font_size_preset("default")
        qapp.processEvents()
        assert (_point_size(prompt), _point_size(conversion_description), _point_size(version)) == (16, 12, 12)

        window._apply_font_size_preset("xlarge")
        qapp.processEvents()
        assert (_point_size(prompt), _point_size(conversion_description), _point_size(version)) == (19, 15, 15)
    finally:
        window.close()
        ThemeManager.reset_instance()


def test_xlarge_drop_pyramid_keeps_every_label_inside_its_row(qapp, qtbot) -> None:
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.input_area_vm import InputAreaViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel
    from docwen_gui.widgets.input_area import InputArea

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_font_size_preset("xlarge")
    main_vm = MainWindowViewModel(controller=None)
    input_vm = InputAreaViewModel(main_vm=main_vm)
    widget = InputArea(view_model=input_vm)
    widget.setFixedSize(460, 316)
    widget.show()
    input_vm.set_mode("batch")
    for _ in range(3):
        qapp.processEvents()
    try:
        prompt = widget._prompt_label
        qtbot.waitUntil(
            lambda: all(
                type_label.sizeHint().width() <= type_label.width()
                and value_label.sizeHint().width() <= value_label.width()
                for _row, _layout, type_label, value_label in widget._type_prompt_rows
            ),
            timeout=3000,
        )
        assert not prompt.wordWrap()
        assert prompt.sizeHint().width() <= prompt.width()
        for row, _layout, type_label, value_label in widget._type_prompt_rows:
            row_left = row.mapTo(widget._empty_center_panel, row.rect().topLeft()).x()
            row_right = row.mapTo(widget._empty_center_panel, row.rect().topRight()).x()
            assert row_left >= 0
            assert row_right < widget._empty_center_panel.width()
            assert type_label.geometry().left() >= 0
            assert value_label.geometry().right() < row.width()
            assert type_label.geometry().right() < value_label.geometry().left()
            assert type_label.sizeHint().width() <= type_label.width()
            assert value_label.sizeHint().width() <= value_label.width()
    finally:
        widget.close()
        ThemeManager.reset_instance()


def test_xlarge_batch_toolbar_wraps_before_labels_are_squeezed(qapp, qtbot) -> None:
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.widgets.batch_list import BatchList

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_font_size_preset("xlarge")
    widget = BatchList(view_model=BatchListViewModel())
    widget.setFixedWidth(375)
    widget.resize(375, 640)
    widget.show()
    qapp.processEvents()
    try:
        required_width = (
            widget.filter_button.sizeHint().width()
            + widget._summary_reorder_frame.sizeHint().width()
            + widget._summary_header_layout.horizontalSpacing()
        )
        widget._summary_header.setFixedWidth(required_width)
        widget._sync_summary_header_layout()
        qtbot.waitUntil(lambda: widget._summary_header_compact, timeout=3000)
        for button in (
            widget.filter_button,
            widget.move_up_button,
            widget.move_down_button,
            widget.sort_button,
        ):
            assert button.sizeHint().width() <= button.width()
    finally:
        widget.close()
        ThemeManager.reset_instance()


def test_batch_status_pulse_returns_to_semantic_font_size(qapp) -> None:
    from docwen_gui.view_models.batch_list_vm import BatchFileEntry
    from docwen_gui.widgets.batch_list import BatchEntryItemWidget

    entry = BatchFileEntry(
        file_path="sample.docx",
        file_name="sample.docx",
        detected_format="docx",
        workflow_category="document",
    )
    widget = BatchEntryItemWidget(entry)
    base_size = _point_size(widget.info_badge)
    try:
        widget._status_badge_pulse_base_size = base_size
        widget._apply_pulse(1)
        widget._apply_pulse(2)
        assert _point_size(widget.info_badge) == base_size + 2
        widget._finish_pulse()
        assert _point_size(widget.info_badge) == base_size
    finally:
        widget.close()


def test_xlarge_batch_entry_wraps_before_name_and_badge_collide(qapp) -> None:
    from PySide6.QtWidgets import QListWidget, QListWidgetItem

    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.batch_list_vm import BatchFileEntry
    from docwen_gui.widgets.batch_list import BatchEntryItemWidget

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "dark")
    manager.apply_font_size_preset("xlarge")
    entry = BatchFileEntry(
        file_path="very_long_filename_for_status_balance_review_document_v3.docx",
        file_name="very_long_filename_for_status_balance_review_document_v3.docx",
        detected_format="docx",
        workflow_category="document",
        size_bytes=1_048_576,
        status="completed",
        output_path="very_long_filename_for_status_balance_review_document_v3.md",
    )
    list_widget = QListWidget()
    list_widget.resize(400, 600)
    item = QListWidgetItem(list_widget)
    widget = BatchEntryItemWidget(entry)
    list_widget.setItemWidget(item, widget)
    widget.bind_list_item(list_widget, item)
    list_widget.show()
    try:
        for _ in range(3):
            qapp.processEvents()
        assert widget._is_compact
        assert widget.name_label.text().replace("\u200b", "") == entry.file_name
        wrapped_name_height = widget.name_label.heightForWidth(widget.name_label.width())
        assert wrapped_name_height <= widget.name_label.height()
        assert widget.info_badge.sizeHint().width() <= widget.info_badge.width()
        item_layout = widget.layout()
        assert item_layout is not None
        assert item.sizeHint().height() >= item_layout.totalHeightForWidth(list_widget.viewport().width())
    finally:
        list_widget.close()
        ThemeManager.reset_instance()


def test_xlarge_layout_render_controls_stay_inside_right_panel(qapp) -> None:
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
    from docwen_gui.widgets.conversion_panel import ConversionPanel

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "dark")
    manager.apply_font_size_preset("xlarge")
    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(view_model=vm)
    widget.resize(576, 760)
    vm.set_file_info("layout", "pdf", file_path="annual-report-2026.pdf")
    widget.show()
    try:
        for _ in range(3):
            qapp.processEvents()
        render_button = widget._layout_render_button
        assert render_button is not None
        row = render_button.parentWidget()
        assert row is not None
        assert render_button.sizeHint().width() <= render_button.width()
        assert render_button.geometry().right() < row.width()
        assert render_button.geometry().top() > widget._layout_render_format_combo.geometry().bottom()
    finally:
        widget.close()
        ThemeManager.reset_instance()


def test_xlarge_conversion_detail_controls_wrap_without_clipping(qapp) -> None:
    from PySide6.QtWidgets import QLabel

    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
    from docwen_gui.widgets.conversion_panel import ConversionPanel

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_font_size_preset("xlarge")
    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(view_model=vm)
    widget.resize(576, 760)
    widget.show()
    try:
        vm.set_file_info("spreadsheet", "xlsx", file_path="invoice-summary-2026.xlsx")
        for _ in range(3):
            qapp.processEvents()
        consent_label = widget.findChild(QLabel, "conversionWrappingCheckLabel")
        assert consent_label is not None
        assert consent_label.wordWrap()
        assert consent_label.heightForWidth(consent_label.width()) <= consent_label.height()
        consent_parent = consent_label.parentWidget()
        assert consent_parent is not None
        assert consent_label.geometry().right() < consent_parent.width()

        vm.set_file_info("image", "png", file_path="scanned-contract-page-01.png")
        for _ in range(3):
            qapp.processEvents()
        assert widget._size_limit_edit.sizeHint().width() <= widget._size_limit_edit.width()
        assert widget._size_unit_combo.sizeHint().width() <= widget._size_unit_combo.width()
        assert widget._size_unit_combo.currentText() == "KB"
    finally:
        widget.close()
        ThemeManager.reset_instance()


def test_xlarge_guide_actions_wrap_before_labels_are_squeezed(qapp, qtbot) -> None:
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
    from docwen_gui.widgets.info_area import InfoArea

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "dark")
    manager.apply_font_size_preset("xlarge")
    vm = InfoAreaViewModel()
    widget = InfoArea(view_model=vm)
    widget.setFixedWidth(576)
    widget.resize(576, 500)
    widget.show()
    vm.set_task_summary(
        operation_id="failed-xlarge",
        state="failed",
        tone="danger",
        guide_actions=[
            {"action_key": "open_output_dir", "target_path": "/tmp/out"},
            {"action_key": "view_failed_details", "target_path": "/tmp/failed.json"},
            {"action_key": "retry_failed", "target_path": ""},
            {"action_key": "add_more_files", "target_path": ""},
        ],
    )
    try:
        for _ in range(3):
            qapp.processEvents()
        buttons = widget.find_guide_buttons()
        assert len(buttons) == 4
        spacing = max(0, widget._status_guide_actions_layout.horizontalSpacing())
        one_row_width = sum(button.sizeHint().width() for button in buttons) + spacing * (len(buttons) - 1)
        widget._status_guide_actions_widget.setFixedWidth(one_row_width)
        widget._sync_guide_button_layout()
        qtbot.waitUntil(
            lambda: len({button.geometry().top() for button in buttons}) >= 2,
            timeout=3000,
        )
        assert len({button.geometry().top() for button in buttons}) >= 2
        for button in buttons:
            assert button.sizeHint().width() <= button.width()
            assert button.geometry().right() < widget._status_guide_actions_widget.width()
    finally:
        widget.close()
        ThemeManager.reset_instance()


def test_existing_and_later_settings_widgets_agree(qapp) -> None:
    from PySide6.QtWidgets import QLabel, QPushButton

    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_font_size_preset("default")
    existing = SettingsDialog(view_model=SettingsViewModel())
    existing.show()
    qapp.processEvents()
    later = None
    try:
        existing_title = existing.findChild(QLabel, "settingsTabTitle")
        existing_apply = existing.findChild(QPushButton, "settingsApplyButton")
        assert existing_title is not None
        assert existing_apply is not None
        assert (_point_size(existing_title), _point_size(existing_apply)) == (18, 12)

        manager.apply_font_size_preset("xlarge")
        qapp.processEvents()
        assert (_point_size(existing_title), _point_size(existing_apply)) == (21, 15)

        later = SettingsDialog(view_model=SettingsViewModel())
        later.show()
        qapp.processEvents()
        later_title = later.findChild(QLabel, "settingsTabTitle")
        later_apply = later.findChild(QPushButton, "settingsApplyButton")
        assert later_title is not None
        assert later_apply is not None
        assert (_point_size(later_title), _point_size(later_apply)) == (21, 15)
    finally:
        existing.close()
        if later is not None:
            later.close()
        ThemeManager.reset_instance()


def test_xlarge_settings_pages_do_not_require_horizontal_scrolling(qapp) -> None:
    from PySide6.QtWidgets import QScrollArea

    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_font_size_preset("xlarge")
    dialog = SettingsDialog(view_model=SettingsViewModel())
    dialog.show()
    try:
        for index in range(dialog._tab_widget.count()):
            dialog._tab_widget.setCurrentIndex(index)
            for _ in range(2):
                qapp.processEvents()
            page = dialog._tab_widget.widget(index)
            assert page is not None
            scroll = page.findChild(QScrollArea, "settingsTabScrollArea")
            assert scroll is not None
            assert scroll.horizontalScrollBar().maximum() == 0, dialog._tab_widget.tabText(index)
    finally:
        dialog.close()
        ThemeManager.reset_instance()
