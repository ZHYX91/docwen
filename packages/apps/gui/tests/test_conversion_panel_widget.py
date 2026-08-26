"""Smoke tests for ConversionPanel widget.

These tests validate widget construction, ViewModel binding, and
category-based layout switching. Require a QApplication instance.
"""

from collections.abc import Generator

import pytest
from PySide6.QtWidgets import QApplication, QCheckBox, QLabel, QLineEdit, QScrollArea
from tests.support.gui_vm_fakes import FakeMainWindowViewModel

from docwen_gui.i18n import t
from docwen_gui.styles._hex_helper import _hex_to_rgba
from docwen_gui.styles.conversion_panel import build_conversion_panel_stylesheet
from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.styles.theme_semantics import get_theme_class_color
from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
from docwen_gui.widgets.conversion_panel import ConversionPanel

pytestmark = pytest.mark.gui


def _combo_icon_center_color(combo, index: int) -> str:
    image = combo.itemIcon(index).pixmap(12, 12).toImage()
    return image.pixelColor(image.width() // 2, image.height() // 2).name().upper()


@pytest.fixture
def vm() -> ConversionPanelViewModel:
    return ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]


@pytest.fixture
def widget(qapp: QApplication, vm: ConversionPanelViewModel) -> "Generator[ConversionPanel, None, None]":
    w = ConversionPanel(view_model=vm)
    yield w
    w.deleteLater()


# ── Construction ──────────────────────────────────────────────────────


class TestConstruction:
    def test_widget_created(self, widget: ConversionPanel) -> None:
        assert widget is not None

    def test_object_name(self, widget: ConversionPanel) -> None:
        assert widget.objectName() == "conversionPanelRoot"

    def test_focus_policy(self, widget: ConversionPanel) -> None:
        from PySide6.QtCore import Qt

        assert widget.focusPolicy() == Qt.FocusPolicy.StrongFocus

    def test_view_model_access(self, widget: ConversionPanel) -> None:
        assert widget.view_model is not None
        assert isinstance(widget.view_model, ConversionPanelViewModel)

    def test_hint_label_visible_initially(self, widget: ConversionPanel) -> None:
        # With no file set, hint label should be visible
        hint = widget._hint_label
        assert hint is not None

    def test_panel_avoids_repeating_file_mode_and_format_context(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("document", "docx", file_path="/work/report.docx")

        assert widget.findChild(QLabel, "conversionPanelTitle") is None
        assert widget.findChild(QLabel, "conversionPanelSubtitle") is None
        assert widget.findChild(QLabel, "conversionSectionMeta") is None
        assert "report.docx" not in " ".join(label.text() for label in widget.findChildren(QLabel))

    def test_extra_group_hidden_initially(self, widget: ConversionPanel) -> None:
        assert widget.extra_group.isVisible() is False

    def test_stylesheet_overrides_action_buttons(self) -> None:
        stylesheet = build_conversion_panel_stylesheet("dark")

        assert "conversionPanelHeader" not in stylesheet
        assert "conversionSectionMeta" not in stylesheet
        assert "QWidget#conversionPanelRoot QPushButton#conversionActionButton" in stylesheet
        assert "background-color: palette(highlight)" in stylesheet
        assert "QWidget#conversionPanelRoot QPushButton#conversionActionButton:disabled" in stylesheet
        assert "background-color: palette(alternate-base)" in stylesheet

    def test_semantic_card_borders_keep_distinct_readable_theme_tones(self) -> None:
        light = build_conversion_panel_stylesheet("light")
        dark = build_conversion_panel_stylesheet("dark")

        for tone in ("success", "danger", "warning"):
            light_color = get_theme_class_color(tone, "light")
            dark_color = get_theme_class_color(tone, "dark")
            assert f'[accentTone="{tone}"]' in light
            assert f"border: 1px solid {_hex_to_rgba(light_color, 130)};" in light
            assert f"background-color: {_hex_to_rgba(light_color, 10)};" in light
            assert f"border: 1px solid {_hex_to_rgba(dark_color, 132)};" in dark
            assert f"background-color: {_hex_to_rgba(dark_color, 22)};" in dark


# ── Category Switching ────────────────────────────────────────────────


class TestCategorySwitch:
    def test_document_category_builds_ui(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        # Just verify VM state transitions work; widget rebuild is async via signal
        assert vm.file_category == "document"

    def test_spreadsheet_category(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("spreadsheet", "xlsx", file_path="/test.xlsx")
        assert vm.file_category == "spreadsheet"

    def test_image_category(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("image", "png", file_path="/test.png")
        assert vm.file_category == "image"

    def test_layout_category(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        assert vm.file_category == "layout"

    def test_switch_between_categories(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        assert vm.file_category == "document"
        vm.set_file_info("image", "png", file_path="/test.png")
        assert vm.file_category == "image"

    def test_reset_to_no_category(self, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        vm.reset()
        assert vm.file_category is None
        assert vm.current_format == ""
        assert vm.current_file_path is None


# ── Widget Rebuilds on VM State Change ───────────────────────────────


class TestWidgetRebuild:
    def test_rebuild_on_set_file_info(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        # After state change, widgets should be built
        assert widget.conversion_combo is not None
        assert widget.conversion_button is not None

    def test_rebuild_on_reset(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        vm.reset()
        # After reset, hint label shows, no combos
        assert widget.conversion_combo is None

    def test_document_has_saveas_combo(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        assert widget.saveas_combo is not None
        assert widget.saveas_button is not None
        assert widget.conversion_button is not None
        assert widget.conversion_button.objectName() == "conversionActionButton"
        assert widget.saveas_button.objectName() == "conversionActionButton"

    def test_format_swatch_recolors_during_live_theme_preview(
        self,
        qapp: QApplication,
        widget: ConversionPanel,
        vm: ConversionPanelViewModel,
    ) -> None:
        manager = ThemeManager.get_instance()
        previous_theme = manager.get_current_theme()
        manager.initialize(qapp, "light")
        try:
            vm.set_file_info("layout", "pdf", file_path="/test.pdf")
            combo = widget._layout_export_combo
            assert combo is not None
            docx_index = combo.findText("DOCX")
            assert _combo_icon_center_color(combo, docx_index) == get_theme_class_color("primary", "light").upper()

            manager.apply_theme("dark")
            qapp.processEvents()

            assert _combo_icon_center_color(combo, docx_index) == get_theme_class_color("primary", "dark").upper()
        finally:
            manager.apply_theme(previous_theme)
            qapp.processEvents()

    def test_csv_spreadsheet_keeps_route_backed_pdf_saveas(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("spreadsheet", "csv", file_path="/test.csv")
        assert widget.saveas_combo is not None
        assert [widget.saveas_combo.itemText(i) for i in range(widget.saveas_combo.count())] == ["PDF"]

    def test_tsv_spreadsheet_hides_unreachable_pdf_saveas(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("spreadsheet", "tsv", file_path="/test.tsv")
        assert widget.saveas_combo is None
        assert widget.saveas_button is None
        assert widget._saveas_group.isHidden()

    def test_layout_has_render_widgets(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        # Layout should have export combo and render widgets
        assert widget._layout_export_combo is not None
        assert widget._layout_render_format_combo is not None
        assert [widget._layout_export_combo.itemText(i) for i in range(widget._layout_export_combo.count())] == [
            "DOCX",
            "DOC",
            "ODT",
            "RTF",
        ]
        assert [
            widget._layout_render_format_combo.itemText(i) for i in range(widget._layout_render_format_combo.count())
        ] == [
            "PNG",
            "JPG",
            "TIF",
        ]

    def test_layout_pdf_does_not_show_same_format_convert_button(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        assert widget.conversion_combo is None
        assert widget.conversion_button is None
        assert widget._conversion_group.isHidden()
        assert not widget._saveas_group.isHidden()

    def test_layout_ofd_shows_pdf_conversion_but_not_pdf_tools(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("layout", "ofd", file_path="/test.ofd")
        assert widget.conversion_combo is not None
        assert [widget.conversion_combo.itemText(i) for i in range(widget.conversion_combo.count())] == ["PDF"]
        assert widget._merge_pdfs_button is None
        assert widget._split_pdf_button is None

    def test_layout_extra_has_merge_split(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        # Extra group visibility flag set to True (hiddenness=False)
        assert not widget.extra_group.isHidden()
        assert widget._merge_pdfs_button is not None
        assert widget._merge_pdfs_button.objectName() == "conversionActionButton"

    def test_image_has_compress_options(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("image", "png", file_path="/test.png")
        # Image should have compression widgets
        assert widget._compress_btn_group is not None

    def test_image_conversion_combo_hides_same_format(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("image", "png", file_path="/test.png")
        assert widget.conversion_combo is not None
        assert "PNG" not in [widget.conversion_combo.itemText(i) for i in range(widget.conversion_combo.count())]
        assert widget.conversion_combo.currentText() == "JPG"

    def test_image_panel_uses_vertical_scroll_when_height_is_constrained(
        self,
        qapp: QApplication,
        widget: ConversionPanel,
        vm: ConversionPanelViewModel,
    ) -> None:
        vm.set_file_info("image", "png", file_path="/test.png")
        widget.resize(360, 360)
        widget.show()
        qapp.processEvents()

        scroll_area = widget.findChild(QScrollArea, "conversionPanelScrollArea")
        assert scroll_area is not None
        assert scroll_area.verticalScrollBar().maximum() > 0
        assert scroll_area.horizontalScrollBar().maximum() == 0

    def test_same_category_format_and_mode_update_in_place(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("image", "png", file_path="/first.png")
        combo = widget.conversion_combo
        size_edit = widget._size_limit_edit

        vm.set_file_info("image", "png", file_path="/second.png")

        assert widget.conversion_combo is combo
        assert widget._size_limit_edit is size_edit

    def test_structural_context_change_rebuilds_controls(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("image", "png", file_path="/test.png")
        image_combo = widget.conversion_combo

        vm.set_file_info("document", "docx", file_path="/test.docx")

        assert widget.conversion_combo is not image_combo

    def test_image_option_edit_preserves_widget_focus_and_scroll(
        self,
        qapp: QApplication,
        widget: ConversionPanel,
        vm: ConversionPanelViewModel,
    ) -> None:
        vm.set_file_info("image", "jpg", file_path="/test.jpg")
        combo = widget.conversion_combo
        size_edit = widget._size_limit_edit
        assert combo is not None
        assert size_edit is not None

        vm.compress_mode = "limit_size"
        assert widget.conversion_combo is combo
        assert widget._size_limit_edit is size_edit
        assert "JPG" in [combo.itemText(index) for index in range(combo.count())]

        widget.resize(360, 260)
        widget.show()
        qapp.processEvents()
        scroll_area = widget.findChild(QScrollArea, "conversionPanelScrollArea")
        assert scroll_area is not None
        scrollbar = scroll_area.verticalScrollBar()
        assert scrollbar.maximum() > 0

        size_edit.setFocus()
        qapp.processEvents()
        assert size_edit.hasFocus()
        scrollbar.setValue(scrollbar.maximum())
        qapp.processEvents()
        scroll_value = scrollbar.value()

        size_edit.setText("512")
        qapp.processEvents()

        assert vm.size_limit == 512
        assert widget._size_limit_edit is size_edit
        assert size_edit.hasFocus()
        assert scrollbar.value() == scroll_value

    def test_validation_changes_update_existing_controls(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        checkbox = widget._validation_checkboxes["symbol_pairing"]
        validate_button = widget._validate_button

        vm.set_validation_option("symbol_pairing", False)

        assert widget._validation_checkboxes["symbol_pairing"] is checkbox
        assert checkbox.isChecked() is False
        assert widget._validate_button is validate_button

    def test_pdf_metadata_updates_in_place_without_repeating_filename(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("layout", "pdf", file_path="/annual-report.pdf")
        page_edit = widget._page_input_edit
        info_label = widget._pdf_info_label
        assert page_edit is not None
        assert info_label is not None

        vm.set_pdf_info(12, "annual-report.pdf")

        assert widget._page_input_edit is page_edit
        assert widget._pdf_info_label is info_label
        assert "12" in info_label.text()
        assert info_label.toolTip() == "annual-report.pdf"
        assert "annual-report.pdf" not in " ".join(label.text() for label in widget.findChildren(QLabel))


# ── Conversion Requests ───────────────────────────────────────────────


class TestConversionRequests:
    def test_convert_button_click(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        emitted: list = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append(f))
        btn = widget.conversion_button
        if btn is not None:
            btn.click()
            assert emitted == ["doc"]

    def test_spreadsheet_convert_button_uses_first_reachable_target(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("spreadsheet", "xlsx", file_path="/test.xlsx")
        emitted: list = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append(f))
        btn = widget.conversion_button
        if btn is not None:
            btn.click()
            assert emitted == ["xls"]

    def test_policy02_xlsx_to_ods_uses_masked_request_scoped_password_and_explicit_consent(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("spreadsheet", "xlsx", file_path="/protected.xlsx")
        password_edit = widget._spreadsheet_password_edit
        consent = widget._spreadsheet_protection_loss_checkbox
        assert password_edit is not None
        assert password_edit.echoMode() == QLineEdit.EchoMode.Password
        assert consent is not None
        assert consent.isChecked() is False

        emitted: list[tuple[str, str, dict]] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))
        assert widget.conversion_combo is not None
        widget.conversion_combo.setCurrentText("ODS")
        password_edit.setText("test")
        consent.setChecked(True)
        assert widget.conversion_button is not None
        widget.conversion_button.click()

        assert emitted == [
            (
                "ods",
                "/protected.xlsx",
                {
                    "spreadsheet_password": "test",
                    "allow_spreadsheet_protection_loss": True,
                },
            )
        ]
        assert password_edit.text() == ""
        assert consent.isChecked() is False

    def test_policy02_password_controls_are_disabled_in_batch_to_prevent_credential_reuse(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info(
            "spreadsheet",
            "xlsx",
            file_path="/one.xlsx",
            file_list=["/one.xlsx", "/two.xlsx"],
            ui_mode="batch",
        )

        assert widget._spreadsheet_password_edit is not None
        assert widget._spreadsheet_password_edit.isEnabled() is False
        assert widget._spreadsheet_protection_loss_checkbox is not None
        assert widget._spreadsheet_protection_loss_checkbox.isEnabled() is False

    def test_policy02_options_do_not_leak_to_non_ods_target(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("spreadsheet", "xlsx", file_path="/protected.xlsx")
        assert widget._spreadsheet_password_edit is not None
        assert widget._spreadsheet_protection_loss_checkbox is not None
        widget._spreadsheet_password_edit.setText("test")
        widget._spreadsheet_protection_loss_checkbox.setChecked(True)
        assert widget.conversion_combo is not None
        widget.conversion_combo.setCurrentText("XLS")
        emitted: list[tuple[str, str, dict]] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))

        assert widget.conversion_button is not None
        widget.conversion_button.click()

        assert emitted[-1][2] == {}

    def test_tsv_convert_button_uses_only_reachable_target(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("spreadsheet", "tsv", file_path="/test.tsv")
        emitted: list = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append(f))
        btn = widget.conversion_button
        if btn is not None:
            btn.click()
            assert emitted == ["xlsx"]

    def test_validate_button_click(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")
        emitted: list = []
        vm.named_action_requested.connect(lambda n, fp, o: emitted.append(n))
        btn = widget._validate_button
        if btn is not None:
            btn.click()
            assert len(emitted) >= 1

    def test_document_validation_options_use_locale_labels(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("document", "docx", file_path="/test.docx")

        labels = {checkbox.text() for checkbox in widget.findChildren(QCheckBox)}
        expected = {
            t("conversion_panel.document.symbol_pairing"),
            t("conversion_panel.document.typos_rule"),
            t("conversion_panel.document.symbol_correction"),
            t("conversion_panel.document.sensitive_word"),
        }
        assert expected.issubset(labels)
        assert "Symbol Pairing" not in labels

    def test_image_convert_click_emits_compression_options(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.compress_mode = "limit_size"
        vm.size_limit = 512
        vm.size_unit = "KB"
        vm.set_file_info("image", "png", file_path="/test.png")
        emitted: list[tuple[str, str, dict]] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))
        assert widget.conversion_combo is not None
        widget.conversion_combo.setCurrentText("JPG")

        btn = widget.conversion_button
        assert btn is not None
        btn.click()

        assert emitted
        assert emitted[-1][2] == {
            "compress_mode": "limit_size",
            "size_limit": 512,
            "size_unit": "KB",
        }

    def test_image_limit_size_combo_allows_current_jpg_for_original_compression(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.compress_mode = "limit_size"
        vm.size_limit = 128
        vm.size_unit = "KB"
        vm.set_file_info("image", "jpg", file_path="/test.jpg")
        emitted: list[tuple[str, str, dict]] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))

        assert widget.conversion_combo is not None
        items = [widget.conversion_combo.itemText(i) for i in range(widget.conversion_combo.count())]
        assert "JPG" in items
        widget.conversion_combo.setCurrentText("JPG")

        btn = widget.conversion_button
        assert btn is not None
        btn.click()

        assert emitted[-1] == (
            "jpg",
            "/test.jpg",
            {"compress_mode": "limit_size", "size_limit": 128, "size_unit": "KB"},
        )

    def test_image_pdf_click_emits_normalized_quality(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.pdf_quality = "fit_a4"
        vm.set_file_info("image", "png", file_path="/test.png")
        emitted: list[tuple[str, str, dict]] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))

        btn = widget.saveas_button
        assert btn is not None
        btn.click()

        assert emitted[-1][0] == "pdf"
        assert emitted[-1][2] == {"quality_mode": "a4"}

    def test_merge_tiff_click_emits_normalized_mode(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.tiff_mode = "rgb"
        vm.set_file_info("image", "png", file_path="/test.png")
        emitted: list[tuple[str, str, dict]] = []
        vm.named_action_requested.connect(lambda n, fp, o: emitted.append((n, fp, o)))

        btn = widget._merge_tiff_button
        assert btn is not None
        btn.click()

        assert emitted[-1][0] == "merge_images_to_tiff"
        assert emitted[-1][2] == {"mode": "RGB"}

    def test_merge_tables_click_emits_plugin_merge_mode(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.merge_mode = 2
        vm.set_file_info("spreadsheet", "xlsx", file_path="/test.xlsx")
        emitted: list[tuple[str, str, dict]] = []
        vm.named_action_requested.connect(lambda n, fp, o: emitted.append((n, fp, o)))

        btn = widget._merge_tables_button
        assert btn is not None
        btn.click()

        assert emitted[-1][0] == "merge_tables"
        assert emitted[-1][2] == {"merge_mode": "col"}

    def test_layout_render_click_emits_render_dpi(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.render_dpi = 600
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        emitted: list[tuple[str, str, dict]] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))

        assert widget._layout_render_format_combo is not None
        widget._layout_render_format_combo.setCurrentText("PNG")
        assert widget._layout_render_dpi_combo is not None
        widget._layout_render_dpi_combo.setCurrentText("600")
        btn = widget._layout_render_button
        assert btn is not None
        btn.click()

        assert emitted[-1][0] == "png"
        assert emitted[-1][2] == {"render_dpi": 600}

    def test_layout_export_click_can_emit_odt(self, widget: ConversionPanel, vm: ConversionPanelViewModel) -> None:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        emitted: list[tuple[str, str, dict]] = []
        vm.conversion_requested.connect(lambda f, fp, o: emitted.append((f, fp, o)))

        assert widget._layout_export_combo is not None
        widget._layout_export_combo.setCurrentText("ODT")
        btn = widget._layout_export_button
        assert btn is not None
        btn.click()

        assert emitted[-1][0] == "odt"

    def test_layout_split_click_emits_custom_page_options(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        vm.set_pdf_info(3, "test.pdf")
        emitted: list[tuple[str, str, dict]] = []
        vm.named_action_requested.connect(lambda n, fp, o: emitted.append((n, fp, o)))

        assert widget._page_input_edit is not None
        widget._page_input_edit.setText("1")
        btn = widget._split_pdf_button
        assert btn is not None
        assert btn.isEnabled()
        btn.click()

        assert emitted[-1] == (
            "split_pdf",
            "/test.pdf",
            {"split_mode": "custom", "pages": [1]},
        )

    def test_layout_split_click_emits_every_page_mode(
        self, widget: ConversionPanel, vm: ConversionPanelViewModel
    ) -> None:
        vm.set_file_info("layout", "pdf", file_path="/test.pdf")
        vm.set_pdf_info(3, "test.pdf")
        emitted: list[tuple[str, str, dict]] = []
        vm.named_action_requested.connect(lambda n, fp, o: emitted.append((n, fp, o)))

        assert widget._page_input_edit is not None
        widget._page_input_edit.setText("*")
        btn = widget._split_pdf_button
        assert btn is not None
        assert btn.isEnabled()
        btn.click()

        assert emitted[-1] == (
            "split_pdf",
            "/test.pdf",
            {"split_mode": "every_page"},
        )


# ── Focus Management ──────────────────────────────────────────────────


class TestFocusManagement:
    def test_panel_focused_signal(self, widget: ConversionPanel) -> None:
        emitted: list[int] = []
        widget.panel_focused.connect(lambda: emitted.append(1))
        # Simulate focus in
        from PySide6.QtCore import QEvent
        from PySide6.QtGui import QFocusEvent

        event = QFocusEvent(QEvent.Type.FocusIn)
        widget.focusInEvent(event)
        assert len(emitted) == 1


# ── Property Access ───────────────────────────────────────────────────


class TestProperties:
    def test_conversion_combo_none_initially(self, widget: ConversionPanel) -> None:
        assert widget.conversion_combo is None

    def test_conversion_button_none_initially(self, widget: ConversionPanel) -> None:
        assert widget.conversion_button is None

    def test_saveas_combo_none_initially(self, widget: ConversionPanel) -> None:
        assert widget.saveas_combo is None

    def test_saveas_button_none_initially(self, widget: ConversionPanel) -> None:
        assert widget.saveas_button is None
