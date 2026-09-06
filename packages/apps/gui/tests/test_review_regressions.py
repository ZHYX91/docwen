"""Cross-state regressions for execution ownership and responsive analysis."""

from __future__ import annotations

import threading
from types import SimpleNamespace

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QComboBox, QVBoxLayout, QWidget

pytestmark = pytest.mark.gui


def test_active_inputs_survive_delete_and_finished_details_survive_row_removal(
    main_window, qtbot, tmp_path, monkeypatch
):
    from docwen_core.models.result import ConversionErrorInfo, ConversionResult
    from docwen_gui.dialogs.task_details import TaskDetailsDialog
    from docwen_gui.main_window import _normalize_path

    release = threading.Event()

    class Controller:
        def prepare_execution_cancellation(self, *args, **kwargs):
            return object()

        def release_execution_cancellation(self, *args):
            pass

        def execute_single(self, request):
            release.wait(5)
            return ConversionResult(
                task_id=request.request_id,
                success=False,
                error=ConversionErrorInfo(error_type="conversion_error", message="Output locked"),
            )

    source = tmp_path / "sample.md"
    source.write_text("# Sample", encoding="utf-8")
    path = _normalize_path(str(source))
    window = main_window
    window._input_area_vm.set_mode("batch")
    window._input_area_vm.add_files([path])
    window.show()
    window._batch_list.select_file(path)
    context = {
        "request_id": "op",
        "file_path": path,
        "file_paths": [path],
        "options": {},
        "target_format": "docx",
        "action_name": "convert",
    }

    def project():
        window.view_model.begin_execution_telemetry("op", ("op",))
        window._set_execution_file_status(path, "processing", operation_id="op")
        window._action_area_vm.show_cancel()

    try:
        assert window._launch_execution_thread(
            controller=Controller(),
            request=SimpleNamespace(request_id="op"),
            context=context,
            project_reserved_execution=project,
        )
        window._batch_list.setFocus()
        qtbot.keyClick(window._batch_list, Qt.Key.Key_Delete)
        assert window._batch_list_vm.get_file_entry(path) is not None
        assert not window._batch_list_vm.remove_file(path)
        window.view_model.clear_files()
        assert window.view_model.files
    finally:
        release.set()
        qtbot.waitUntil(lambda: not window._active_threads)

    assert window._batch_list_vm.remove_file(path)
    assert not window.view_model.files
    window._handle_task_guide_action("view_failed_details", path)
    dialog = window.findChild(TaskDetailsDialog)
    assert dialog is not None
    assert "Output locked" in dialog.details.toPlainText()
    assert window._task_history.get("op").failed_paths == [path]
    dialog.close()
    retries = []
    monkeypatch.setattr(window, "_start_execution", lambda **kwargs: retries.append(kwargs))
    window._retry_failed_request()
    assert [request["file_path"] for request in retries] == [path]
    assert window._batch_list_vm.get_file_entry(path) is not None


def test_form_reflows_for_runtime_font_and_text_without_window_resize(qtbot, qapp):
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.widgets.panel_card import FormRow

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "light")
    manager.apply_font_size_preset("default")
    host = QWidget()
    qtbot.addWidget(host)
    layout = QVBoxLayout(host)
    row = FormRow("Output format", QComboBox())
    layout.addWidget(row)
    host.setFixedSize(750, 200)
    host.show()
    qtbot.waitUntil(lambda: row.label.width() == row.label.fontMetrics().horizontalAdvance(row.label.text()))
    original_width = row.width()
    try:
        manager.apply_font_size_preset("xlarge")
        qtbot.waitUntil(lambda: row.label.width() >= row.label.fontMetrics().horizontalAdvance(row.label.text()))
        assert row.width() == original_width
        row.label.setText("A much longer output format label")
        qtbot.waitUntil(lambda: row.label.width() >= row.label.fontMetrics().horizontalAdvance(row.label.text()))
    finally:
        manager.apply_font_size_preset("default")
        host.close()
        ThemeManager.reset_instance()


def test_pdf_workflow_fits_default_columns_with_large_typography(main_window_with_controller, qtbot, tmp_path, qapp):
    import fitz
    from PySide6.QtCore import QPoint
    from PySide6.QtWidgets import QScrollArea

    from docwen_gui.styles.theme_manager import ThemeManager

    source = tmp_path / "layout.pdf"
    with fitz.open() as document:
        document.new_page().insert_text((72, 72), "Layout check")
        document.save(source)
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "dark")
    manager.apply_font_size_preset("xlarge")
    window = main_window_with_controller
    try:
        window._input_area_vm.set_mode("batch")
        window._input_area_vm.add_files([str(source)])
        window.show()
        qtbot.waitUntil(lambda: window._conversion_panel._pdf_info_label is not None)
        qtbot.wait(200)
        page_input = window._conversion_panel._page_input_edit
        assert page_input.height() >= page_input.fontMetrics().height() + 10
        assert page_input.parentWidget().rect().contains(page_input.geometry())
        info_label = window._conversion_panel._pdf_info_label
        assert page_input.mapTo(window, QPoint()).y() + page_input.height() <= info_label.mapTo(window, QPoint()).y()
        page_input.setText("1-3,5")
        assert page_input.text() == "1-3,5"
        qtbot.waitUntil(lambda: window._conversion_panel_vm.pdf_total_pages == 1)
        page_input.setText("*")
        assert not window._conversion_panel._page_warning_label.isHidden()
        assert window._conversion_panel._page_warning_label.text()
        page_input.clear()
        assert window._conversion_panel._page_warning_label.isHidden()
        for name in ("centerWorkflowScroll", "conversionPanelScrollArea"):
            scroll = window.findChild(QScrollArea, name)
            assert scroll is not None
            assert scroll.horizontalScrollBar().maximum() == 0, (
                name,
                scroll.viewport().width(),
                [
                    (child.objectName(), type(child).__name__, child.minimumSizeHint().width(), child.minimumWidth())
                    for child in scroll.findChildren(QWidget)
                    if max(child.minimumSizeHint().width(), child.minimumWidth()) > scroll.viewport().width()
                ],
            )
    finally:
        manager.apply_font_size_preset("default")
        manager.apply_theme("light")
        window.close()


def test_mixed_spreadsheets_share_routes_and_protection_gate(qtbot, tmp_path):
    from openpyxl import Workbook
    from tests.support.gui_vm_fakes import FakeMainWindowViewModel

    from docwen_gui.view_models.conversion_panel_vm import ConversionPanelViewModel
    from docwen_gui.widgets.conversion_panel import ConversionPanel

    xlsx = tmp_path / "book.xlsx"
    csv = tmp_path / "table.csv"
    Workbook().save(xlsx)
    csv.write_text("a,b\n1,2", encoding="utf-8")
    formats = {str(xlsx): "xlsx", str(csv): "csv"}
    vm = ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]
    widget = ConversionPanel(vm)
    qtbot.addWidget(widget)
    targets = []
    for path, fmt in formats.items():
        vm.set_file_info("spreadsheet", fmt, path, list(formats), "batch", source_formats=formats)
        widget._conversion_combo.setCurrentText("ODS")
        if vm.spreadsheet_analysis_pending:
            assert not widget._conversion_button.isEnabled()
        qtbot.waitUntil(lambda: not vm.spreadsheet_analysis_pending)
        assert widget._conversion_combo.currentText() == "ODS"
        assert widget._conversion_button.isEnabled(), (
            fmt,
            vm.spreadsheet_protection_info,
            widget._conversion_combo.currentData(),
        )
        assert not vm.spreadsheet_unknown_files
        assert [info.path for info in vm.spreadsheet_protection_info] == [str(xlsx)]
        targets.append(vm.route_choices_result.targets)
    assert targets[0] == targets[1]
    vm.close()


def test_analysis_is_nonblocking_cached_and_rejects_stale_results(qtbot, monkeypatch, tmp_path):
    from docwen_gui.spreadsheet_protection import SpreadsheetProtectionInfo
    from docwen_gui.view_models.spreadsheet_analysis import SpreadsheetAnalysis

    first = tmp_path / "a.xlsx"
    second = tmp_path / "b.xlsx"
    first.write_bytes(b"a")
    second.write_bytes(b"b")
    started = threading.Event()
    release = threading.Event()
    calls = []

    def inspect(path):
        calls.append(path)
        if path == str(first):
            started.set()
            release.wait(3)
        return SpreadsheetProtectionInfo(path, "none")

    monkeypatch.setattr("docwen_gui.view_models.spreadsheet_analysis.inspect_xlsx_protection", inspect)
    model = SpreadsheetAnalysis()
    try:
        model.select((str(first),))
        qtbot.waitUntil(started.is_set)
        assert model.pending  # GUI event loop remains available while analysis waits.
        model.select((str(second),))
        qtbot.waitUntil(lambda: not model.pending)
        release.set()
        qtbot.wait(50)
        assert [info.path for info in model.results] == [str(second)]
        model.select((str(second),))
        assert not model.pending
        model.select((str(second),), refresh=True)
        qtbot.waitUntil(lambda: not model.pending)
        assert calls.count(str(second)) == 1
        second.write_bytes(b"changed")
        model.select((str(second),), refresh=True)
        qtbot.waitUntil(lambda: not model.pending)
        assert calls.count(str(second)) == 2
    finally:
        release.set()
        model.close()


def test_follow_system_survives_font_changes_and_ignores_system_when_explicit(qapp, monkeypatch):
    from docwen_gui.styles.theme_manager import ThemeManager

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    scheme = ["light"]
    monkeypatch.setattr(manager, "_resolve_system_theme", lambda name: scheme[0] if name == "system" else name)
    manager.initialize(qapp, "system")
    try:
        manager.apply_font_size_preset("large")
        scheme[0] = "dark"
        qapp.styleHints().colorSchemeChanged.emit(Qt.ColorScheme.Dark)
        assert manager.get_current_theme() == "dark"
        manager.apply_theme("light")
        qapp.styleHints().colorSchemeChanged.emit(Qt.ColorScheme.Dark)
        assert manager.get_current_theme() == "light"
    finally:
        manager.apply_font_size_preset("default")
        ThemeManager.reset_instance()
