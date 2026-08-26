"""Focused tests split from test_gui_e2e_conversion.py."""

from __future__ import annotations

from ._gui_e2e_conversion_support import (
    _E2E_CONVERSION_TIMEOUT_MS,
    Path,
    _has_document_pdf_backend,
    _has_spreadsheet_pdf_backend,
    _read_pdf_text,
    _wait_for,
    os,
    pytest,
    shutdown_main_window,
    t,
)


@pytest.mark.gui
class TestDocxToMdConversion:
    def test_conversion_completes_successfully(self, main_window_with_controller, sample_docx: Path) -> None:
        from PySide6.QtTest import QSignalSpy
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import _normalize_path

        window = main_window_with_controller
        vm = window.view_model
        file_path_str = str(sample_docx)
        normalized = _normalize_path(file_path_str)

        vm.add_files([file_path_str])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        assert vm.has_files
        assert window._action_area_vm.visible

        controller = vm.controller
        assert controller is not None
        assert controller.has_runtime

        spy = QSignalSpy(vm.task_summary_changed)
        window._action_area_vm.request_conversion("md")

        if app is not None:
            app.processEvents()

        def _conversion_finished() -> bool:
            entry = window._batch_list_vm.get_file_entry(normalized)
            if entry is None:
                return False
            return entry.status in ("completed", "failed")

        assert _wait_for(_conversion_finished, timeout_ms=_E2E_CONVERSION_TIMEOUT_MS)

        if app is not None:
            app.processEvents()

        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "completed"
        assert entry.output_path

        output_path = Path(entry.output_path)
        assert output_path.exists()
        assert output_path.stat().st_size > 0
        content = output_path.read_text(encoding="utf-8")
        assert "E2E Test Document" in content
        assert "# " in content

        assert _wait_for(lambda: spy.count() >= 1, timeout_ms=2000, interval_ms=20)
        spy_count = spy.count()
        assert spy_count >= 1
        summary = spy.at(spy_count - 1)[0]
        assert isinstance(summary, dict)
        assert summary.get("status") == "completed"
        assert window._action_area_vm.cancel_visible is False
        assert window._info_area_vm.has_task_summary
        assert window._info_area_vm._task_summary.state == "success"
        assert window._info_area_vm._task_summary.completed_count == 1
        assert window._info_area_vm._task_summary.failed_count == 0
        assert window._info_area_vm.history_rows
        latest = window._info_area_vm.history_rows[-1]
        assert latest.message_type == "success"
        assert latest.show_location is True
        assert latest.navigate_file_path == str(output_path)
        assert latest.message == t("info_area.history_completed", name=sample_docx.name)
        guide_keys = {action["action_key"] for action in window._info_area_vm.guide_actions}
        assert {"open_output_dir", "add_more_files"}.issubset(guide_keys)

    def test_conversion_uses_custom_output_directory_setting(
        self, main_window_with_controller, sample_docx: Path, tmp_path: Path
    ) -> None:
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import _normalize_path

        window = main_window_with_controller
        controller = window.view_model.controller
        assert controller is not None
        assert controller.config_port is not None
        output_dir = tmp_path / "exports"
        assert controller.config_port.set("output.directory.mode", "custom")
        assert controller.config_port.set("output.directory.custom_path", str(output_dir))
        assert controller.config_port.set("output.directory.create_date_subfolder", False)

        file_path_str = str(sample_docx)
        normalized = _normalize_path(file_path_str)
        window.view_model.add_files([file_path_str])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        window._action_area_vm.request_conversion("md")
        if app is not None:
            app.processEvents()

        def _conversion_finished() -> bool:
            entry = window._batch_list_vm.get_file_entry(normalized)
            return entry is not None and (entry.status == "failed" or entry.output_path is not None)

        assert _wait_for(_conversion_finished, timeout_ms=_E2E_CONVERSION_TIMEOUT_MS)

        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "completed"
        assert entry.output_path
        output_path = Path(entry.output_path)
        assert output_path.exists()
        assert output_path.parent.parent == output_dir
        assert output_path.parent.name == output_path.stem
        assert (output_path.parent / "docwen-node.json").is_file()
        assert "E2E Test Document" in output_path.read_text(encoding="utf-8")

    def test_conversion_preserves_unowned_legacy_output(
        self, main_window_with_controller, sample_docx: Path, tmp_path: Path
    ) -> None:
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import _normalize_path

        window = main_window_with_controller
        controller = window.view_model.controller
        assert controller is not None
        assert controller.config_port is not None
        output_dir = tmp_path / "exports"
        output_dir.mkdir()
        existing = output_dir / "e2e_test.md"
        existing.write_text("existing GUI result", encoding="utf-8")

        assert controller.config_port.set("output.directory.mode", "custom")
        assert controller.config_port.set("output.directory.custom_path", str(output_dir))
        assert controller.config_port.set("output.directory.create_date_subfolder", False)

        file_path_str = str(sample_docx)
        normalized = _normalize_path(file_path_str)
        window.view_model.add_files([file_path_str])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        window._action_area_vm.request_conversion("md")
        if app is not None:
            app.processEvents()

        def _conversion_finished() -> bool:
            entry = window._batch_list_vm.get_file_entry(normalized)
            return entry is not None and (entry.status == "failed" or entry.output_path is not None)

        assert _wait_for(_conversion_finished, timeout_ms=_E2E_CONVERSION_TIMEOUT_MS)

        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "completed"
        assert entry.output_path
        output_path = Path(entry.output_path)
        assert output_path.exists()
        assert output_path.parent.parent == output_dir
        assert output_path.parent.name == output_path.stem
        assert (output_path.parent / "docwen-node.json").is_file()
        assert existing.read_text(encoding="utf-8") == "existing GUI result"
        assert "E2E Test Document" in output_path.read_text(encoding="utf-8")

    def test_conversion_rejects_when_no_runtime(self, qapp, sample_docx: Path) -> None:
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import MainWindow, _normalize_path
        from docwen_gui.view_models.main_window_vm import MainWindowViewModel

        vm = MainWindowViewModel(controller=None)
        window = MainWindow(view_model=vm)
        window.setup_ui()
        try:
            vm.add_files([str(sample_docx)])
            app = QApplication.instance()
            if app is not None:
                app.processEvents()

            window._action_area_vm.request_conversion("md")

            if app is not None:
                app.processEvents()

            assert len(window._active_threads) == 0

            normalized = _normalize_path(str(sample_docx))
            entry = window._batch_list_vm.get_file_entry(normalized)
            if entry is not None:
                assert entry.status in ("pending", "")
        finally:
            shutdown_main_window(window)

    def test_clear_files_before_conversion(self, main_window_with_controller, sample_docx: Path) -> None:
        from PySide6.QtWidgets import QApplication

        window = main_window_with_controller
        vm = window.view_model

        vm.add_files([str(sample_docx)])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()
        assert vm.has_files

        vm.clear_files()
        if app is not None:
            app.processEvents()

        assert not vm.has_files
        assert vm.selected_file is None
        assert not window._action_area_vm.visible

    def test_multiple_files_loaded(self, main_window_with_controller, sample_docx: Path, tmp_path: Path) -> None:
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import _normalize_path

        window = main_window_with_controller
        vm = window.view_model

        docx2_path = tmp_path / "e2e_test_2.docx"
        docx2_path.write_bytes(sample_docx.read_bytes())

        vm.add_files([str(sample_docx), str(docx2_path)])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        assert len(vm.files) == 2
        assert vm.selected_file is not None

        normalized1 = _normalize_path(str(sample_docx))
        normalized2 = _normalize_path(str(docx2_path))
        files_in_batch = window._batch_list_vm.get_files()
        assert any(normalized1 in f for f in files_in_batch)
        assert any(normalized2 in f for f in files_in_batch)


@pytest.mark.gui_smoke
class TestOfficeBackedGuiExecution:
    @pytest.mark.slow
    def test_docx_to_pdf_runs_through_gui_thread_with_external_office_backend(
        self,
        main_window_with_controller,
        sample_docx: Path,
        tmp_path: Path,
    ) -> None:
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import _normalize_path

        if os.environ.get("DOCWEN_RUN_EXTERNAL_OFFICE_GUI_SMOKE") != "1":
            pytest.skip("Set DOCWEN_RUN_EXTERNAL_OFFICE_GUI_SMOKE=1 to run the external Office GUI smoke")
        if not _has_document_pdf_backend():
            pytest.skip("No Word/WPS Writer COM or LibreOffice backend available for DOCX->PDF GUI smoke")

        window = main_window_with_controller
        controller = window.view_model.controller
        assert controller is not None
        assert controller.config_port is not None
        output_dir = tmp_path / "office_gui_exports"
        assert controller.config_port.set("output.directory.mode", "custom")
        assert controller.config_port.set("output.directory.custom_path", str(output_dir))
        assert controller.config_port.set("output.directory.create_date_subfolder", False)

        normalized = _normalize_path(str(sample_docx))
        window.view_model.add_files([str(sample_docx)])

        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        assert _wait_for(
            lambda: window._conversion_panel_vm.current_file_path == str(sample_docx),
            timeout_ms=5000,
            interval_ms=20,
        )
        window._conversion_panel_vm.request_conversion("pdf")

        if app is not None:
            app.processEvents()

        def _conversion_finished() -> bool:
            entry = window._batch_list_vm.get_file_entry(normalized)
            return entry is not None and entry.status in ("completed", "failed")

        assert _wait_for(_conversion_finished, timeout_ms=_E2E_CONVERSION_TIMEOUT_MS)
        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "completed"
        assert entry.output_path

        output_path = Path(entry.output_path)
        assert output_path.name == "e2e_test.pdf"
        assert output_path.parent == output_dir
        assert output_path.exists()
        assert output_path.stat().st_size > 0
        assert output_path.read_bytes().startswith(b"%PDF")
        output_text = _read_pdf_text(output_path)
        assert "E2E Test Document" in output_text
        assert "conversion verification" in output_text
        assert window._info_area_vm.has_task_summary
        assert window._info_area_vm._task_summary.state == "success"

    @pytest.mark.slow
    def test_xlsx_to_pdf_runs_through_gui_thread_with_external_office_backend(
        self,
        main_window_with_controller,
        tmp_path: Path,
    ) -> None:
        from openpyxl import Workbook
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import _normalize_path

        if os.environ.get("DOCWEN_RUN_EXTERNAL_OFFICE_GUI_SMOKE") != "1":
            pytest.skip("Set DOCWEN_RUN_EXTERNAL_OFFICE_GUI_SMOKE=1 to run the external Office GUI smoke")
        if not _has_spreadsheet_pdf_backend():
            pytest.skip("No Excel/WPS Spreadsheet COM or LibreOffice backend available for XLSX->PDF GUI smoke")

        source = tmp_path / "office-sheet.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Smoke"
        sheet.append(["Name", "Score"])
        sheet.append(["Alice", 95])
        sheet.append(["Bob", 88])
        workbook.save(source)
        workbook.close()

        window = main_window_with_controller
        controller = window.view_model.controller
        assert controller is not None
        assert controller.config_port is not None
        output_dir = tmp_path / "spreadsheet_gui_exports"
        assert controller.config_port.set("output.directory.mode", "custom")
        assert controller.config_port.set("output.directory.custom_path", str(output_dir))
        assert controller.config_port.set("output.directory.create_date_subfolder", False)

        normalized = _normalize_path(str(source))
        window.view_model.add_files([str(source)])

        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        assert _wait_for(
            lambda: window._conversion_panel_vm.current_file_path == str(source),
            timeout_ms=5000,
            interval_ms=20,
        )
        window._conversion_panel_vm.request_conversion("pdf")

        if app is not None:
            app.processEvents()

        def _conversion_finished() -> bool:
            entry = window._batch_list_vm.get_file_entry(normalized)
            return entry is not None and entry.status in ("completed", "failed")

        assert _wait_for(_conversion_finished, timeout_ms=_E2E_CONVERSION_TIMEOUT_MS)
        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "completed"
        assert entry.output_path

        output_path = Path(entry.output_path)
        assert output_path.name == "office-sheet.pdf"
        assert output_path.parent == output_dir
        assert output_path.exists()
        assert output_path.stat().st_size > 0
        assert output_path.read_bytes().startswith(b"%PDF")
        output_text = _read_pdf_text(output_path)
        assert all(token in output_text for token in ("Name", "Score", "Alice", "95", "Bob", "88"))
        assert window._info_area_vm.has_task_summary
        assert window._info_area_vm._task_summary.state == "success"

    @pytest.mark.slow
    def test_markdown_to_pdf_runs_through_gui_thread_with_external_office_backend(
        self,
        main_window_with_controller,
        tmp_path: Path,
    ) -> None:
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import _normalize_path

        if os.environ.get("DOCWEN_RUN_EXTERNAL_OFFICE_GUI_SMOKE") != "1":
            pytest.skip("Set DOCWEN_RUN_EXTERNAL_OFFICE_GUI_SMOKE=1 to run the external Office GUI smoke")
        if not _has_document_pdf_backend():
            pytest.skip("No Word/WPS Writer COM or LibreOffice backend available for Markdown->PDF GUI smoke")

        source = tmp_path / "office-markdown.md"
        source.write_text(
            "# Office Markdown Smoke\n\n"
            "This file exercises the GUI Markdown to PDF route through the current runtime.\n",
            encoding="utf-8",
        )

        window = main_window_with_controller
        controller = window.view_model.controller
        assert controller is not None
        assert controller.config_port is not None
        output_dir = tmp_path / "markdown_gui_exports"
        assert controller.config_port.set("output.directory.mode", "custom")
        assert controller.config_port.set("output.directory.custom_path", str(output_dir))
        assert controller.config_port.set("output.directory.create_date_subfolder", False)

        normalized = _normalize_path(str(source))
        window.view_model.add_files([str(source)])

        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        assert _wait_for(lambda: window._action_area_vm.visible, timeout_ms=5000, interval_ms=20)
        window._action_area_vm.request_conversion("pdf")

        if app is not None:
            app.processEvents()

        def _conversion_finished() -> bool:
            entry = window._batch_list_vm.get_file_entry(normalized)
            return entry is not None and entry.status in ("completed", "failed")

        assert _wait_for(_conversion_finished, timeout_ms=_E2E_CONVERSION_TIMEOUT_MS)
        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "completed"
        assert entry.output_path

        output_path = Path(entry.output_path)
        assert output_path.name == "office-markdown.pdf"
        assert output_path.parent == output_dir
        assert output_path.exists()
        assert output_path.stat().st_size > 0
        assert output_path.read_bytes().startswith(b"%PDF")
        output_text = _read_pdf_text(output_path)
        assert "Office Markdown Smoke" in output_text
        assert "current runtime" in output_text
        assert window._info_area_vm.has_task_summary
        assert window._info_area_vm._task_summary.state == "success"
