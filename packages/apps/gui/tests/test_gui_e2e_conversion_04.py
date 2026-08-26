"""Focused tests split from test_gui_e2e_conversion.py."""

from __future__ import annotations

from ._gui_e2e_conversion_support import (
    _E2E_CONVERSION_TIMEOUT_MS,
    Path,
    _wait_for,
    pytest,
)


@pytest.mark.gui
class TestLayoutGuiExecution:
    def test_pdf_to_docx_runs_through_gui_thread_and_places_output(
        self,
        main_window_with_controller,
        tmp_path: Path,
        monkeypatch,
    ) -> None:
        import zipfile

        import fitz
        from PySide6.QtWidgets import QApplication

        import docwen_plugin_layout.to_document.converter as layout_converter
        from docwen_core.office_bridge import BridgeResult
        from docwen_gui.main_window import _normalize_path

        observed_priority: list[str] = []

        def fake_convert_with_backend_priority(
            input_path,
            output_path,
            *,
            source_format,
            backend_priority,
            com_candidates,
            libreoffice_format,
            cancel=None,
            failure_subject,
        ):
            del input_path, output_path, cancel, failure_subject
            observed_priority.extend(backend_priority)
            assert source_format == "pdf"
            assert set(com_candidates) == {"msoffice_word"}
            assert libreoffice_format == "docx"
            return BridgeResult(False, message="external PDF import skipped in GUI smoke")

        def fake_pdf2docx(
            input_path: Path,
            output_path: Path,
            *,
            cancellation=None,
        ) -> BridgeResult:
            del input_path
            assert cancellation is not None
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(output_path, "w") as archive:
                archive.writestr("word/document.xml", "<document/>")
            return BridgeResult(True, output_path=str(output_path), backend="pdf2docx")

        monkeypatch.setattr(layout_converter, "convert_with_backend_priority", fake_convert_with_backend_priority)
        monkeypatch.setattr(layout_converter, "_convert_pdf_with_pdf2docx", fake_pdf2docx)

        source = tmp_path / "layout-smoke.pdf"
        doc = fitz.open()
        page = doc.new_page(width=300, height=180)
        page.insert_text((48, 90), "GUI PDF to DOCX smoke")
        doc.save(source)
        doc.close()

        window = main_window_with_controller
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
        window._conversion_panel_vm.request_conversion("docx")

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
        assert observed_priority == ["msoffice_word", "libreoffice"]

        output_path = Path(entry.output_path)
        assert output_path.name == "layout-smoke.docx"
        assert output_path.exists()
        assert output_path.stat().st_size > 0
        with zipfile.ZipFile(output_path) as archive:
            assert "word/document.xml" in archive.namelist()
        assert window._info_area_vm.has_task_summary
        assert window._info_area_vm._task_summary.state == "success"

    def test_split_pdf_custom_page_runs_through_gui_thread_and_places_outputs(
        self,
        main_window_with_controller,
        tmp_path: Path,
    ) -> None:
        import fitz
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "contract.pdf"
        doc = fitz.open()
        for label in ("page 1", "page 2", "page 3"):
            page = doc.new_page(width=240, height=160)
            page.insert_text((48, 80), label)
        doc.save(source)
        doc.close()

        window = main_window_with_controller
        normalized = _normalize_path(str(source))
        window.view_model.add_files([str(source)])

        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        assert _wait_for(
            lambda: (
                window._conversion_panel._page_input_edit is not None
                and window._conversion_panel._split_pdf_button is not None
                and window._conversion_panel_vm.pdf_total_pages == 3
            ),
            timeout_ms=5000,
            interval_ms=20,
        )
        assert window._conversion_panel._page_input_edit is not None
        assert window._conversion_panel._split_pdf_button is not None
        window._conversion_panel._page_input_edit.setText("1")

        if app is not None:
            app.processEvents()

        assert window._conversion_panel._split_pdf_button.isEnabled()
        window._conversion_panel._split_pdf_button.click()

        if app is not None:
            app.processEvents()

        def _split_finished() -> bool:
            entry = window._batch_list_vm.get_file_entry(normalized)
            return entry is not None and entry.status in ("completed", "failed")

        assert _wait_for(_split_finished, timeout_ms=_E2E_CONVERSION_TIMEOUT_MS)
        entry = window._batch_list_vm.get_file_entry(normalized)
        assert entry is not None
        assert entry.status == "completed"
        assert entry.output_path

        part1 = Path(entry.output_path)
        part2 = part1.with_name("contract_part2.pdf")
        assert part1.name == "contract_part1.pdf"
        assert part1.exists()
        assert part2.exists()

        part1_doc = fitz.open(part1)
        part2_doc = fitz.open(part2)
        try:
            assert part1_doc.page_count == 1
            assert part2_doc.page_count == 2
        finally:
            part1_doc.close()
            part2_doc.close()
