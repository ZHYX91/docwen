"""Focused tests split from test_gui_e2e_conversion.py."""

from __future__ import annotations

from ._gui_e2e_conversion_support import (
    _E2E_CONVERSION_TIMEOUT_MS,
    BatchFileEntry,
    Path,
    _wait_for,
    pytest,
)


@pytest.mark.gui
class TestAggregateGuiExecution:
    @staticmethod
    def _wait_for_entries_finished(window, paths: list[str]) -> None:
        from docwen_gui.main_window import _normalize_path

        normalized = [_normalize_path(path) for path in paths]

        def _finished() -> bool:
            entries = [window._batch_list_vm.get_file_entry(path) for path in normalized]
            if any(entry is None for entry in entries):
                return False
            return all(entry.status in ("completed", "failed") for entry in entries if entry is not None)

        assert _wait_for(_finished, timeout_ms=_E2E_CONVERSION_TIMEOUT_MS)

    @staticmethod
    def _completed_entries(window, paths: list[str]) -> list[BatchFileEntry]:
        from docwen_gui.main_window import _normalize_path

        entries = [window._batch_list_vm.get_file_entry(_normalize_path(path)) for path in paths]
        assert all(entry is not None and entry.status == "completed" for entry in entries)
        return [entry for entry in entries if entry is not None]

    def test_merge_images_to_tiff_failure_marks_all_participants_failed(
        self, main_window_with_controller, tmp_path: Path
    ) -> None:
        from PIL import Image
        from PySide6.QtWidgets import QApplication

        good = tmp_path / "good.png"
        broken = tmp_path / "broken.png"

        Image.new("RGB", (32, 24), (255, 0, 0)).save(good)
        broken.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 64)

        window = main_window_with_controller
        window.view_model.add_files([str(good), str(broken)])

        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        assert window._conversion_panel._merge_tiff_button is not None
        window._conversion_panel._merge_tiff_button.click()

        if app is not None:
            app.processEvents()

        self._wait_for_entries_finished(window, [str(good), str(broken)])
        entries = [
            window._batch_list_vm.get_file_entry(path)
            for path in (str(good).replace("\\", "/"), str(broken).replace("\\", "/"))
        ]
        assert all(entry is not None and entry.status == "failed" for entry in entries)
        assert all(entry is not None and not entry.output_path for entry in entries)
        assert all(entry is not None and entry.error_message for entry in entries)
        assert any("broken.png" in (entry.error_message or "") for entry in entries if entry is not None)
        assert any("Failed to load image" in (entry.error_message or "") for entry in entries if entry is not None)

        assert window._info_area_vm.has_task_summary
        assert window._info_area_vm.history_rows
        latest = window._info_area_vm.history_rows[-1]
        assert latest.message_type == "danger"
        assert "broken.png" in latest.message
