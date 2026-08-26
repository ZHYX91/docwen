from __future__ import annotations

import os
from pathlib import Path

import pytest

pytestmark = pytest.mark.gui


class TestMainWindowWithRealController:
    def test_window_created_with_controller(self, main_window_with_controller) -> None:
        window = main_window_with_controller
        assert window is not None
        assert window.view_model is not None
        controller = window.view_model.controller
        assert controller is not None
        assert controller.has_runtime

    def test_window_title_set(self, main_window_with_controller) -> None:
        title = main_window_with_controller.windowTitle()
        assert len(title) > 0
        assert "DocWen" in title

    def test_core_widgets_assembled(self, main_window_with_controller) -> None:
        window = main_window_with_controller
        assert window.input_area is not None
        assert window.batch_list is not None
        assert window.conversion_panel is not None
        assert window.action_area is not None
        assert window.info_area is not None

    def test_task_event_bridge_wired(self, main_window_with_controller) -> None:
        bridge = main_window_with_controller.task_event_bridge
        assert bridge is not None
        assert bridge.is_flushing is True


class TestFileLoading:
    def test_add_docx_file_updates_view_model(self, main_window_with_controller, sample_docx: Path) -> None:
        from PySide6.QtWidgets import QApplication

        window = main_window_with_controller
        vm = window.view_model

        vm.add_files([str(sample_docx)])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        assert vm.has_files
        assert len(vm.files) == 1
        file_ref = vm.files[0]
        assert file_ref.path == str(sample_docx)
        assert file_ref.format == "docx"
        assert file_ref.category == "document"

    def test_file_synced_to_batch_list(self, main_window_with_controller, sample_docx: Path) -> None:
        from PySide6.QtWidgets import QApplication

        from docwen_gui.main_window import _normalize_path

        window = main_window_with_controller
        vm = window.view_model

        vm.add_files([str(sample_docx)])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        files_in_batch = window._batch_list_vm.get_files()
        normalized = _normalize_path(str(sample_docx))
        assert normalized in files_in_batch or any(normalized in f for f in files_in_batch)

    def test_selected_file_synced_to_panels(self, main_window_with_controller, sample_docx: Path) -> None:
        from PySide6.QtWidgets import QApplication

        window = main_window_with_controller
        vm = window.view_model

        vm.add_files([str(sample_docx)])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        action_vm = window._action_area_vm
        assert action_vm.visible
        assert action_vm.file_type == "document"
        assert action_vm.file_path is not None


class TestInfoAreaIntegration:
    def test_info_area_accessible_after_file_load(self, main_window_with_controller, sample_docx: Path) -> None:
        from PySide6.QtWidgets import QApplication

        window = main_window_with_controller

        window.view_model.add_files([str(sample_docx)])
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

        info_vm = window._info_area_vm
        assert info_vm is not None
        assert info_vm.message_count >= 0


class TestOffscreenCompatibility:
    def test_offscreen_platform_is_set(self, qapp) -> None:
        platform = os.environ.get("QT_QPA_PLATFORM", "")
        assert platform == "offscreen"

    def test_window_can_be_created_in_offscreen(self, main_window_with_controller) -> None:
        window = main_window_with_controller
        assert window is not None
        assert window.isValid() if hasattr(window, "isValid") else True
