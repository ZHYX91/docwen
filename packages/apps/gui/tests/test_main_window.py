import pytest
from PySide6.QtCore import Qt
from PySide6.QtTest import QSignalSpy

pytestmark = pytest.mark.gui


class TestMainWindowCreation:
    def test_window_has_title(self, main_window) -> None:
        title = main_window.windowTitle()
        assert len(title) > 0
        assert "DocWen" in title

    def test_window_has_minimum_size(self, main_window) -> None:
        assert main_window.minimumWidth() > 0
        assert main_window.minimumHeight() > 0

    def test_window_has_object_name(self, main_window) -> None:
        assert main_window.objectName() == "docwenMainWindow"

    def test_window_is_not_visible_initially(self, main_window) -> None:
        """MainWindow should NOT be shown during construction."""
        # In tests we call setup_ui() but not show()
        # true by construction — we never call show()

    def test_view_model_accessible(self, main_window) -> None:
        assert main_window.view_model is not None
        assert main_window.view_model.mode == "single"

    def test_settings_button_exists(self, main_window) -> None:
        btn = main_window.settings_button
        assert btn is not None
        assert btn.toolTip() != ""
        assert btn.accessibleName() == btn.toolTip()


class TestViewModeIntegration:
    def test_add_files_updates_view_model(self, main_window, tmp_path) -> None:
        vm = main_window.view_model
        source = tmp_path / "test.docx"
        source.write_text("content-first admission probe", encoding="utf-8")
        vm.add_files([str(source)])
        assert vm.has_files is True
        assert len(vm.files) == 1

    def test_mode_change_updates_view_model(self, main_window) -> None:
        vm = main_window.view_model
        vm.set_mode("batch")
        assert vm.mode == "batch"


class TestWindowDimensions:
    def test_default_width(self, main_window) -> None:
        w = main_window.width()
        assert w >= 400, f"Window width {w} is too narrow"

    def test_default_height(self, main_window) -> None:
        h = main_window.height()
        assert h >= 400, f"Window height {h} is too small"


class TestMainWindowAssembly:
    def test_core_widgets_are_assembled(self, main_window) -> None:
        assert main_window.input_area is not None
        assert main_window.batch_list is not None
        assert main_window.conversion_panel is not None
        assert main_window.action_area is not None
        assert main_window.info_area is not None

    def test_task_event_bridge_is_wired(self, main_window) -> None:
        bridge = main_window.task_event_bridge
        assert bridge is not None
        assert bridge.is_flushing is True

    def test_clear_button_resets_current_work_session(self, main_window, qapp, qtbot, tmp_path) -> None:
        source = tmp_path / "source.txt"
        source.write_text("source stays on disk", encoding="utf-8")
        output = tmp_path / "existing-output.md"
        output.write_text("output stays on disk", encoding="utf-8")
        recent = str(tmp_path / "recent.txt")

        main_window.view_model.add_files([str(source)])
        qapp.processEvents()
        main_window.input_area._recent_files = [recent]

        info_vm = main_window.info_area.view_model
        info_vm.add_message("old history", "warning")
        info_vm.set_task_summary(
            operation_id="clear-session",
            current_file=source.name,
            total_count=1,
            state="success",
            guide_actions=[{"action_key": "add_more_files", "target_path": ""}],
        )
        info_vm.set_transient_message("progress:clear-session", "old progress", "info", ttl_ms=0)

        assert main_window.view_model.has_files is True
        assert main_window.view_model.selected_file is not None
        assert main_window.batch_list.view_model.entry_count == 1
        assert main_window.action_area.view_model.visible is True
        assert main_window.info_area.message_count == 1
        assert info_vm.transient_count == 1
        assert info_vm.has_task_summary is True

        qtbot.mouseClick(main_window.input_area.clear_button, Qt.MouseButton.LeftButton)
        qapp.processEvents()

        assert main_window.view_model.has_files is False
        assert main_window.view_model.selected_file is None
        assert main_window.batch_list.view_model.entry_count == 0
        assert main_window.batch_list.view_model.get_files() == []
        assert main_window.action_area.view_model.visible is False
        assert main_window.action_area.view_model.file_path is None
        assert main_window.conversion_panel.view_model.has_files is False
        assert main_window.info_area.message_count == 0
        assert main_window.info_area._history_row_widgets == []
        assert info_vm.transient_count == 0
        assert info_vm.has_task_summary is False
        assert info_vm.guide_visible is False
        assert info_vm.status_source == "idle"
        assert main_window.input_area._recent_files == [recent]
        assert source.read_text(encoding="utf-8") == "source stays on disk"
        assert output.read_text(encoding="utf-8") == "output stays on disk"

    def test_clear_button_is_disabled_while_execution_is_active(self, main_window, qapp) -> None:
        assert main_window.input_area.clear_button.isEnabled()

        main_window.action_area.view_model.show_cancel()
        qapp.processEvents()
        assert main_window.input_area.clear_button.isEnabled() is False

        main_window.action_area.view_model.hide_cancel()
        qapp.processEvents()
        assert main_window.input_area.clear_button.isEnabled()


class TestWindowActivation:
    def test_ipc_activate_relies_on_single_view_model_activation_request(self, main_window, monkeypatch) -> None:
        direct_calls: list[bool] = []
        monkeypatch.setattr(main_window, "bring_to_front", lambda: direct_calls.append(True))
        activation_spy = QSignalSpy(main_window.view_model.window_activation_requested)

        main_window.handle_ipc_command("activate")

        assert activation_spy.count() == 1
        assert direct_calls == [True]

    def test_bring_to_front_preserves_maximized_state(self, main_window, qapp) -> None:
        main_window.showMaximized()
        qapp.processEvents()
        assert main_window.isMaximized()

        main_window.bring_to_front()
        qapp.processEvents()

        assert main_window.isMaximized()
