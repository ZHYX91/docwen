"""Focused tests split from test_main_window_features.py."""

from __future__ import annotations

from ._main_window_features_support import (
    pytest,
)

pytestmark = pytest.mark.gui


class TestAlwaysOnTop:
    """Verify always-on-top toggle behavior."""

    def test_default_is_disabled(self, main_window) -> None:
        assert main_window.is_always_on_top_enabled() is False
        assert main_window._always_on_top_enabled is False

    def test_set_enabled(self, main_window) -> None:
        main_window.set_window_always_on_top(True)
        assert main_window._always_on_top_enabled is True
        assert main_window.is_always_on_top_enabled() is True
        # Clean up
        main_window.set_window_always_on_top(False)

    def test_set_disabled(self, main_window) -> None:
        main_window.set_window_always_on_top(True)
        main_window.set_window_always_on_top(False)
        assert main_window._always_on_top_enabled is False
        assert main_window.is_always_on_top_enabled() is False

    def test_toggle_from_disabled(self, main_window) -> None:
        main_window.set_window_always_on_top(False)
        main_window.toggle_always_on_top()
        assert main_window._always_on_top_enabled is True
        # Clean up
        main_window.set_window_always_on_top(False)

    def test_toggle_from_enabled(self, main_window) -> None:
        main_window.set_window_always_on_top(True)
        main_window.toggle_always_on_top()
        assert main_window._always_on_top_enabled is False

    def test_property_set(self, main_window) -> None:
        main_window.set_window_always_on_top(True)
        assert main_window.property("windowAlwaysOnTop") is True
        main_window.set_window_always_on_top(False)
        assert main_window.property("windowAlwaysOnTop") is False


class TestUserVisibleParitySmoke:
    """Smoke-test user-visible GUI parity states without running conversions."""

    def _window_with_config(self, qapp, values: dict[str, object] | None = None):
        from docwen_gui.main_window import MainWindow
        from docwen_gui.view_models.main_window_vm import MainWindowViewModel

        class _Cfg:
            def __init__(self) -> None:
                self.values = dict(values or {})
                self.calls: list[tuple[str, object]] = []
                self.get_calls: list[str] = []

            def get(self, key: str, default=None):
                self.get_calls.append(key)
                return self.values.get(key, default)

            def set(self, key: str, value) -> bool:
                self.calls.append((key, value))
                self.values[key] = value
                return True

        class _Controller:
            def __init__(self) -> None:
                self.config_port = _Cfg()

            def stop(self) -> None:
                return None

        ctrl = _Controller()
        vm = MainWindowViewModel(controller=ctrl)  # type: ignore[arg-type]
        window = MainWindow(view_model=vm)
        window.setup_ui()
        return window, ctrl

    def test_window_geometry_save_respects_disable_env(self, qapp, monkeypatch: pytest.MonkeyPatch) -> None:
        window, ctrl = self._window_with_config(qapp)
        try:
            window.move(111, 222)
            window.resize(777, 666)
            monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")

            window._save_gui_state()

            assert ctrl.config_port.calls == []
        finally:
            window.close()

    def test_window_geometry_restore_reads_config_values(self, qapp) -> None:
        window, ctrl = self._window_with_config(
            qapp,
            {
                "gui.window.geometry_schema_version": 2,
                "gui.window.center_panel_screen_x": 123,
                "gui.window.window_y": 20,
                "gui.window.default_width": 640,
                "gui.window.default_height": 760,
            },
        )
        try:
            assert {
                "gui.window.geometry_schema_version",
                "gui.window.center_panel_screen_x",
                "gui.window.window_y",
                "gui.window.default_width",
                "gui.window.default_height",
            }.issubset(set(ctrl.config_port.get_calls))
            assert window.pos().x() + window._center_column_offset() == 123
            assert window.pos().y() == 20
            assert window.size().width() == 640
            assert window.size().height() == 760
        finally:
            window.close()

    def test_system_tray_setup_respects_disabled_default(
        self,
        qapp,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        monkeypatch.setattr(
            "docwen_gui.main_window.QSystemTrayIcon.isSystemTrayAvailable",
            staticmethod(lambda: True),
        )

        window, ctrl = self._window_with_config(qapp)
        try:
            assert "gui.notifications.system_tray" in ctrl.config_port.get_calls
            assert window._system_tray_icon is None
        finally:
            window.close()

    def test_system_tray_setup_creates_icon_when_enabled(
        self,
        qapp,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        monkeypatch.setattr(
            "docwen_gui.main_window.QSystemTrayIcon.isSystemTrayAvailable",
            staticmethod(lambda: True),
        )

        window, _ctrl = self._window_with_config(
            qapp,
            {"gui.notifications.system_tray": True},
        )
        try:
            assert window._system_tray_icon is not None
            assert window._system_tray_icon.isVisible()
        finally:
            window.close()

    def test_task_completion_notification_respects_min_elapsed_and_tray_config(
        self,
        qapp,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import time

        window, _ctrl = self._window_with_config(
            qapp,
            {
                "gui.notifications.task_completion": True,
                "gui.notifications.min_elapsed_seconds": 5,
                "gui.notifications.system_tray": True,
                "gui.notifications.system_tray_timeout_ms": 1500,
            },
        )
        messages: list[tuple[str, str, int]] = []

        class _Tray:
            def showMessage(self, title, body, icon, timeout_ms):
                messages.append((title, body, timeout_ms))

            def hide(self):
                return None

        try:
            monkeypatch.setattr(
                "docwen_gui.main_window.QSystemTrayIcon.isSystemTrayAvailable",
                staticmethod(lambda: True),
            )
            window._system_tray_icon = _Tray()  # type: ignore[assignment]
            window._start_time = time.monotonic()

            window._maybe_notify_task_completion({"display_name": "sample.docx"})
            assert messages == []

            window._start_time = time.monotonic() - 6
            window._maybe_notify_task_completion({"display_name": "sample.docx"})

            assert messages
            assert messages[0][0]
            assert messages[0][1]
            assert messages[0][2] == 1500
        finally:
            window.close()

    def test_task_completion_notification_uses_task_summary_state(
        self,
        qapp,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        from docwen_gui.i18n import get_locale, set_locale

        previous_locale = get_locale()
        set_locale("en_US")
        window, _ctrl = self._window_with_config(
            qapp,
            {
                "gui.notifications.task_completion": True,
                "gui.notifications.min_elapsed_seconds": 0,
                "gui.notifications.system_tray": True,
                "gui.notifications.system_tray_timeout_ms": 1500,
            },
        )
        messages: list[tuple[str, str, int]] = []

        class _Tray:
            def showMessage(self, title, body, icon, timeout_ms):
                messages.append((title, body, timeout_ms))

            def hide(self):
                return None

        try:
            monkeypatch.setattr(
                "docwen_gui.main_window.QSystemTrayIcon.isSystemTrayAvailable",
                staticmethod(lambda: True),
            )
            window._system_tray_icon = _Tray()  # type: ignore[assignment]
            window._info_area_vm.set_task_summary(
                operation_id="failed-op",
                current_file="failed.docx",
                completed_count=0,
                total_count=2,
                failed_count=2,
                state="failed",
                tone="danger",
            )

            window._maybe_notify_task_completion({"display_name": "failed.docx"})

            assert messages
            assert "failed" in messages[-1][1]
            assert "Task completed" not in messages[-1][1]

            window._info_area_vm.set_task_summary(
                operation_id="partial-op",
                current_file="mixed-batch",
                completed_count=2,
                total_count=3,
                failed_count=1,
                state="partial",
                tone="warning",
            )

            window._maybe_notify_task_completion({"display_name": "mixed-batch"})

            assert "1 failed" in messages[-1][1]
            assert "Task completed" not in messages[-1][1]

            window._info_area_vm.set_task_summary(
                operation_id="cancelled-op",
                current_file="cancelled.docx",
                completed_count=0,
                total_count=1,
                failed_count=0,
                cancelled_count=1,
                state="cancelled",
                tone="warning",
            )

            window._maybe_notify_task_completion({"display_name": "cancelled.docx"})

            assert "cancelled" in messages[-1][1]
            assert "Task completed" not in messages[-1][1]
        finally:
            set_locale(previous_locale)
            window.close()
