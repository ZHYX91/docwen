"""Focused tests split from test_entrypoints.py."""

from __future__ import annotations

from ._entrypoints_support import (
    Any,
    Callable,
    SimpleNamespace,
    TaskEvent,
    _Signal,
    cast,
    pytest,
    run_subprocess,
)

pytestmark = pytest.mark.unit


class TestGuiEntry:
    def test_make_event_callback_adapts_runtime_events(self) -> None:
        from docwen_bundle.gui_entry import _make_event_callback

        received: list[tuple[str, dict[str, object]]] = []
        callback = _make_event_callback(lambda event_type, payload: received.append((event_type, payload)))
        event = TaskEvent(
            event_type="task_progress",
            task_id="task-1",
            sequence=1,
            payload={"percent": 50.0},
        )

        callback(event)

        assert received == [("task_progress", {"task_id": "task-1", "percent": 50.0})]

    def test_gui_entry_returns_bootstrap_exit_code(self, monkeypatch: pytest.MonkeyPatch) -> None:
        import docwen_bundle.gui_bootstrap as gui_bootstrap_module
        import docwen_bundle.gui_entry as gui_entry
        import docwen_runtime.logging as runtime_logging

        pre_init_levels: list[str] = []
        monkeypatch.setattr(runtime_logging, "pre_init_logging", pre_init_levels.append)

        monkeypatch.setattr(
            gui_bootstrap_module,
            "bootstrap_gui",
            lambda app_name, argv: SimpleNamespace(
                should_exit=True,
                exit_code=9,
                files_to_add=[],
                instance_lock=None,
                ipc_dir=None,
            ),
        )

        assert gui_entry.main(["docwen-gui"]) == 9
        assert pre_init_levels == ["INFO"]

    @pytest.mark.parametrize("missing_module", ["PySide6.QtWidgets", "qfluentwidgets"])
    def test_gui_entry_reports_required_dependency_without_traceback(
        self,
        monkeypatch: pytest.MonkeyPatch,
        capsys: pytest.CaptureFixture[str],
        missing_module: str,
    ) -> None:
        import docwen_bundle.gui_entry as gui_entry

        def _raise_missing(_argv=None) -> int:
            raise ModuleNotFoundError(
                f"No module named {missing_module!r}",
                name=missing_module,
            )

        monkeypatch.setattr(gui_entry, "_main_with_guard_active", _raise_missing)

        exit_code = gui_entry.main(["docwen-gui"])

        captured = capsys.readouterr()
        assert exit_code == 4
        assert missing_module.partition(".")[0] in captured.err
        assert "required dependency" in captured.err
        assert "Traceback" not in captured.err

    @pytest.mark.parametrize("missing_root", ["PySide6", "qfluentwidgets"])
    def test_gui_entry_reports_required_dependency_from_fresh_interpreter(
        self,
        missing_root: str,
    ) -> None:
        import sys
        import textwrap

        script = textwrap.dedent(
            """
            import importlib.abc
            import sys

            missing_root = sys.argv[1]

            class _BlockRequiredGuiDependency(importlib.abc.MetaPathFinder):
                def find_spec(self, fullname, path=None, target=None):
                    if fullname == missing_root or fullname.startswith(f"{missing_root}."):
                        raise ModuleNotFoundError(
                            f"No module named {fullname!r}",
                            name=fullname,
                        )
                    return None

            sys.meta_path.insert(0, _BlockRequiredGuiDependency())

            from docwen_bundle.gui_entry import main

            raise SystemExit(main(["docwen-gui"]))
            """
        )

        completed = run_subprocess(
            [sys.executable, "-c", script, missing_root],
            check=False,
        )

        assert isinstance(completed.stdout, str)
        assert isinstance(completed.stderr, str)
        assert completed.returncode == 4
        assert completed.stdout == ""
        assert missing_root in completed.stderr
        assert "required dependency" in completed.stderr
        assert "Traceback" not in completed.stderr

    def test_gui_entry_does_not_mask_unrelated_import_failure(self, monkeypatch: pytest.MonkeyPatch) -> None:
        import docwen_bundle.gui_entry as gui_entry

        def _raise_internal_missing(_argv=None) -> int:
            raise ModuleNotFoundError("No module named 'docwen_internal_bug'", name="docwen_internal_bug")

        monkeypatch.setattr(gui_entry, "_main_with_guard_active", _raise_internal_missing)

        with pytest.raises(ModuleNotFoundError, match="docwen_internal_bug"):
            gui_entry.main(["docwen-gui"])

    def test_gui_entry_wires_controller_window_and_cleanup(self, monkeypatch: pytest.MonkeyPatch) -> None:
        import docwen_application.controller as controller_module
        import docwen_bundle.config_port as config_port_module
        import docwen_bundle.gui_bootstrap as gui_bootstrap_module
        import docwen_bundle.gui_entry as gui_entry
        import docwen_bundle.runtime_factory as runtime_factory_module
        import docwen_gui.app as gui_app_module
        import docwen_gui.i18n as gui_i18n_module
        import docwen_gui.qt_bridge.task_event_bridge as bridge_module

        cleanup_events: list[str] = []
        instance_lock = SimpleNamespace(released=False)

        def _release_instance_lock() -> None:
            cleanup_events.append("lock_release")
            instance_lock.released = True

        instance_lock.release = _release_instance_lock

        monkeypatch.setattr(
            gui_bootstrap_module,
            "bootstrap_gui",
            lambda app_name, argv: SimpleNamespace(
                should_exit=False,
                exit_code=0,
                files_to_add=["startup.docx"],
                instance_lock=instance_lock,
                ipc_dir=None,
            ),
        )

        class _App:
            def __init__(self) -> None:
                self.aboutToQuit = _Signal()
                self.exec_called = False

            def exec(self) -> int:
                self.exec_called = True
                cleanup_events.append("app_exec")
                self.aboutToQuit.emit()
                cleanup_events.append("app_exec_return")
                return 0

        app = _App()
        monkeypatch.setattr(gui_app_module, "create_qapplication", lambda argv=None: app)
        initialized_theme: list[tuple[object, object]] = []
        configured_locales: list[str] = []
        monkeypatch.setattr(
            gui_app_module,
            "_initialize_application_theme",
            lambda qt_app, controller: initialized_theme.append((qt_app, controller)),
        )
        monkeypatch.setattr(gui_i18n_module, "set_locale", configured_locales.append)

        bridge = SimpleNamespace(events=[])
        bridge.enqueue = lambda event_type, payload: bridge.events.append((event_type, payload))
        monkeypatch.setattr(bridge_module, "TaskEventBridge", lambda: bridge)

        # Sentinel for later assignment — type stays narrow once set
        _sentinel: object = object()
        runtime_port = SimpleNamespace()
        _captured_event_callback: object = _sentinel
        _captured_controller: object = _sentinel
        _captured_window_controller: object = _sentinel
        _captured_window_bridge: object = _sentinel
        _captured_initial_files: object = _sentinel

        def _fake_create_runtime_port(*, config_loader=None, event_callback=None):
            nonlocal _captured_event_callback
            assert config_loader is not None
            _captured_event_callback = event_callback
            return runtime_port

        monkeypatch.setattr(runtime_factory_module, "create_runtime_port", _fake_create_runtime_port)

        class _ConfigPort:
            def __init__(self, loader=None) -> None:
                self.loader = loader

            def get(self, key: str, default=None):
                return "en_US" if key == "gui.language.locale" else default

        monkeypatch.setattr(config_port_module, "ConfigPortAdapter", _ConfigPort)

        class _Controller:
            def __init__(self, runtime_port=None, config_port=None) -> None:
                self.runtime_port = runtime_port
                self.config_port = config_port
                self.started = False
                self.stopped = False
                nonlocal _captured_controller
                _captured_controller = self

            def start(self) -> None:
                self.started = True

            def stop(self) -> None:
                self.stopped = True

        monkeypatch.setattr(controller_module, "ApplicationController", _Controller)

        window = SimpleNamespace(shown=False, close=lambda: None)
        window.show = lambda: setattr(window, "shown", True)

        def _fake_create_main_window(*, controller=None, task_event_bridge=None, initial_files=None):
            nonlocal _captured_window_controller, _captured_window_bridge, _captured_initial_files
            _captured_window_controller = controller
            _captured_window_bridge = task_event_bridge
            _captured_initial_files = initial_files
            return window

        monkeypatch.setattr(gui_app_module, "create_main_window", _fake_create_main_window)
        control_server = SimpleNamespace(stopped=False)

        def _begin_stop() -> None:
            cleanup_events.append("server_begin_stop")

        def _stop() -> None:
            cleanup_events.append("server_stop")
            control_server.stopped = True

        control_server.begin_stop = _begin_stop
        control_server.stop = _stop

        def _fake_start_gui_control(_window, *, app):
            app.aboutToQuit.connect(control_server.begin_stop)
            return control_server

        monkeypatch.setattr(gui_entry, "_start_gui_control", _fake_start_gui_control)

        exit_code = gui_entry.main(["docwen-gui", "startup.docx"])

        assert exit_code == 0
        assert app.exec_called is True
        assert window.shown is True
        assert _captured_initial_files == ["startup.docx"]
        assert _captured_window_bridge is bridge
        assert _captured_window_controller is _captured_controller
        assert initialized_theme == [(app, _captured_controller)]
        assert configured_locales == ["en_US"]
        ctrl = cast(_Controller, _captured_controller)
        assert ctrl.runtime_port is runtime_port
        assert isinstance(ctrl.config_port, _ConfigPort)
        assert ctrl.started is True

        cast(Callable[[TaskEvent], object], _captured_event_callback)(
            TaskEvent(
                event_type="task_started",
                task_id="task-1",
                sequence=1,
                payload={"stage": "queued"},
            )
        )
        assert bridge.events == [("task_started", {"task_id": "task-1", "stage": "queued"})]

        assert ctrl.stopped is False
        assert control_server.stopped is True
        assert instance_lock.released is True
        assert cleanup_events == [
            "app_exec",
            "server_begin_stop",
            "app_exec_return",
            "server_stop",
            "lock_release",
        ]

    def test_gui_control_stop_failure_still_releases_instance_lock(self) -> None:
        import docwen_bundle.gui_entry as gui_entry

        events: list[str] = []

        class _Server:
            def stop(self) -> None:
                events.append("server_stop")
                raise RuntimeError("cleanup failed")

        instance_lock = SimpleNamespace(release=lambda: events.append("lock_release"))

        with pytest.raises(RuntimeError, match="cleanup failed"):
            gui_entry._stop_gui_control_and_release_instance_lock(_Server(), instance_lock)

        assert events == ["server_stop", "lock_release"]

    def test_start_gui_control_marshals_requests_to_gui_thread(self, monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
        import threading
        import time

        import docwen_bundle.gui_entry as gui_entry
        import docwen_runtime.control as control_module

        calls: list[str] = []
        servers: list[object] = []

        class _FakeControlServer:
            def __init__(self, handler, *, app_name) -> None:
                self.handler = handler
                self.app_name = app_name
                servers.append(self)

            def start(self) -> None:
                calls.append("start")

            def stop(self) -> None:
                calls.append("stop")

            def begin_stop(self) -> None:
                calls.append("stop")

        monkeypatch.setattr(control_module, "ControlServer", _FakeControlServer)

        installed: list[Callable[[], None]] = []

        def _fake_install_timer(app, drain_pending_commands) -> object:
            assert app is fake_app
            installed.append(drain_pending_commands)
            return SimpleNamespace(stop=lambda: calls.append("timer_stop"))

        monkeypatch.setattr(gui_entry, "_install_gui_control_poll_timer", _fake_install_timer)

        fake_app = SimpleNamespace(aboutToQuit=_Signal())
        window = SimpleNamespace(handled=[], opened_settings=[], isVisible=lambda: True)
        window.handle_ipc_command = lambda action, path=None: window.handled.append((action, path))
        window.supported_settings_sections = lambda: ("proofread",)
        window.open_settings = lambda section, *, deadline=None: (
            window.opened_settings.append(section) or {"accepted": True, "section": section, "reused": False}
        )
        sample = tmp_path / "sample.md"
        sample.write_text("hello", encoding="utf-8")

        server = gui_entry._start_gui_control(window, app=fake_app)

        assert isinstance(server, _FakeControlServer)
        assert len(installed) == 1
        assert calls == ["start"]

        result: dict[str, object] = {}
        worker = threading.Thread(target=lambda: result.update(server.handler("open", {"file": str(sample.resolve())})))
        worker.start()
        time.sleep(0.02)
        installed[0]()
        worker.join(1)

        assert result["accepted"] is True
        assert window.handled == [("open_file", str(sample.resolve()))]

        status: dict[str, object] = {}
        worker = threading.Thread(target=lambda: status.update(server.handler("status", {})))
        worker.start()
        time.sleep(0.02)
        installed[0]()
        worker.join(1)
        assert status["supported_actions"] == ["status", "activate", "open", "open_settings"]
        assert status["settings_sections"] == ["proofread"]

        settings_result: dict[str, object] = {}
        worker = threading.Thread(
            target=lambda: settings_result.update(server.handler("open_settings", {"section": "proofread"}))
        )
        worker.start()
        time.sleep(0.02)
        installed[0]()
        worker.join(1)
        assert settings_result == {
            "accepted": True,
            "running": True,
            "action": "open_settings",
            "section": "proofread",
            "reused": False,
        }
        assert window.opened_settings == ["proofread"]

        settings_errors: list[Any] = []

        def _open_unknown_settings() -> None:
            try:
                server.handler("open_settings", {"section": "document"})
            except control_module.ControlRequestError as exc:
                settings_errors.append(exc)

        worker = threading.Thread(target=_open_unknown_settings)
        worker.start()
        time.sleep(0.02)
        installed[0]()
        worker.join(1)
        assert len(settings_errors) == 1
        assert settings_errors[0].code == "settings_section_unavailable"

        expired_errors: list[Any] = []

        def _open_expired_settings() -> None:
            try:
                server.handler(
                    "open_settings",
                    {
                        "section": "proofread",
                        "_deadline_monotonic": time.monotonic() + 0.03,
                    },
                )
            except control_module.ControlRequestError as exc:
                expired_errors.append(exc)

        worker = threading.Thread(target=_open_expired_settings)
        worker.start()
        time.sleep(0.08)
        worker.join(1)
        installed[0]()
        assert len(expired_errors) == 1
        assert expired_errors[0].code == "control_timeout"
        assert window.opened_settings == ["proofread"]

        expired_action_errors: list[Any] = []

        def _activate_after_deadline() -> None:
            try:
                server.handler(
                    "activate",
                    {"_deadline_monotonic": time.monotonic() + 0.03},
                )
            except control_module.ControlRequestError as exc:
                expired_action_errors.append(exc)

        worker = threading.Thread(target=_activate_after_deadline)
        worker.start()
        time.sleep(0.08)
        worker.join(1)
        installed[0]()
        assert len(expired_action_errors) == 1
        assert expired_action_errors[0].code == "control_timeout"
        assert window.handled == [("open_file", str(sample.resolve()))]

        shutdown_errors: list[Any] = []

        def _activate_while_gui_quits() -> None:
            try:
                server.handler(
                    "activate",
                    {"_deadline_monotonic": time.monotonic() + 30.0},
                )
            except control_module.ControlRequestError as exc:
                shutdown_errors.append(exc)

        worker = threading.Thread(target=_activate_while_gui_quits)
        worker.start()
        time.sleep(0.02)
        quit_started = time.monotonic()
        fake_app.aboutToQuit.emit()
        worker.join(0.5)
        assert not worker.is_alive()
        assert time.monotonic() - quit_started < 0.5
        assert len(shutdown_errors) == 1
        assert shutdown_errors[0].code == "gui_stopping"
        assert window.handled == [("open_file", str(sample.resolve()))]
        assert calls == ["start", "timer_stop", "stop"]

        with pytest.raises(control_module.ControlRequestError) as exc_info:
            server.handler("activate", {"_deadline_monotonic": time.monotonic() + 30.0})
        assert exc_info.value.code == "gui_stopping"

    def test_explicit_control_deadline_is_not_capped_at_fifteen_seconds(
        self,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        import docwen_bundle.gui_entry as gui_entry

        monkeypatch.setattr(gui_entry.time, "monotonic", lambda: 100.0)

        assert gui_entry._control_wait_timeout(None) == 15.0
        assert gui_entry._control_wait_timeout(130.0) == 30.0
        assert gui_entry._control_wait_timeout(99.0) == 0.0
