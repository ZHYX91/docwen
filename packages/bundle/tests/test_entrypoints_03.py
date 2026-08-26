"""Focused tests split from test_entrypoints.py."""

from __future__ import annotations

from ._entrypoints_support import (
    SimpleNamespace,
    _Signal,
    pytest,
)

pytestmark = pytest.mark.unit


class TestGuiAutocloseSingleHome:
    """C1 closure: the test-autoclose helper must live in exactly one place.

    The plan removed the duplicate ``run_gui_with_ipc`` entry and required the
    duplicate ``_schedule_test_autoclose`` to be collapsed to a single copy.
    Dependency honesty dictates the single home: ``docwen_gui.app`` (the Qt-owning
    layer). ``docwen_bundle`` already imports ``docwen_gui.app``, and the
    import-linter contract forbids the reverse (gui → bundle). So the bundle
    entry must call the gui-side helper, not keep its own private copy.
    """

    def test_schedule_test_autoclose_defined_only_in_gui_app(self) -> None:
        """The helper must not be (re)defined inside docwen_bundle."""
        import ast
        from pathlib import Path

        repo_root = Path(__file__).resolve().parents[3]
        bundle_src = repo_root / "packages" / "bundle" / "src"

        offenders: list[str] = []
        for py_file in bundle_src.rglob("*.py"):
            source = py_file.read_text(encoding="utf-8")
            tree = ast.parse(source, filename=str(py_file))
            for node in ast.walk(tree):
                if isinstance(node, ast.FunctionDef) and node.name == "_schedule_test_autoclose":
                    offenders.append(str(py_file.relative_to(repo_root)))

        assert offenders == [], (
            "_schedule_test_autoclose must not be defined in docwen_bundle "
            "(single home is docwen_gui.app); found def in: " + ", ".join(offenders)
        )

    def test_gui_app_exposes_schedule_test_autoclose(self) -> None:
        """The single home (docwen_gui.app) must still define the helper."""
        from docwen_gui import app as gui_app

        assert hasattr(gui_app, "_schedule_test_autoclose")
        assert callable(gui_app._schedule_test_autoclose)

    def test_gui_shutdown_entrypoints_route_through_main_window_close(self) -> None:
        from pathlib import Path

        repo_root = Path(__file__).resolve().parents[3]
        gui_app = (repo_root / "packages/apps/gui/src/docwen_gui/app.py").read_text(encoding="utf-8")
        release_smoke = (repo_root / "packages/apps/gui/src/docwen_gui/release_smoke.py").read_text(encoding="utf-8")
        bundle_entry = (repo_root / "packages/bundle/src/docwen_bundle/gui_entry.py").read_text(encoding="utf-8")

        assert "app.quit" not in gui_app
        assert "app.quit" not in release_smoke
        assert "controller.stop()" not in bundle_entry
        assert "_schedule_test_autoclose(app, window)" in bundle_entry

    def test_bundle_gui_entry_uses_gui_side_autoclose(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """bundle.gui_entry.main must schedule autoclose via docwen_gui.app's helper, not a local one."""
        import docwen_bundle.gui_bootstrap as gui_bootstrap_module
        import docwen_bundle.gui_entry as gui_entry
        import docwen_gui.app as gui_app_module

        # bundle must NOT carry its own copy anymore.
        assert not hasattr(gui_entry, "_schedule_test_autoclose"), (
            "docwen_bundle.gui_entry still defines its own _schedule_test_autoclose; "
            "it should call docwen_gui.app._schedule_test_autoclose instead."
        )

        called: list[bool] = []
        monkeypatch.setattr(
            gui_app_module,
            "_schedule_test_autoclose",
            lambda app, window: called.append(True),
        )
        monkeypatch.setattr(
            gui_bootstrap_module,
            "bootstrap_gui",
            lambda app_name, argv: SimpleNamespace(
                should_exit=False,
                exit_code=0,
                files_to_add=[],
                instance_lock=None,
                ipc_dir=None,
            ),
        )

        class _App:
            def __init__(self) -> None:
                self.aboutToQuit = _Signal()

            def exec(self) -> int:
                return 0

        monkeypatch.setattr(gui_app_module, "create_qapplication", lambda argv=None: _App())
        monkeypatch.setattr(
            gui_app_module,
            "create_main_window",
            lambda **kw: SimpleNamespace(show=lambda: None, close=lambda: None),
        )
        monkeypatch.setattr(gui_app_module, "_initialize_application_theme", lambda app, controller: None)

        # Stub the bundle-internal wiring collaborators so main() reaches the autoclose call.
        import docwen_application.controller as controller_module
        import docwen_bundle.config_port as config_port_module
        import docwen_bundle.runtime_factory as runtime_factory_module
        import docwen_gui.qt_bridge.task_event_bridge as bridge_module

        monkeypatch.setattr(bridge_module, "TaskEventBridge", lambda: SimpleNamespace(enqueue=lambda *a, **kw: None))
        monkeypatch.setattr(runtime_factory_module, "create_runtime_port", lambda **kw: SimpleNamespace())
        monkeypatch.setattr(
            config_port_module,
            "ConfigPortAdapter",
            lambda _loader: SimpleNamespace(get=lambda _key, default=None: default),
        )

        class _Controller:
            def start(self) -> None: ...
            def stop(self) -> None: ...

        monkeypatch.setattr(controller_module, "ApplicationController", lambda **kw: _Controller())

        gui_entry.main(["docwen-gui"])

        assert called == [True], (
            "bundle.gui_entry.main did not delegate autoclose scheduling to docwen_gui.app._schedule_test_autoclose."
        )


class TestGuiMainDelegation:
    def test_docwen_gui_main_delegates_to_bundle_gui_entry(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """python -m docwen_gui must delegate to bundle.gui_entry.main, not run_gui_with_ipc."""
        import docwen_bundle.gui_entry as gui_entry
        import docwen_gui.__main__ as gui_main

        _called: bool = False
        _argv: object = None

        def _fake_bundle_main(argv=None):
            nonlocal _called, _argv
            _called = True
            _argv = argv
            return 0

        monkeypatch.setattr(gui_entry, "main", _fake_bundle_main)

        exit_code = gui_main.main()

        assert exit_code == 0
        assert _called is True

    def test_pyi_gui_entry_delegates_and_preserves_exit_code(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """PyInstaller entry must delegate to bundle.gui_entry.main and preserve its exit code."""
        import docwen_bundle.gui_entry as gui_entry
        import docwen_bundle.pyi_gui_entry as pyi

        monkeypatch.setattr(gui_entry, "main", lambda argv=None: 7)

        # pyi_gui_entry runs under __name__ == "__main__"; we exercise the delegate function directly
        exit_code = pyi._delegate()

        assert exit_code == 7

    def test_pyi_gui_entry_prepares_frozen_multiprocessing(self, monkeypatch: pytest.MonkeyPatch) -> None:
        import docwen_bundle.pyi_gui_entry as pyi

        calls: list[bool] = []
        monkeypatch.setattr(pyi.multiprocessing, "freeze_support", lambda: calls.append(True))

        pyi._prepare_multiprocessing()

        assert calls == [True]
