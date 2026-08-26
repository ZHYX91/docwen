"""Focused tests split from test_main_window_window_behavior.py."""

from __future__ import annotations

from ._main_window_window_behavior_support import (
    Any,
    QWidget,
    _make_window,
    _root_grid,
    pytest,
)

pytestmark = pytest.mark.gui


def test_settings_dialog_source_signal_refreshes_main_window_policy(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_core.models.file_ref import FileRef
    from docwen_gui.widgets.settings import dialog as dialog_module

    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.expand_side_panels": False,
        },
    )
    centered: list[bool] = []

    class _Signal:
        def __init__(self) -> None:
            self._callbacks: list[Any] = []

        def connect(self, callback: Any) -> None:
            self._callbacks.append(callback)

        def emit(self) -> None:
            for callback in self._callbacks:
                callback()

    class _Dialog:
        def __init__(self, *, parent: QWidget, view_model: object) -> None:
            del parent, view_model
            self.settings_source_changed = _Signal()
            self.destroyed = _Signal()
            self.finished = _Signal()

        def activate_section(self, _section: str) -> bool:
            return True

        def open(self) -> None:
            controller.config_port.values.update(
                {
                    "gui.window.remember_gui_state": False,
                    "gui.window.auto_center": True,
                    "gui.window.expand_side_panels": True,
                }
            )
            self.settings_source_changed.emit()

        def raise_(self) -> None:
            return None

        def activateWindow(self) -> None:
            return None

    try:
        window._view_model.set_selected_file(  # pyright: ignore[reportPrivateUsage]
            FileRef(
                path="C:/tmp/test.docx",
                format="docx",
                category="document",
                size_bytes=0,
            )
        )
        assert _root_grid(window).columnStretch(2) == 0
        monkeypatch.setattr(window, "_center_on_screen", lambda: centered.append(True))
        monkeypatch.setattr(dialog_module, "SettingsDialog", _Dialog)

        window._open_settings_dialog()  # pyright: ignore[reportPrivateUsage]

        assert window._window_behavior.remember_gui_state is False  # pyright: ignore[reportPrivateUsage]
        assert window._window_behavior.auto_center is True  # pyright: ignore[reportPrivateUsage]
        assert window._window_behavior.expand_side_panels is True  # pyright: ignore[reportPrivateUsage]
        assert _root_grid(window).columnStretch(2) == 3
        assert centered == []
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_main_window_reuses_owned_nonblocking_settings_dialog(main_window, monkeypatch) -> None:
    from docwen_gui.widgets.settings import dialog as dialog_module

    instances: list[Any] = []
    monkeypatch.setattr(main_window, "bring_to_front", lambda: None)

    class _Signal:
        def __init__(self) -> None:
            self._callbacks: list[Any] = []

        def connect(self, callback: Any) -> None:
            self._callbacks.append(callback)

        def emit(self) -> None:
            for callback in self._callbacks:
                callback()

    class _Dialog:
        def __init__(self, *, parent: QWidget, view_model: object) -> None:
            del parent, view_model
            self.settings_source_changed = _Signal()
            self.destroyed = _Signal()
            self.finished = _Signal()
            self.sections: list[str] = []
            self.open_count = 0
            self.focus_count = 0
            instances.append(self)

        def activate_section(self, section: str) -> bool:
            self.sections.append(section)
            return section == "proofread"

        def current_section(self) -> str:
            return self.sections[-1] if self.sections else "general"

        def open(self) -> None:
            self.open_count += 1

        def raise_(self) -> None:
            self.focus_count += 1

        def activateWindow(self) -> None:
            self.focus_count += 1

        def deleteLater(self) -> None:
            return None

    monkeypatch.setattr(dialog_module, "SettingsDialog", _Dialog)

    first = main_window.open_settings("proofread")
    second = main_window.open_settings("proofread")

    assert first == {"accepted": True, "section": "proofread", "reused": False}
    assert second == {"accepted": True, "section": "proofread", "reused": True}
    assert len(instances) == 1
    assert instances[0].sections == ["proofread", "proofread"]
    assert instances[0].open_count == 2
    assert instances[0].focus_count == 4
    instances[0].destroyed.emit()
    assert main_window._settings_dialog is None  # pyright: ignore[reportPrivateUsage]


def test_main_window_discards_settings_dialog_when_construction_crosses_deadline(
    main_window,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_gui.main_window as main_window_module
    from docwen_gui.widgets.settings import dialog as dialog_module

    now = [100.0]
    instances: list[Any] = []
    monkeypatch.setattr(main_window_module.time, "monotonic", lambda: now[0])
    monkeypatch.setattr(main_window, "bring_to_front", lambda: pytest.fail("must not focus after timeout"))

    class _Signal:
        def connect(self, _callback: Any) -> None:
            return None

    class _Dialog:
        def __init__(self, *, parent: QWidget, view_model: object) -> None:
            del parent, view_model
            self.settings_source_changed = _Signal()
            self.destroyed = _Signal()
            self.finished = _Signal()
            self.open_count = 0
            self.focus_count = 0
            self.deleted = False
            instances.append(self)
            now[0] = 102.0

        def activate_section(self, _section: str) -> bool:
            raise AssertionError("expired dialog must not activate a page")

        def open(self) -> None:
            self.open_count += 1

        def raise_(self) -> None:
            self.focus_count += 1

        def activateWindow(self) -> None:
            self.focus_count += 1

        def deleteLater(self) -> None:
            self.deleted = True

    monkeypatch.setattr(dialog_module, "SettingsDialog", _Dialog)

    result = main_window.open_settings("proofread", deadline=101.0)

    assert result == {
        "accepted": False,
        "section": "proofread",
        "reused": False,
        "error_code": "control_timeout",
    }
    assert len(instances) == 1
    assert instances[0].deleted is True
    assert instances[0].open_count == 0
    assert instances[0].focus_count == 0
    assert main_window._settings_dialog is None  # pyright: ignore[reportPrivateUsage]


def test_main_window_real_settings_dialog_is_nonblocking_owned_singleton(
    main_window,
    qapp,
) -> None:
    from docwen_gui.view_models.settings_vm import SECTION_PROOFREAD

    first = main_window.open_settings("proofread")
    first_dialog = main_window._settings_dialog  # pyright: ignore[reportPrivateUsage]
    qapp.processEvents()

    assert first == {"accepted": True, "section": "proofread", "reused": False}
    assert first_dialog is not None
    assert first_dialog.isModal() is True
    assert first_dialog.isVisible() is True
    assert first_dialog.current_section() == "proofread"

    first_dialog.view_model.set_field(SECTION_PROOFREAD, "sensitive_word", False)
    second = main_window.open_settings("proofread")
    qapp.processEvents()

    assert second == {"accepted": True, "section": "proofread", "reused": True}
    assert main_window._settings_dialog is first_dialog  # pyright: ignore[reportPrivateUsage]
    assert first_dialog.view_model.config.proofread.sensitive_word is False

    first_dialog.view_model.cancel_changes()
    first_dialog.reject()
    qapp.processEvents()
    assert main_window._settings_dialog is None  # pyright: ignore[reportPrivateUsage]

    reopened = main_window.open_settings("proofread")
    reopened_dialog = main_window._settings_dialog  # pyright: ignore[reportPrivateUsage]
    qapp.processEvents()

    assert reopened == {"accepted": True, "section": "proofread", "reused": False}
    assert reopened_dialog is not None and reopened_dialog is not first_dialog
    reopened_dialog.reject()
    qapp.processEvents()


def test_main_window_failed_public_settings_page_is_unavailable_without_owner(
    main_window,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.widgets.settings import dialog as dialog_module

    original = dialog_module._TAB_SPECS["proofread"]  # pyright: ignore[reportPrivateUsage]

    def _fail_proofread(_view_model: object) -> QWidget:
        raise RuntimeError("proofread page failed")

    monkeypatch.setitem(
        dialog_module._TAB_SPECS,  # pyright: ignore[reportPrivateUsage]
        "proofread",
        dialog_module._TabSpec(original.module_name, original.class_name, _fail_proofread),  # pyright: ignore[reportPrivateUsage]
    )

    result = main_window.open_settings("proofread")

    assert result == {"accepted": False, "section": "proofread", "reused": False}
    assert main_window._settings_dialog is None  # pyright: ignore[reportPrivateUsage]
