"""Focused tests split from test_main_window_window_behavior.py."""

from __future__ import annotations

from ._main_window_window_behavior_support import (
    _HIDDEN,
    _VISIBLE_BOTH,
    _VISIBLE_LEFT,
    _VISIBLE_RIGHT,
    MainWindowUiProjection,
    QPoint,
    QTest,
    QWidget,
    _make_window,
    _root_grid,
    pytest,
)

pytestmark = pytest.mark.gui


def test_state_save_excludes_visible_context_panel_width(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.default_width": 460,
            "gui.window.default_height": 760,
            "gui.window.right_panel_width": 300,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        base_width = window.width()
        window._view_model.ui_projection_changed.emit(_VISIBLE_RIGHT)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        fit = window._screen_fit_rect(window.screen())  # pyright: ignore[reportPrivateUsage]
        if fit is None or fit.width < base_width + window._RIGHT_PANEL_MIN_WIDTH + 8:  # pyright: ignore[reportPrivateUsage]
            pytest.skip("test screen cannot fit the configured right panel")

        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        saved = controller.config_port.set_many_calls[-1]
        assert saved["gui.window.default_width"] == base_width
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_maximized_save_uses_paired_normal_frame_rect_and_anchor(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.expand_side_panels": True,
            "gui.window.center_panel_screen_x": 250,
            "gui.window.window_y": 20,
            "gui.window.default_width": 700,
            "gui.window.default_height": 720,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        window._view_model.ui_projection_changed.emit(_VISIBLE_LEFT)  # pyright: ignore[reportPrivateUsage]
        window.move(50, 30)
        window.resize(700, 720)
        qapp.processEvents()
        normal_rect = window._normal_window_rect()  # pyright: ignore[reportPrivateUsage]
        normal_offset = window._center_column_offset()  # pyright: ignore[reportPrivateUsage]

        window.showMaximized()
        qapp.processEvents()
        assert window.isMaximized()
        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        expected_base_width = max(
            window.minimumWidth(),
            normal_rect.width
            - window._context_panel_width_contribution(  # pyright: ignore[reportPrivateUsage]
                left_visible=True,
                right_visible=False,
            ),
        )
        assert controller.config_port.set_many_calls[-1] == {
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": normal_rect.x + normal_offset,
            "gui.window.window_y": normal_rect.y,
            "gui.window.default_width": expected_base_width,
            "gui.window.default_height": normal_rect.height,
        }
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_panel_change_while_maximized_preserves_normal_anchor_for_save_and_restore(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.expand_side_panels": True,
            "gui.window.center_panel_screen_x": 300,
            "gui.window.window_y": 20,
            "gui.window.default_width": 476,
            "gui.window.default_height": 720,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        preserved_anchor = center_column.mapToGlobal(QPoint(0, 0)).x()

        window.showMaximized()
        qapp.processEvents()
        window._view_model.ui_projection_changed.emit(_VISIBLE_LEFT)  # pyright: ignore[reportPrivateUsage]
        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        assert controller.config_port.set_many_calls[-1]["gui.window.center_panel_screen_x"] == preserved_anchor

        window.showNormal()
        qapp.processEvents()
        qapp.processEvents()  # drain the queued native-normal-frame finalizer
        assert center_column.mapToGlobal(QPoint(0, 0)).x() == preserved_anchor
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_move_during_queued_normal_restore_cancels_the_pending_panel_anchor(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.expand_side_panels": True,
            "gui.window.center_panel_screen_x": 300,
            "gui.window.window_y": 20,
            "gui.window.default_width": 460,
            "gui.window.default_height": 720,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None

        window.showMaximized()
        qapp.processEvents()
        window._view_model.ui_projection_changed.emit(_VISIBLE_LEFT)  # pyright: ignore[reportPrivateUsage]
        window.showNormal()
        qapp.processEvents()

        window.move(100, 20)
        qapp.processEvents()
        qapp.processEvents()
        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        actual_anchor = center_column.mapToGlobal(QPoint(0, 0)).x()
        assert window.pos() == QPoint(100, 20)
        assert window._pending_normal_center_anchor is None  # pyright: ignore[reportPrivateUsage]
        assert controller.config_port.set_many_calls[-1]["gui.window.center_panel_screen_x"] == actual_anchor
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_save_during_queued_normal_restore_uses_the_paired_normal_rect(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.expand_side_panels": True,
            "gui.window.center_panel_screen_x": 300,
            "gui.window.window_y": 20,
            "gui.window.default_width": 460,
            "gui.window.default_height": 720,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        normal_rect = window._last_normal_window_rect  # pyright: ignore[reportPrivateUsage]
        assert normal_rect is not None

        window.showMaximized()
        qapp.processEvents()
        window._view_model.ui_projection_changed.emit(_VISIBLE_LEFT)  # pyright: ignore[reportPrivateUsage]
        window.showNormal()
        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        saved = controller.config_port.set_many_calls[-1]
        assert saved["gui.window.center_panel_screen_x"] == 300
        assert saved["gui.window.window_y"] == normal_rect.y
        assert saved["gui.window.default_width"] == normal_rect.width
        assert saved["gui.window.default_height"] == normal_rect.height
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_one_pixel_native_normal_rect_difference_does_not_strand_panel_anchor(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from docwen_gui.window_geometry import WindowRect

    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.expand_side_panels": True,
            "gui.window.center_panel_screen_x": 300,
            "gui.window.window_y": 20,
            "gui.window.default_width": 460,
            "gui.window.default_height": 720,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        normal_rect = window._last_normal_window_rect  # pyright: ignore[reportPrivateUsage]
        assert normal_rect is not None

        window.showMaximized()
        qapp.processEvents()
        window._view_model.ui_projection_changed.emit(_VISIBLE_LEFT)  # pyright: ignore[reportPrivateUsage]
        window._last_normal_window_rect = WindowRect(  # pyright: ignore[reportPrivateUsage]
            x=normal_rect.x + 1,
            y=normal_rect.y,
            width=normal_rect.width,
            height=normal_rect.height,
        )

        window.showNormal()
        qapp.processEvents()
        qapp.processEvents()

        assert center_column.mapToGlobal(QPoint(0, 0)).x() == 300
        assert window._pending_normal_center_anchor is None  # pyright: ignore[reportPrivateUsage]
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_reverse_order_normal_restore_uses_bounded_quiet_fallback(qapp) -> None:
    from docwen_gui.window_geometry import WindowRect

    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.center_panel_screen_x": 300,
            "gui.window.window_y": 20,
            "gui.window.default_width": 460,
            "gui.window.default_height": 720,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        current = WindowRect(
            x=window.pos().x(),
            y=window.pos().y(),
            width=window.width(),
            height=window.height(),
        )
        target_anchor = center_column.mapToGlobal(QPoint(0, 0)).x() + 20
        window._last_normal_window_rect = WindowRect(  # pyright: ignore[reportPrivateUsage]
            x=current.x + 1,
            y=current.y,
            width=current.width,
            height=current.height,
        )
        window._pending_normal_center_anchor = target_anchor  # pyright: ignore[reportPrivateUsage]
        window._pending_normal_frame_restored = False  # pyright: ignore[reportPrivateUsage]
        window._pending_anchor_restore_scheduled = True  # pyright: ignore[reportPrivateUsage]

        window._apply_pending_normal_center_anchor()  # pyright: ignore[reportPrivateUsage]
        QTest.qWait(100)

        assert center_column.mapToGlobal(QPoint(0, 0)).x() == target_anchor
        assert window._pending_normal_center_anchor is None  # pyright: ignore[reportPrivateUsage]
        assert window._pending_anchor_restore_scheduled is False  # pyright: ignore[reportPrivateUsage]
    finally:
        window.close()


def test_unknown_future_geometry_schema_is_not_downgraded_on_close(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 999,
            "gui.window.center_panel_screen_x": 700,
            "gui.window.window_y": 400,
            "gui.window.default_width": 700,
            "gui.window.default_height": 760,
        },
    )
    try:
        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        assert controller.config_port.set_many_calls == []
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_state_save_reloads_remember_policy_before_writing(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
        },
    )
    try:
        controller.config_port.values["gui.window.remember_gui_state"] = False

        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        assert controller.config_port.set_calls == []
        assert controller.config_port.set_many_calls == []
        assert window._window_behavior.remember_gui_state is False  # pyright: ignore[reportPrivateUsage]
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_geometry_source_reset_is_not_undone_on_close_without_a_later_move(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 650,
            "gui.window.window_y": 20,
            "gui.window.default_width": 560,
            "gui.window.default_height": 760,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        controller.config_port.values.update(
            {
                "gui.window.center_panel_screen_x": 420,
                "gui.window.window_y": 0,
                "gui.window.default_width": 476,
                "gui.window.default_height": 860,
            }
        )
        window._apply_runtime_window_settings()  # pyright: ignore[reportPrivateUsage]

        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        assert controller.config_port.set_many_calls == []

        window._view_model.ui_projection_changed.emit(_VISIBLE_LEFT)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        assert controller.config_port.set_many_calls == []

        window.move(100, 10)
        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        assert len(controller.config_port.set_many_calls) == 1
        assert controller.config_port.set_many_calls[0]["gui.window.window_y"] == 10
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


@pytest.mark.parametrize(
    ("expand_side_panels", "projection", "expected_left", "expected_right"),
    [
        (False, _HIDDEN, 0, 0),
        (False, _VISIBLE_LEFT, 0, 0),
        (False, _VISIBLE_RIGHT, 0, 0),
        (False, _VISIBLE_BOTH, 0, 0),
        (True, _HIDDEN, 0, 0),
        (True, _VISIBLE_LEFT, 2, 0),
        (True, _VISIBLE_RIGHT, 0, 3),
        (True, _VISIBLE_BOTH, 2, 3),
    ],
)
def test_expand_policy_never_allocates_hidden_columns(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
    expand_side_panels: bool,
    projection: MainWindowUiProjection,
    expected_left: int,
    expected_right: int,
) -> None:
    window, _controller = _make_window(
        qapp,
        {"gui.window.expand_side_panels": expand_side_panels},
    )
    try:
        window._view_model.ui_projection_changed.emit(projection)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        layout = _root_grid(window)
        assert layout.columnStretch(0) == expected_left
        assert layout.columnStretch(1) == 4
        assert layout.columnStretch(2) == expected_right
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_runtime_policy_refreshes_flags_and_layout_without_recentering(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.expand_side_panels": False,
        },
    )
    centered: list[bool] = []
    try:
        from docwen_core.models.file_ref import FileRef

        window._view_model.set_selected_file(  # pyright: ignore[reportPrivateUsage]
            FileRef(
                path="C:/tmp/test.docx",
                format="docx",
                category="document",
                size_bytes=0,
            )
        )
        assert _root_grid(window).columnStretch(2) == 0
        controller.config_port.values.update(
            {
                "gui.window.remember_gui_state": False,
                "gui.window.auto_center": True,
                "gui.window.expand_side_panels": True,
            }
        )
        monkeypatch.setattr(window, "_center_on_screen", lambda: centered.append(True))

        window._apply_runtime_window_settings()  # pyright: ignore[reportPrivateUsage]

        assert window._window_behavior.remember_gui_state is False  # pyright: ignore[reportPrivateUsage]
        assert window._window_behavior.auto_center is True  # pyright: ignore[reportPrivateUsage]
        assert window._window_behavior.expand_side_panels is True  # pyright: ignore[reportPrivateUsage]
        assert _root_grid(window).columnStretch(2) == 3
        assert centered == []
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()
