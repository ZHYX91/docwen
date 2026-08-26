"""Focused tests split from test_main_window_window_behavior.py."""

from __future__ import annotations

from ._main_window_window_behavior_support import (
    _HIDDEN,
    _VISIBLE_BOTH,
    _VISIBLE_LEFT,
    _VISIBLE_RIGHT,
    Any,
    MainWindowUiProjection,
    QPoint,
    QRect,
    QWidget,
    RightPanelSlot,
    WindowRect,
    _make_window,
    pytest,
)

pytestmark = pytest.mark.gui


@pytest.mark.parametrize(
    ("projection_sequence", "expected_left", "expected_right"),
    [
        ((_VISIBLE_RIGHT, _VISIBLE_BOTH), True, True),
        ((_VISIBLE_BOTH, _VISIBLE_LEFT), True, False),
        ((_VISIBLE_LEFT, _VISIBLE_RIGHT), False, True),
    ],
)
def test_panel_change_while_maximized_restores_pending_normal_frame_with_runtime_width_plan(
    qapp,
    projection_sequence: tuple[MainWindowUiProjection, ...],
    expected_left: bool,
    expected_right: bool,
) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 400,
            "gui.window.window_y": 20,
            "gui.window.default_width": 476,
            "gui.window.default_height": 760,
            "gui.window.expand_side_panels": False,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        normal_size = window.size()
        preserved_anchor = center_column.mapToGlobal(QPoint(0, 0)).x()

        window.showMaximized()
        qapp.processEvents()
        for projection in projection_sequence:
            window._view_model.ui_projection_changed.emit(projection)  # pyright: ignore[reportPrivateUsage]
            qapp.processEvents()

        maximized_plan = window._panel_width_plan(  # pyright: ignore[reportPrivateUsage]
            left_visible=expected_left,
            right_visible=expected_right,
        )
        assert window._left_panel_frame.minimumWidth() == maximized_plan.left  # pyright: ignore[reportPrivateUsage]
        assert center_column.minimumWidth() == maximized_plan.center
        assert window._right_panel_frame.minimumWidth() == maximized_plan.right  # pyright: ignore[reportPrivateUsage]
        assert window._pending_normal_center_anchor == preserved_anchor  # pyright: ignore[reportPrivateUsage]

        window.showNormal()
        qapp.processEvents()
        qapp.processEvents()

        restored_plan = window._panel_width_plan(  # pyright: ignore[reportPrivateUsage]
            left_visible=expected_left,
            right_visible=expected_right,
        )
        assert window.size() == normal_size
        assert center_column.mapToGlobal(QPoint(0, 0)).x() == preserved_anchor
        assert window._pending_normal_center_anchor is None  # pyright: ignore[reportPrivateUsage]
        assert window._left_panel_frame.minimumWidth() == restored_plan.left  # pyright: ignore[reportPrivateUsage]
        assert center_column.minimumWidth() == restored_plan.center
        assert window._right_panel_frame.minimumWidth() == restored_plan.right  # pyright: ignore[reportPrivateUsage]
        assert (restored_plan.left == 0) is (not expected_left)
        assert (restored_plan.right == 0) is (not expected_right)
    finally:
        window.close()


def test_simulated_half_screen_round_trip_restores_collapsed_frame_and_discards_hidden_widths(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 400,
            "gui.window.window_y": 20,
            "gui.window.default_width": 476,
            "gui.window.default_height": 760,
            "gui.window.expand_side_panels": False,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        fit = window._screen_fit_rect(window.screen())  # pyright: ignore[reportPrivateUsage]
        assert fit is not None
        snapped_width = max(window.minimumWidth(), fit.width // 2)
        snapped_height = min(window.height(), fit.height)
        snapped_x = fit.x + max(0, (fit.width - snapped_width) // 2)
        window.setGeometry(snapped_x, fit.y, snapped_width, snapped_height)
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        collapsed_rect = window.geometry()
        collapsed_anchor = center_column.mapToGlobal(QPoint(0, 0)).x()

        window._view_model.ui_projection_changed.emit(_VISIBLE_LEFT)  # pyright: ignore[reportPrivateUsage]
        window._view_model.ui_projection_changed.emit(_VISIBLE_BOTH)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        expanded_plan = window._panel_width_plan(  # pyright: ignore[reportPrivateUsage]
            left_visible=True,
            right_visible=True,
        )
        assert expanded_plan.left > 0
        assert expanded_plan.center > 0
        assert expanded_plan.right > 0
        assert window._left_panel_frame.minimumWidth() == expanded_plan.left  # pyright: ignore[reportPrivateUsage]
        assert center_column.minimumWidth() == expanded_plan.center
        assert window._right_panel_frame.minimumWidth() == expanded_plan.right  # pyright: ignore[reportPrivateUsage]

        window._view_model.ui_projection_changed.emit(_HIDDEN)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        assert window.geometry() == collapsed_rect
        assert center_column.mapToGlobal(QPoint(0, 0)).x() == collapsed_anchor
        assert window._left_panel_frame.minimumWidth() == 0  # pyright: ignore[reportPrivateUsage]
        assert window._right_panel_frame.minimumWidth() == 0  # pyright: ignore[reportPrivateUsage]
        collapsed_plan = window._panel_width_plan(  # pyright: ignore[reportPrivateUsage]
            left_visible=False,
            right_visible=False,
        )
        assert center_column.minimumWidth() == collapsed_plan.center
    finally:
        window.close()


def test_batch_panel_round_trips_keep_frame_reachable_and_restore_collapsed_rect(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 459,
            "gui.window.window_y": -10,
            "gui.window.default_width": 476,
            "gui.window.default_height": 796,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        screen = window.screen() or qapp.primaryScreen()
        assert screen is not None
        work_area = screen.availableGeometry()
        initial_normal = (window.pos(), window.size())
        initial_frame = window.frameGeometry()
        assert work_area.contains(initial_frame), (initial_normal, initial_frame, work_area)

        for _ in range(3):
            window._view_model.set_mode("batch")  # pyright: ignore[reportPrivateUsage]
            qapp.processEvents()
            batch_normal = (window.pos(), window.size())
            batch_frame = window.frameGeometry()
            assert work_area.contains(batch_frame), (
                initial_normal,
                initial_frame,
                batch_normal,
                batch_frame,
                work_area,
            )

            window._view_model.set_mode("single")  # pyright: ignore[reportPrivateUsage]
            qapp.processEvents()
            assert (window.pos(), window.size()) == initial_normal
            assert work_area.contains(window.frameGeometry())

        for _ in range(3):
            window._view_model.ui_projection_changed.emit(_VISIBLE_RIGHT)  # pyright: ignore[reportPrivateUsage]
            qapp.processEvents()
            assert work_area.contains(window.frameGeometry())
            window._view_model.ui_projection_changed.emit(_HIDDEN)  # pyright: ignore[reportPrivateUsage]
            qapp.processEvents()
            assert (window.pos(), window.size()) == initial_normal
            assert work_area.contains(window.frameGeometry())
    finally:
        window.close()


def test_atomic_normal_frame_geometry_projects_native_margins_to_set_geometry(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    window, _controller = _make_window(qapp, {})

    try:
        monkeypatch.setattr(window, "geometry", lambda: QRect(100, 120, 500, 400))
        monkeypatch.setattr(window, "frameGeometry", lambda: QRect(90, 80, 520, 460))
        calls: list[tuple[int, int, int, int]] = []
        monkeypatch.setattr(window, "setGeometry", lambda x, y, width, height: calls.append((x, y, width, height)))

        window._set_normal_frame_geometry(  # pyright: ignore[reportPrivateUsage]
            WindowRect(x=200, y=300, width=600, height=500),
        )

        assert calls == [(210, 340, 600, 500)]
    finally:
        window.close()


def test_internal_panel_round_trips_save_the_original_collapsed_geometry(
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
            "gui.window.center_panel_screen_x": 459,
            "gui.window.window_y": 40,
            "gui.window.default_width": 476,
            "gui.window.default_height": 760,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        initial_pos = window.pos()
        initial_size = window.size()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        initial_anchor = center_column.mapToGlobal(QPoint(0, 0)).x()

        for _ in range(3):
            window._view_model.set_mode("batch")  # pyright: ignore[reportPrivateUsage]
            qapp.processEvents()
            window._view_model.set_mode("single")  # pyright: ignore[reportPrivateUsage]
            qapp.processEvents()

        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]
        saved = controller.config_port.set_many_calls[-1]
        assert saved == {
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": initial_anchor,
            "gui.window.window_y": initial_pos.y(),
            "gui.window.default_width": initial_size.width(),
            "gui.window.default_height": initial_size.height(),
        }
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_normal_panel_transition_commits_top_level_geometry_once(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 700,
            "gui.window.default_width": 476,
            "gui.window.default_height": 760,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        fit = window._screen_fit_rect(window.screen())  # pyright: ignore[reportPrivateUsage]
        if fit is None or fit.width < window.width() + window._RIGHT_PANEL_MIN_WIDTH + 8:  # pyright: ignore[reportPrivateUsage]
            pytest.skip("test screen cannot fit the configured right panel")

        original_set_geometry = window.setGeometry
        geometry_calls: list[tuple[Any, ...]] = []

        def record_geometry(*args: Any) -> None:
            geometry_calls.append(args)
            original_set_geometry(*args)

        def reject_split_geometry(*_args: object) -> None:
            raise AssertionError("panel transition must not split resize and move")

        monkeypatch.setattr(window, "setGeometry", record_geometry)
        monkeypatch.setattr(window, "resize", reject_split_geometry)
        monkeypatch.setattr(window, "move", reject_split_geometry)

        window._view_model.ui_projection_changed.emit(_VISIBLE_RIGHT)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        assert len(geometry_calls) == 1
    finally:
        window.close()


def test_normal_panel_transition_preserves_pre_visibility_vertical_geometry(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 700,
            "gui.window.default_width": 476,
            "gui.window.default_height": 760,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        fit = window._screen_fit_rect(window.screen())  # pyright: ignore[reportPrivateUsage]
        if fit is None or fit.width < window.width() + window._RIGHT_PANEL_MIN_WIDTH + 8:  # pyright: ignore[reportPrivateUsage]
            pytest.skip("test screen cannot fit the configured right panel")
        assert fit is not None

        window.setGeometry(
            window.x(),
            fit.y + 40,
            window.width(),
            max(window.minimumHeight(), fit.height - 120),
        )
        qapp.processEvents()
        before = window.geometry()
        right_panel = window._right_panel_frame  # pyright: ignore[reportPrivateUsage]
        original_set_visible = right_panel.setVisible

        def perturb_height_after_visibility_change(visible: bool) -> None:
            original_set_visible(visible)
            window.resize(window.width(), window.height() + 37)

        monkeypatch.setattr(right_panel, "setVisible", perturb_height_after_visibility_change)
        window._view_model.ui_projection_changed.emit(_VISIBLE_RIGHT)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        assert abs(window.y() - before.y()) <= 8
        assert window.height() == before.height()
    finally:
        window.close()


def test_file_clear_content_reset_and_panel_close_share_one_geometry_transaction(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 700,
            "gui.window.default_width": 476,
            "gui.window.default_height": 760,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        fit = window._screen_fit_rect(window.screen())  # pyright: ignore[reportPrivateUsage]
        if fit is None or fit.width < window.width() + window._RIGHT_PANEL_MIN_WIDTH + 8:  # pyright: ignore[reportPrivateUsage]
            pytest.skip("test screen cannot fit the configured right panel")

        window._view_model.ui_projection_changed.emit(_VISIBLE_RIGHT)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        before = window.geometry()
        original_reset = window._action_area_vm.reset  # pyright: ignore[reportPrivateUsage]

        def perturb_height_during_content_reset() -> None:
            original_reset()
            window.resize(window.width(), window.height() + 37)

        monkeypatch.setattr(window._action_area_vm, "reset", perturb_height_during_content_reset)  # pyright: ignore[reportPrivateUsage]
        window._on_files_cleared()  # pyright: ignore[reportPrivateUsage]
        window._view_model.ui_projection_changed.emit(_HIDDEN)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        assert window.height() == before.height()
        assert window.updatesEnabled() is True
    finally:
        window.close()


def test_history_row_does_not_raise_window_height_when_right_panel_closes(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 700,
            "gui.window.default_width": 476,
            "gui.window.default_height": 760,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        fit = window._screen_fit_rect(window.screen())  # pyright: ignore[reportPrivateUsage]
        if fit is None or fit.width < window.width() + window._RIGHT_PANEL_MIN_WIDTH + 8:  # pyright: ignore[reportPrivateUsage]
            pytest.skip("test screen cannot fit the configured right panel")

        window._view_model.ui_projection_changed.emit(_VISIBLE_RIGHT)  # pyright: ignore[reportPrivateUsage]
        window._info_area_vm.add_message(  # pyright: ignore[reportPrivateUsage]
            "Received file from another instance: README.md",
            "info",
            show_location=True,
            file_path="README.md",
        )
        qapp.processEvents()
        before = window.geometry()

        window._on_files_cleared()  # pyright: ignore[reportPrivateUsage]
        window._view_model.ui_projection_changed.emit(_HIDDEN)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        assert window.height() == before.height()
    finally:
        window.close()


def test_right_slot_switch_does_not_move_or_resize_the_window(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.default_width": 460,
            "gui.window.default_height": 760,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        window._view_model.ui_projection_changed.emit(_VISIBLE_RIGHT)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        before = (window.pos(), window.size())
        template_projection = MainWindowUiProjection(
            left_panel_visible=False,
            right_panel_visible=True,
            right_panel_slot=RightPanelSlot.TEMPLATE,
            center_action_visible=True,
            info_area_visible=True,
            conversion_context=None,
            template_context=None,
        )

        window._view_model.ui_projection_changed.emit(template_projection)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()

        assert (window.pos(), window.size()) == before
    finally:
        window.close()


def test_state_save_is_skipped_when_remember_policy_is_disabled(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    window, controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": False,
            "gui.window.auto_center": False,
        },
    )
    try:
        window.move(111, 222)
        window.resize(777, 766)

        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        assert controller.config_port.set_calls == []
        assert controller.config_port.set_many_calls == []
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_state_save_uses_one_sparse_write_when_remember_policy_is_enabled(
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
        window.move(111, 222)
        window.resize(777, 766)
        actual_size = window.size()

        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        assert controller.config_port.set_calls == []
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        center_screen_x = center_column.mapToGlobal(QPoint(0, 0)).x()
        assert controller.config_port.set_many_calls == [
            {
                "gui.window.geometry_schema_version": 2,
                "gui.window.center_panel_screen_x": center_screen_x,
                "gui.window.window_y": 222,
                "gui.window.default_width": actual_size.width(),
                "gui.window.default_height": actual_size.height(),
            }
        ]
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()
