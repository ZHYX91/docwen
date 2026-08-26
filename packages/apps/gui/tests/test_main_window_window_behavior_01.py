"""Focused tests split from test_main_window_window_behavior.py."""

from __future__ import annotations

from ._main_window_window_behavior_support import (
    _HIDDEN,
    _VISIBLE_BOTH,
    _VISIBLE_LEFT,
    _VISIBLE_RIGHT,
    QPoint,
    QWidget,
    _ConfigPort,
    _make_window,
    _root_grid,
    pytest,
)

pytestmark = pytest.mark.gui


def test_general_model_fallback_matches_shipped_auto_center_default() -> None:
    from docwen_gui.models.settings_config import SettingsConfig

    assert SettingsConfig().gui.auto_center is False


def test_window_behavior_policy_falls_back_for_missing_broken_or_untyped_ports() -> None:
    from docwen_gui.window_behavior import (
        DEFAULT_WINDOW_BEHAVIOR,
        load_window_behavior_policy,
    )

    class _BrokenPort:
        def get(self, _key: str, _default: object = None) -> object:
            raise RuntimeError("unavailable")

    class _UntypedPort:
        def get(self, key: str, _default: object = None) -> object:
            return {
                "gui.window.remember_gui_state": 1,
                "gui.window.auto_center": "false",
                "gui.window.expand_side_panels": None,
            }[key]

    assert load_window_behavior_policy(None) == DEFAULT_WINDOW_BEHAVIOR
    assert load_window_behavior_policy(_BrokenPort()) == DEFAULT_WINDOW_BEHAVIOR  # type: ignore[arg-type]
    assert load_window_behavior_policy(_UntypedPort()) == DEFAULT_WINDOW_BEHAVIOR  # type: ignore[arg-type]


def test_window_behavior_policy_reads_bounded_configured_panel_widths() -> None:
    from docwen_gui.window_behavior import load_window_behavior_policy

    valid = _ConfigPort(
        {
            "gui.window.center_panel_width": 520,
            "gui.window.left_panel_width": 410,
            "gui.window.right_panel_width": 320,
        }
    )
    policy = load_window_behavior_policy(valid)
    assert policy.center_panel_width == 520
    assert policy.left_panel_width == 410
    assert policy.right_panel_width == 320

    invalid = _ConfigPort(
        {
            "gui.window.center_panel_width": True,
            "gui.window.left_panel_width": 20,
            "gui.window.right_panel_width": "500",
        }
    )
    fallback = load_window_behavior_policy(invalid)
    assert fallback.center_panel_width == 460
    assert fallback.left_panel_width == 400
    assert fallback.right_panel_width == 300


def test_panel_width_plan_keeps_hidden_columns_inert_and_compresses_narrow_viewports() -> None:
    from docwen_gui.window_behavior import DEFAULT_WINDOW_BEHAVIOR, PanelWidthPlan, plan_panel_widths

    center_only = plan_panel_widths(
        DEFAULT_WINDOW_BEHAVIOR,
        container_width=476,
        horizontal_margins=16,
        spacing=8,
        left_visible=False,
        right_visible=False,
    )
    assert center_only == PanelWidthPlan(left=0, center=460, right=0)

    three_columns = plan_panel_widths(
        DEFAULT_WINDOW_BEHAVIOR,
        container_width=796,
        horizontal_margins=16,
        spacing=8,
        left_visible=True,
        right_visible=True,
    )
    assert three_columns == PanelWidthPlan(left=263, center=303, right=198)
    assert sum((three_columns.left, three_columns.center, three_columns.right)) == 796 - 16 - (2 * 8)


def test_panel_width_plan_scales_semantic_intent_before_allocating_dpi_viewport() -> None:
    from docwen_gui.window_behavior import DEFAULT_WINDOW_BEHAVIOR, PanelWidthPlan, plan_panel_widths

    plan = plan_panel_widths(
        DEFAULT_WINDOW_BEHAVIOR,
        container_width=1_482,
        horizontal_margins=16,
        spacing=8,
        left_visible=True,
        right_visible=True,
        scale_factor=1.25,
    )

    assert plan == PanelWidthPlan(left=500, center=575, right=375)


def test_fresh_window_exposes_the_reference_center_width_inside_current_margins(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_width": 460,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        assert window.width() == 476
        assert center_column.width() == 460
    finally:
        window.close()


def test_fresh_empty_batch_restores_base_width_plus_left_panel(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.default_mode": "batch",
            "gui.window.center_panel_screen_x": 700,
            "gui.window.default_width": 476,
            "gui.window.default_height": 760,
            "gui.window.center_panel_width": 460,
            "gui.window.left_panel_width": 300,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        fit = window._screen_fit_rect(window.screen())  # pyright: ignore[reportPrivateUsage]
        expected_width = 476 + 300 + _root_grid(window).horizontalSpacing()
        if fit is None or fit.width < expected_width:
            pytest.skip("test screen cannot fit the configured empty-batch layout")
        assert window._left_panel_frame.isVisible() is True  # pyright: ignore[reportPrivateUsage]
        assert window.width() == expected_width
        assert center_column.width() == 460
    finally:
        window.close()


def test_main_window_reads_shipped_policy_through_real_config_port(
    qapp,
    tmp_path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from pathlib import Path

    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_gui.main_window import MainWindow
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel
    from docwen_gui.window_behavior import DEFAULT_WINDOW_BEHAVIOR

    project_configs = Path(__file__).resolve().parents[4] / "configs"
    controller = ApplicationController(
        config_port=ConfigPortAdapter(base_dir=project_configs, user_dir=tmp_path / "configs")
    )
    window = MainWindow(view_model=MainWindowViewModel(controller=controller))
    try:
        assert window._window_behavior == DEFAULT_WINDOW_BEHAVIOR  # pyright: ignore[reportPrivateUsage]
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_real_general_reset_survives_main_window_close(
    qapp,
    tmp_path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from pathlib import Path

    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_gui.main_window import MainWindow
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

    project_configs = Path(__file__).resolve().parents[4] / "configs"
    config_port = ConfigPortAdapter(base_dir=project_configs, user_dir=tmp_path / "configs")
    assert config_port.set_many(
        {
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 650,
            "gui.window.window_y": 20,
            "gui.window.default_width": 560,
            "gui.window.default_height": 760,
        }
    )
    controller = ApplicationController(config_port=config_port)
    window = MainWindow(view_model=MainWindowViewModel(controller=controller))
    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    try:
        assert config_port.reset_group("general") is True
        window._apply_runtime_window_settings()  # pyright: ignore[reportPrivateUsage]

        window._save_gui_state()  # pyright: ignore[reportPrivateUsage]

        assert config_port.get("gui.window.geometry_schema_version") == 2
        assert config_port.get("gui.window.center_panel_screen_x") == 420
        assert config_port.get("gui.window.window_y") == 0
        assert config_port.get("gui.window.default_width") == 476
        assert config_port.get("gui.window.default_height") == 860
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


@pytest.mark.parametrize(
    ("remember_gui_state", "auto_center", "expected_size"),
    [
        (False, False, (476, 860)),
        (True, True, (1_200, 1_000)),
        (False, True, (476, 860)),
    ],
)
def test_startup_center_policy_retains_size_only_when_state_is_remembered(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
    remember_gui_state: bool,
    auto_center: bool,
    expected_size: tuple[int, int],
) -> None:
    from docwen_gui.main_window import MainWindow

    centered: list[bool] = []

    def _center(window: MainWindow) -> None:
        centered.append(True)
        window.move(17, 29)

    monkeypatch.setattr(MainWindow, "_center_on_screen", _center)
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": remember_gui_state,
            "gui.window.auto_center": auto_center,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 123,
            "gui.window.window_y": 234,
            "gui.window.default_width": 1_200,
            "gui.window.default_height": 1_000,
        },
    )
    try:
        assert centered == [True]
        assert window.pos().x() == 17
        assert window.pos().y() == 29
        assert (window.size().width(), window.size().height()) == expected_size
    finally:
        monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
        window.close()


def test_startup_policy_restores_when_remembered_without_auto_center(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 123,
            "gui.window.window_y": 20,
            "gui.window.default_width": 640,
            "gui.window.default_height": 760,
        },
    )
    try:
        assert window.pos().x() + window._center_column_offset() == 123  # pyright: ignore[reportPrivateUsage]
        assert window.pos().y() == 20
        assert window.size().width() == 640
        assert window.size().height() == 760
    finally:
        window.close()


def test_startup_restores_declared_canonical_geometry(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 333,
            "gui.window.window_y": 20,
            "gui.window.default_width": 460,
            "gui.window.default_height": 760,
            "gui.window.min_width": 420,
            "gui.window.min_height": 500,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        center_screen_x = center_column.mapToGlobal(QPoint(0, 0)).x()
        assert center_screen_x == 333
        assert window.pos().y() == 20
        assert window.size().width() == 460
        assert window.size().height() == 760
    finally:
        window.close()


def test_real_canonical_anchor_is_stable_across_shown_save_reload_cycles(
    qapp,
    tmp_path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from pathlib import Path

    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_gui.main_window import MainWindow
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

    project_configs = Path(__file__).resolve().parents[4] / "configs"
    config_port = ConfigPortAdapter(base_dir=project_configs, user_dir=tmp_path / "configs")
    assert config_port.set_many(
        {
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 333,
            "gui.window.window_y": 20,
            "gui.window.default_width": 460,
            "gui.window.default_height": 760,
        }
    )

    observed_anchors: list[int] = []
    for _ in range(3):
        controller = ApplicationController(config_port=config_port)
        window = MainWindow(view_model=MainWindowViewModel(controller=controller))
        monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
        try:
            window.show()
            qapp.processEvents()
            center_column = window.findChild(QWidget, "mainWindowCenterColumn")
            assert center_column is not None
            observed_anchors.append(center_column.mapToGlobal(QPoint(0, 0)).x())
            window._save_gui_state()  # pyright: ignore[reportPrivateUsage]
        finally:
            monkeypatch.setenv("DOCWEN_GUI_DISABLE_STATE_SAVE", "1")
            window.close()

    assert observed_anchors == [333, 333, 333]
    assert config_port.get("gui.window.center_panel_screen_x") == 333


def test_immediate_show_close_does_not_save_before_native_restore_finalizes(
    qapp,
    tmp_path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from pathlib import Path

    from docwen_application.controller import ApplicationController
    from docwen_bundle.config_port import ConfigPortAdapter
    from docwen_gui.main_window import MainWindow
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

    project_configs = Path(__file__).resolve().parents[4] / "configs"
    config_port = ConfigPortAdapter(base_dir=project_configs, user_dir=tmp_path / "configs")
    assert config_port.set_many(
        {
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 333,
            "gui.window.window_y": 20,
            "gui.window.default_width": 460,
            "gui.window.default_height": 760,
        }
    )

    monkeypatch.delenv("DOCWEN_GUI_DISABLE_STATE_SAVE", raising=False)
    for _ in range(3):
        controller = ApplicationController(config_port=config_port)
        window = MainWindow(view_model=MainWindowViewModel(controller=controller))
        window.show()
        window.close()

    assert config_port.get("gui.window.center_panel_screen_x") == 333
    assert config_port.get("gui.window.window_y") == 20


def test_startup_recovers_disconnected_schema_v2_geometry_to_available_screen(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 100_000,
            "gui.window.window_y": 100_000,
            "gui.window.default_width": 640,
            "gui.window.default_height": 760,
        },
    )
    try:
        screen = qapp.primaryScreen()
        assert screen is not None
        available = screen.availableGeometry()
        geometry = window.geometry()
        assert available.contains(geometry.topLeft())
        assert available.contains(geometry.bottomRight())
    finally:
        window.close()


def test_shown_oversized_geometry_fits_native_frame_inside_work_area(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 50_000,
            "gui.window.window_y": 50_000,
            "gui.window.default_width": 2_000,
            "gui.window.default_height": 1_500,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        screen = qapp.primaryScreen()
        assert screen is not None
        assert screen.availableGeometry().contains(window.frameGeometry())
    finally:
        window.close()


def test_visible_panel_transitions_resize_by_configured_width_and_preserve_center(qapp) -> None:
    window, _controller = _make_window(
        qapp,
        {
            "gui.window.remember_gui_state": True,
            "gui.window.auto_center": False,
            "gui.window.geometry_schema_version": 2,
            "gui.window.center_panel_screen_x": 350,
            "gui.window.window_y": 20,
            "gui.window.default_width": 460,
            "gui.window.default_height": 760,
            "gui.window.center_panel_width": 460,
            "gui.window.left_panel_width": 400,
            "gui.window.right_panel_width": 300,
        },
    )
    try:
        window.show()
        qapp.processEvents()
        center_column = window.findChild(QWidget, "mainWindowCenterColumn")
        assert center_column is not None
        expected_anchor = center_column.mapToGlobal(QPoint(0, 0)).x()
        expected_center_width = center_column.width()
        center_only_width = window.width()
        fit = window._screen_fit_rect(window.screen())  # pyright: ignore[reportPrivateUsage]
        grid = _root_grid(window)
        spacing = grid.horizontalSpacing()
        full_width = center_only_width + 400 + 300 + (2 * spacing)
        has_room_for_full_three_column_layout = fit is not None and fit.width >= full_width

        window._view_model.ui_projection_changed.emit(_VISIBLE_RIGHT)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        if has_room_for_full_three_column_layout:
            assert window.width() == center_only_width + 300 + spacing
            assert center_column.mapToGlobal(QPoint(0, 0)).x() == expected_anchor
            assert center_column.width() == expected_center_width

        window._view_model.ui_projection_changed.emit(_VISIBLE_BOTH)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        if has_room_for_full_three_column_layout:
            assert window.width() == full_width
            assert center_column.mapToGlobal(QPoint(0, 0)).x() == expected_anchor
            assert center_column.width() == expected_center_width

        window._view_model.ui_projection_changed.emit(_VISIBLE_LEFT)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        if has_room_for_full_three_column_layout:
            assert window.width() == center_only_width + 400 + spacing
            assert center_column.mapToGlobal(QPoint(0, 0)).x() == expected_anchor
            assert center_column.width() == expected_center_width

        window._view_model.ui_projection_changed.emit(_HIDDEN)  # pyright: ignore[reportPrivateUsage]
        qapp.processEvents()
        if has_room_for_full_three_column_layout:
            assert window.width() == center_only_width
            assert center_column.mapToGlobal(QPoint(0, 0)).x() == expected_anchor
            assert center_column.width() == expected_center_width
    finally:
        window.close()
