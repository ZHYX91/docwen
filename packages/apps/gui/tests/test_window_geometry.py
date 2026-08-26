from __future__ import annotations

from math import inf, nan
from pathlib import Path

import pytest

from docwen_gui.window_geometry import (
    DEFAULT_MIN_HEIGHT,
    DEFAULT_MIN_WIDTH,
    DEFAULT_WINDOW_HEIGHT,
    DEFAULT_WINDOW_WIDTH,
    ScreenRect,
    WindowRect,
    build_canonical_geometry_values,
    center_window_geometry,
    load_window_geometry_policy,
    load_window_scale_factor,
    normalize_ui_scale,
    recover_window_geometry,
)

pytestmark = pytest.mark.gui


class _ConfigPort:
    def __init__(self, values: dict[str, object]) -> None:
        self.values = dict(values)

    def get(self, key: str, default: object = None) -> object:
        return self.values.get(key, default)


def _identity(value: int) -> int:
    return value


def test_current_default_window_width_owns_outer_margins_around_the_reference_center() -> None:
    assert DEFAULT_WINDOW_WIDTH == 476


def test_canonical_geometry_uses_anchor_and_scales_signed_coordinates() -> None:
    policy = load_window_geometry_policy(
        _ConfigPort(
            {
                "gui.window.geometry_schema_version": 2,
                "gui.window.center_panel_screen_x": -700,
                "gui.window.window_y": -120,
                "gui.window.default_width": 640,
                "gui.window.default_height": 760,
                "gui.window.min_width": 400,
                "gui.window.min_height": 500,
            }
        ),
        center_offset=24,
        scale_value=lambda value: round(value * 1.5),
    )

    assert policy.source == "canonical"
    assert policy.rect == WindowRect(x=-1074, y=-180, width=960, height=1140)
    assert policy.min_width == 600
    assert policy.min_height == 750


def test_invalid_canonical_values_use_declared_defaults() -> None:
    policy = load_window_geometry_policy(
        _ConfigPort(
            {
                "gui.window.geometry_schema_version": 2,
                "gui.window.center_panel_screen_x": nan,
                "gui.window.window_y": True,
                "gui.window.default_width": -1,
                "gui.window.default_height": inf,
                "gui.window.min_width": 0,
                "gui.window.min_height": False,
            }
        ),
        center_offset=20,
        scale_value=_identity,
    )

    assert policy.rect == WindowRect(
        x=400,
        y=0,
        width=DEFAULT_WINDOW_WIDTH,
        height=DEFAULT_WINDOW_HEIGHT,
    )
    assert policy.min_width == DEFAULT_MIN_WIDTH
    assert policy.min_height == DEFAULT_MIN_HEIGHT


def test_unknown_future_schema_uses_safe_defaults_and_disables_save_contract() -> None:
    policy = load_window_geometry_policy(
        _ConfigPort(
            {
                "gui.window.geometry_schema_version": 999,
                "gui.window.center_panel_screen_x": 1_000,
                "gui.window.window_y": 500,
                "gui.window.default_width": 1_200,
                "gui.window.default_height": 900,
                "gui.window.min_width": 9_999,
                "gui.window.min_height": 8_888,
            }
        ),
        center_offset=20,
        scale_value=_identity,
    )

    assert policy.schema_version == 999
    assert policy.schema_supported is False
    assert policy.rect == WindowRect(x=400, y=0, width=476, height=860)
    assert policy.min_width == DEFAULT_MIN_WIDTH
    assert policy.min_height == DEFAULT_MIN_HEIGHT


def test_schema_version_read_error_fails_closed_instead_of_using_configured_values() -> None:
    class _VersionReadError(_ConfigPort):
        def get(self, key: str, default: object = None) -> object:
            if key == "gui.window.geometry_schema_version":
                raise OSError("transient read failure")
            return super().get(key, default)

    policy = load_window_geometry_policy(
        _VersionReadError(
            {
                "gui.window.center_panel_screen_x": 9_999,
                "gui.window.window_y": 8_888,
                "gui.window.default_width": 1_200,
                "gui.window.default_height": 900,
            }
        ),
        center_offset=20,
        scale_value=_identity,
    )

    assert policy.schema_version is None
    assert policy.schema_supported is False
    assert policy.source == "canonical"
    assert policy.rect == WindowRect(x=400, y=0, width=476, height=860)


def test_missing_schema_version_ignores_unversioned_persisted_geometry() -> None:
    policy = load_window_geometry_policy(
        _ConfigPort(
            {
                "gui.window.center_panel_screen_x": 9_999,
                "gui.window.window_y": 8_888,
                "gui.window.default_width": 1_200,
                "gui.window.default_height": 900,
            }
        ),
        center_offset=20,
        scale_value=_identity,
    )

    assert policy.schema_version is None
    assert policy.schema_supported is False
    assert policy.rect == WindowRect(x=400, y=0, width=476, height=860)


@pytest.mark.parametrize("version", ["999", True, 0, -1, 1.0, 1.4])
def test_explicit_invalid_schema_version_is_not_treated_as_missing(version: object) -> None:
    policy = load_window_geometry_policy(
        _ConfigPort(
            {
                "gui.window.geometry_schema_version": version,
                "gui.window.center_panel_screen_x": 777,
                "gui.window.window_y": 33,
                "gui.window.default_width": 600,
                "gui.window.default_height": 760,
                "gui.window.min_width": 9_999,
                "gui.window.min_height": 8_888,
            }
        ),
        center_offset=20,
        scale_value=_identity,
    )

    assert policy.schema_version is None
    assert policy.schema_supported is False
    assert policy.rect == WindowRect(x=400, y=0, width=476, height=860)
    assert policy.min_width == DEFAULT_MIN_WIDTH
    assert policy.min_height == DEFAULT_MIN_HEIGHT


@pytest.mark.parametrize("factor", [1.0, 1.25, 1.5, 2.0])
def test_canonical_scale_round_trip_is_symmetric(factor: float) -> None:
    policy = load_window_geometry_policy(
        _ConfigPort(
            {
                "gui.window.geometry_schema_version": 2,
                "gui.window.center_panel_screen_x": 480,
                "gui.window.window_y": 64,
                "gui.window.default_width": 640,
                "gui.window.default_height": 760,
            }
        ),
        center_offset=round(16 * factor),
        scale_value=lambda value: round(value * factor),
    )
    values = build_canonical_geometry_values(
        policy.rect,
        center_offset=round(16 * factor),
        unscale_value=lambda value: round(value / factor),
    )

    assert values == {
        "gui.window.geometry_schema_version": 2,
        "gui.window.center_panel_screen_x": 480,
        "gui.window.window_y": 64,
        "gui.window.default_width": 640,
        "gui.window.default_height": 760,
    }


def test_negative_coordinate_secondary_screen_is_preserved() -> None:
    recovered = recover_window_geometry(
        WindowRect(x=-1800, y=100, width=900, height=700),
        [
            ScreenRect(x=0, y=0, width=1920, height=1040),
            ScreenRect(x=-1920, y=0, width=1920, height=1040),
        ],
        min_width=420,
        min_height=720,
    )

    assert recovered.screen == ScreenRect(x=-1920, y=0, width=1920, height=1040)
    assert recovered.rect == WindowRect(x=-1800, y=100, width=900, height=720)
    assert recovered.disconnected is False


def test_largest_intersection_selects_the_observed_screen() -> None:
    recovered = recover_window_geometry(
        WindowRect(x=1700, y=100, width=900, height=700),
        [
            ScreenRect(x=0, y=0, width=1920, height=1040),
            ScreenRect(x=1920, y=0, width=1280, height=1024),
        ],
        min_width=420,
        min_height=500,
    )

    assert recovered.screen == ScreenRect(x=1920, y=0, width=1280, height=1024)
    assert recovered.rect == WindowRect(x=1920, y=100, width=900, height=700)


def test_disconnected_screen_geometry_centers_on_nearest_remaining_work_area() -> None:
    recovered = recover_window_geometry(
        WindowRect(x=5000, y=100, width=800, height=700),
        [
            ScreenRect(x=0, y=0, width=1920, height=1040),
            ScreenRect(x=1920, y=0, width=1280, height=1024),
        ],
        min_width=420,
        min_height=500,
    )

    assert recovered.screen == ScreenRect(x=1920, y=0, width=1280, height=1024)
    assert recovered.rect == WindowRect(x=2160, y=162, width=800, height=700)
    assert recovered.disconnected is True


def test_extreme_integer_candidate_cannot_overflow_nearest_screen_selection() -> None:
    recovered = recover_window_geometry(
        WindowRect(x=10**300, y=-(10**300), width=640, height=760),
        [ScreenRect(x=0, y=0, width=800, height=800)],
        min_width=420,
        min_height=720,
    )

    assert recovered.rect == WindowRect(x=80, y=20, width=640, height=760)
    assert recovered.disconnected is True


def test_oversized_geometry_and_minima_fit_a_tiny_work_area() -> None:
    recovered = recover_window_geometry(
        WindowRect(x=-100, y=-100, width=2000, height=1500),
        [ScreenRect(x=50, y=60, width=300, height=200)],
        min_width=420,
        min_height=720,
    )

    assert recovered.rect == WindowRect(x=50, y=60, width=300, height=200)
    assert recovered.effective_min_width == 300
    assert recovered.effective_min_height == 200


def test_no_screen_keeps_coordinates_but_normalizes_size_to_minimum() -> None:
    recovered = recover_window_geometry(
        WindowRect(x=-50, y=20, width=0, height=-1),
        [],
        min_width=420,
        min_height=720,
    )

    assert recovered.rect == WindowRect(x=-50, y=20, width=420, height=720)
    assert recovered.screen is None


def test_centering_caps_size_and_minimum_to_work_area() -> None:
    recovered = center_window_geometry(
        width=1200,
        height=900,
        screen=ScreenRect(x=-1000, y=100, width=800, height=600),
        min_width=900,
        min_height=720,
    )

    assert recovered.rect == WindowRect(x=-1000, y=100, width=800, height=600)
    assert recovered.effective_min_width == 800
    assert recovered.effective_min_height == 600


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (0, None),
        (1.25, 1.25),
        (150, 1.5),
        ("200%", 2.0),
        (" 1.5 ", 1.5),
        (True, None),
        (nan, None),
        ("bad", None),
        (499, None),
    ],
)
def test_normalize_ui_scale(value: object, expected: float | None) -> None:
    assert normalize_ui_scale(value) == expected


@pytest.mark.parametrize(
    ("values", "detected", "expected"),
    [
        ({"gui.dpi.enable_dpi_scaling": False, "gui.dpi.ui_scale": 2}, 1.5, 1.0),
        ({"gui.dpi.enable_dpi_scaling": True, "gui.dpi.ui_scale": 150}, 1.25, 1.5),
        ({"gui.dpi.enable_dpi_scaling": True, "gui.dpi.ui_scale": 0}, 1.25, 1.25),
        ({"gui.dpi.enable_dpi_scaling": "false", "gui.dpi.ui_scale": 0}, 1.5, 1.5),
    ],
)
def test_load_window_scale_factor(
    values: dict[str, object],
    detected: float,
    expected: float,
) -> None:
    assert load_window_scale_factor(_ConfigPort(values), detected_factor=detected) == expected


def test_broken_port_falls_back_without_leaking_exception() -> None:
    class _BrokenPort:
        def get(self, key: str, default: object = None) -> object:
            del key, default
            raise RuntimeError("unavailable")

    policy = load_window_geometry_policy(
        _BrokenPort(),  # type: ignore[arg-type]
        center_offset=20,
        scale_value=_identity,
    )

    assert policy.rect == WindowRect(x=400, y=0, width=476, height=860)
    assert (
        load_window_scale_factor(  # type: ignore[arg-type]
            _BrokenPort(),
            detected_factor=1.25,
        )
        == 1.25
    )


def test_real_sparse_config_port_round_trips_schema_v2_geometry(tmp_path: Path) -> None:
    from docwen_bundle.config_port import ConfigPortAdapter

    project_configs = Path(__file__).resolve().parents[4] / "configs"
    port = ConfigPortAdapter(base_dir=project_configs, user_dir=tmp_path / "configs")

    shipped = load_window_geometry_policy(
        port,
        center_offset=0,
        scale_value=_identity,
    )
    assert shipped.source == "canonical"
    assert shipped.rect == WindowRect(x=420, y=0, width=476, height=860)

    expected = WindowRect(x=-1500, y=75, width=900, height=700)
    assert port.set_many(
        build_canonical_geometry_values(
            expected,
            center_offset=20,
            unscale_value=_identity,
        )
    )
    reloaded = load_window_geometry_policy(
        port,
        center_offset=20,
        scale_value=_identity,
    )
    assert reloaded.source == "canonical"
    assert reloaded.schema_version == 2
    assert reloaded.schema_supported is True
    assert reloaded.rect == expected
