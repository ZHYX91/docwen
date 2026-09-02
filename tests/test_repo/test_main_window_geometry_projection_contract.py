"""Guards for VIS-2026-07-19-139 main-window geometry evidence."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def _read(relative_path: str) -> str:
    return read_source_text(PROJECT_ROOT / relative_path)


def test_geometry_schema_v2_and_recovery_stay_declared() -> None:
    config_path = PROJECT_ROOT / "configs/gui.toml"
    config = tomllib.loads(config_path.read_text(encoding="utf-8"))["window"]
    geometry = _read("packages/apps/gui/src/docwen_gui/window_geometry.py")
    regression = _read("packages/apps/gui/tests/test_window_geometry.py")

    assert config == config | {
        "geometry_schema_version": 2,
        "center_panel_screen_x": 420,
        "window_y": 0,
        "default_width": 476,
        "default_height": 860,
        "min_width": 420,
        "min_height": 720,
    }
    for token in (
        "GEOMETRY_SCHEMA_VERSION = 2",
        "DEFAULT_CENTER_PANEL_SCREEN_X = 420",
        "DEFAULT_WINDOW_Y = 0",
        "DEFAULT_WINDOW_WIDTH = 476",
        "DEFAULT_WINDOW_HEIGHT = 860",
        "DEFAULT_MIN_WIDTH = 420",
        "DEFAULT_MIN_HEIGHT = 720",
        "def _read_schema_version(",
        "def load_window_geometry_policy(",
        "def build_canonical_geometry_values(",
        "def recover_window_geometry(",
        "def normalize_ui_scale(",
    ):
        assert token in geometry

    canonical_writer = geometry.split("def build_canonical_geometry_values", 1)[1].split("def _intersection_area", 1)[0]
    for canonical_key in (
        "gui.window.geometry_schema_version",
        "gui.window.center_panel_screen_x",
        "gui.window.window_y",
        "gui.window.default_width",
        "gui.window.default_height",
    ):
        assert canonical_key in canonical_writer

    for test_name in (
        "test_invalid_canonical_values_use_declared_defaults",
        "test_unknown_future_schema_uses_safe_defaults_and_disables_save_contract",
        "test_schema_version_read_error_fails_closed_instead_of_using_configured_values",
        "test_explicit_invalid_schema_version_is_not_treated_as_missing",
        "test_canonical_scale_round_trip_is_symmetric",
        "test_negative_coordinate_secondary_screen_is_preserved",
        "test_disconnected_screen_geometry_centers_on_nearest_remaining_work_area",
        "test_oversized_geometry_and_minima_fit_a_tiny_work_area",
        "test_real_sparse_config_port_round_trips_schema_v2_geometry",
    ):
        assert f"def {test_name}(" in regression


def test_main_window_geometry_lifecycle_stays_frame_and_anchor_safe() -> None:
    main = _read("packages/apps/gui/src/docwen_gui/main_window.py")
    regression = _read("packages/apps/gui/tests/test_main_window_window_behavior_*.py")

    for token in (
        "QApplication.screens()",
        "screen.availableGeometry()",
        "self.frameGeometry()",
        "QTimer.singleShot(0, self._finalize_shown_window_restore)",
        "spontaneous=event.spontaneous()",
        "QTimer.singleShot(16, self._force_pending_normal_center_anchor)",
        "self._last_normal_window_rect",
        "self._last_normal_center_offset",
        "if not self._persisted_window_geometry_policy.schema_supported:",
        "if self.isVisible() and not self._shown_window_restore_finalized:",
        "values = build_canonical_geometry_values(",
        "if cfg_port.set_many(values):",
        "if self._geometry_source_change_rect == normal_rect:",
        "recovered = recover_window_geometry(",
    ):
        assert token in main

    for test_name in (
        "test_startup_restores_declared_canonical_geometry",
        "test_real_canonical_anchor_is_stable_across_shown_save_reload_cycles",
        "test_immediate_show_close_does_not_save_before_native_restore_finalizes",
        "test_startup_recovers_disconnected_schema_v2_geometry_to_available_screen",
        "test_shown_oversized_geometry_fits_native_frame_inside_work_area",
        "test_visible_panel_transitions_resize_by_configured_width_and_preserve_center",
        "test_normal_panel_transition_commits_top_level_geometry_once",
        "test_fresh_empty_batch_restores_base_width_plus_left_panel",
        "test_maximized_save_uses_paired_normal_frame_rect_and_anchor",
        "test_panel_change_while_maximized_preserves_normal_anchor_for_save_and_restore",
        "test_move_during_queued_normal_restore_cancels_the_pending_panel_anchor",
        "test_save_during_queued_normal_restore_uses_the_paired_normal_rect",
        "test_one_pixel_native_normal_rect_difference_does_not_strand_panel_anchor",
        "test_reverse_order_normal_restore_uses_bounded_quiet_fallback",
        "test_unknown_future_geometry_schema_is_not_downgraded_on_close",
        "test_geometry_source_reset_is_not_undone_on_close_without_a_later_move",
    ):
        assert f"def {test_name}(" in regression
