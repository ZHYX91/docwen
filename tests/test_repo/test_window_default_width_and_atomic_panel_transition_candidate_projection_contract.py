from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]


def _read(relative_path: str) -> str:
    return (ROOT / relative_path).read_text(encoding="utf-8")


def test_current_geometry_contract_matches_the_reference_center_width() -> None:
    config = _read("configs/gui.toml")
    geometry = _read("packages/apps/gui/src/docwen_gui/window_geometry.py")

    assert "geometry_schema_version = 2" in config
    assert "default_width = 476" in config
    assert "center_panel_width = 460" in config
    assert "GEOMETRY_SCHEMA_VERSION = 2" in geometry
    assert "DEFAULT_WINDOW_WIDTH = 476" in geometry


def test_visible_panel_transition_has_one_top_level_geometry_commit() -> None:
    main_window = _read("packages/apps/gui/src/docwen_gui/main_window.py")
    info_area = _read("packages/apps/gui/src/docwen_gui/widgets/info_area.py")

    assert "def _normal_panel_transition_rect(" in main_window
    assert "self.setUpdatesEnabled(False)" in main_window
    assert "self._set_normal_frame_geometry(rect)" in main_window
    assert "def _set_normal_frame_geometry(" in main_window
    assert "self.setGeometry(client_x, client_y, rect.width, rect.height)" in main_window
    assert "self._collapsed_normal_window_rect" in main_window
    assert "def _context_panel_width_contribution(" in main_window
    assert "width=recovered_collapsed.rect.width + contribution" in main_window
    assert "QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Ignored" in info_area
