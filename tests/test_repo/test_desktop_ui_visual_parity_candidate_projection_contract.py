"""Evidence guards for the VIS-2026-07-27-397 desktop UI candidate."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
CARD_NAME = "desktop-ui-visual-parity-stage-card-2026-07-27.md"


def _read(relative_path: str) -> str:
    return (ROOT / relative_path).read_text(encoding="utf-8")


def test_visual_candidate_keeps_the_confirmed_production_repairs() -> None:
    main = _read("packages/apps/gui/src/docwen_gui/main_window.py")
    input_area = _read("packages/apps/gui/src/docwen_gui/widgets/input_area.py")
    info_area = _read("packages/apps/gui/src/docwen_gui/widgets/info_area.py")
    batch_list = _read("packages/apps/gui/src/docwen_gui/widgets/batch_list.py")
    app = _read("packages/apps/gui/src/docwen_gui/app.py")
    bundle_entry = _read("packages/bundle/src/docwen_bundle/gui_entry.py")

    for token in (
        "self._CENTER_PANEL_MIN_WIDTH: int =",
        "self._LEFT_PANEL_MIN_WIDTH: int =",
        "self._RIGHT_PANEL_MIN_WIDTH: int =",
        "def _normal_panel_transition_rect(",
        'setObjectName("bottomBarLeftActions")',
        'setObjectName("bottomBarRightActions")',
        "self._input_area_vm.sync_selection(",
    ):
        assert token in main
    assert "_TWO_SIDE_PANEL_MIN_WIDTH" not in main
    for token in ("_PYRAMID_INDENTS", "_type_prompt_rows", "selection_detail"):
        assert token in input_area
    assert 'setObjectName("infoHistoryEmptyState")' in info_area
    assert "_BATCH_CATEGORY_PIVOT_NARROW_THRESHOLD = 380" in batch_list
    assert "def _initialize_application_theme(" in app
    assert "_initialize_application_theme(app, controller)" in bundle_entry
