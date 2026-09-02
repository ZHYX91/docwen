"""Evidence guards for the VIS-2026-07-27-398 reachable UI candidate."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]
CARD_NAME = "reachable-desktop-ui-polish-stage-card-2026-07-27.md"
LOCALES = (
    "de_DE",
    "en_US",
    "es_ES",
    "fr_FR",
    "ja_JP",
    "ko_KR",
    "pt_BR",
    "ru_RU",
    "vi_VN",
    "zh_CN",
    "zh_TW",
)


def _read(relative_path: str) -> str:
    return (ROOT / relative_path).read_text(encoding="utf-8")


def test_reachable_ui_candidate_keeps_geometry_and_surface_contracts() -> None:
    main = _read("packages/apps/gui/src/docwen_gui/main_window.py")
    policy = _read("packages/apps/gui/src/docwen_gui/window_behavior.py")
    interaction = _read("packages/apps/gui/src/docwen_gui/view_models/interaction.py")
    action = _read("packages/apps/gui/src/docwen_gui/widgets/action_area.py")
    conversion = _read("packages/apps/gui/src/docwen_gui/widgets/conversion_panel.py")
    settings = _read("packages/apps/gui/src/docwen_gui/widgets/settings/base_tab.py")
    batch_dialogs = _read("packages/apps/gui/src/docwen_gui/widgets/batch_dialogs.py")

    for token in (
        "def _normal_panel_transition_rect(",
        "def _context_panel_width_contribution(",
        "main_window.ipc_file_received",
    ):
        assert token in main
    for token in (
        "class _ConfigGetter(Protocol):",
        "center_panel_width: int = 460",
        "left_panel_width: int = 400",
        "right_panel_width: int = 300",
    ):
        assert token in policy

    # Entering batch mode is itself left-panel demand, including an empty list.
    assert "left_visible = context.mode == UiMode.BATCH" in interaction
    assert 'setObjectName("actionPanelTitle")' in action
    assert 'setObjectName("conversionPanelScrollArea")' in conversion
    assert "def _on_vm_state_changed(" in conversion
    assert 'setObjectName("settingsTabTitle")' in settings
    assert "main_window.batch_add_failed_title" in batch_dialogs
    assert "warn(title, message, details=details, parent=parent)" in batch_dialogs


def test_reachable_ui_candidate_localizes_new_owned_feedback() -> None:
    keys = (
        "ipc_file_received",
        "batch_add_failed_title",
        "batch_add_failed_message",
        "batch_add_failed_reason",
    )
    for locale in LOCALES:
        text = _read(f"i18n/locales/{locale}.toml")
        for key in keys:
            assert f"{key} = " in text, f"{locale} is missing {key}"
