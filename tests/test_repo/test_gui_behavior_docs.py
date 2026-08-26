"""Current GUI behavior documentation guards."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.unit

ROOT = Path(__file__).resolve().parents[2]


def test_gui_behavior_spec_covers_current_interaction_contracts() -> None:
    text = (ROOT / "docs" / "specs" / "gui-behavior.md").read_text(encoding="utf-8")

    for token in (
        "Single and batch input modes",
        "Cancellation",
        "persisted, draft and preview layers",
        "Keyboard shortcuts",
        "IPC",
        "docs/assets/screenshots/",
    ):
        assert token in text


def test_gui_capabilities_point_to_current_automated_owners() -> None:
    capabilities = (ROOT / "docs" / "capabilities.md").read_text(encoding="utf-8")
    shortcut = next(line for line in capabilities.splitlines() if line.startswith("| FEAT-GUI-008 |"))
    batch = next(line for line in capabilities.splitlines() if line.startswith("| FEAT-GUI-017 |"))

    assert "main_window.py" in shortcut
    assert "test_main_window_features.py" in shortcut
    assert shortcut.endswith("| verified |")
    assert "batch_list" in batch
    assert "conversion_panel" in batch
    assert "main-window tests" in batch
    assert batch.endswith("| tested |")


def test_gui_behavior_owners_and_regression_entrypoints_exist() -> None:
    paths = (
        "packages/apps/gui/src/docwen_gui/main_window.py",
        "packages/apps/gui/src/docwen_gui/widgets/batch_list.py",
        "packages/apps/gui/src/docwen_gui/widgets/conversion_panel.py",
        "packages/apps/gui/src/docwen_gui/widgets/action_area.py",
        "packages/apps/gui/tests/test_main_window_features_*.py",
        "packages/apps/gui/tests/test_gui_e2e_conversion_*.py",
    )
    for relative_path in paths:
        assert read_source_text(ROOT / relative_path)
