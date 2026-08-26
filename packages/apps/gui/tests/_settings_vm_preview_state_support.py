"""Tests for SettingsViewModel persisted/draft/preview three-layer state machine.

Validates the persisted/draft/preview behavior documented in
``docs/specs/gui-behavior.md``:

- ``persisted`` — config on disk (source of truth)
- ``draft`` — in-dialog edits (not yet persisted)
- ``preview`` — temporary visual effects (theme/opacity applied to UI
  without persisting; Cancel restores to persisted baseline)

Key invariants:
1. Opening settings captures the persisted baseline.
2. Editing changes draft but does NOT persist.
3. Theme/opacity changes preview immediately without writing config.
4. Apply writes draft → persisted and updates baseline.
5. Cancel restores preview → persisted baseline.
6. OK = Apply then close.
7. After Apply, then further edit, then Cancel → restores to post-Apply state.
8. Settings changes after Apply should trigger preview for the applied state.
9. Close without Apply does NOT modify persisted config.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen_application.controller import ApplicationController
from docwen_bundle.config_port import ConfigPortAdapter
from docwen_gui.view_models.settings_vm import SECTION_GUI, SECTION_OUTPUT, SettingsViewModel

pytestmark = pytest.mark.unit

PROJECT_CONFIGS = Path(__file__).resolve().parent.parent.parent.parent.parent / "configs"


@pytest.fixture
def config_port(tmp_path: Path) -> ConfigPortAdapter:
    return ConfigPortAdapter(base_dir=PROJECT_CONFIGS, user_dir=tmp_path / "configs")


@pytest.fixture
def vm(config_port: ConfigPortAdapter) -> SettingsViewModel:
    controller = ApplicationController(config_port=config_port)
    return SettingsViewModel(controller=controller)


__all__ = (
    "SECTION_GUI",
    "SECTION_OUTPUT",
    "ApplicationController",
    "ConfigPortAdapter",
    "SettingsViewModel",
    "config_port",
    "pytest",
    "pytestmark",
    "vm",
)
