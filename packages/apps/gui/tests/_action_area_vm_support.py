"""Model-state tests for ActionAreaViewModel.

These tests validate ViewModel state transitions, 7 setup modes,
signal emission, and option collection. No QApplication required.
"""

import pytest
from tests.support.gui_vm_fakes import FakeMainWindowViewModel

from docwen_gui.i18n import t as _t
from docwen_gui.view_models.action_area_vm import (
    DEFAULT_PROOFREAD_OPTIONS,
    MODE_DOCUMENT,
    MODE_IMAGE,
    MODE_LAYOUT,
    MODE_MD_TO_DOCUMENT,
    MODE_MD_TO_SPREADSHEET,
    MODE_SPREADSHEET,
    PROOFREAD_OPTION_KEYS,
    SENSITIVE_WORD,
    SYMBOL_CORRECTION,
    SYMBOL_PAIRING,
    TYPOS_RULE,
    ActionAreaViewModel,
)

pytestmark = pytest.mark.unit


def _write_minimal_base_config_tree(base_dir) -> None:
    """Seeds a minimal base config tree so ConfigLoader does not crash.

    ConfigLoader expects every CONFIG_FILES entry to exist before the
    three-layer merge can proceed.  An empty tmp_path won't have them.
    """
    from pathlib import Path as _P

    from docwen_runtime.config.registry import CONFIG_FILES

    bd = _P(base_dir)
    for spec in CONFIG_FILES:
        path = bd / spec.rel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text("\n", encoding="utf-8")


@pytest.fixture
def vm() -> ActionAreaViewModel:
    return ActionAreaViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]


def FakeMainWindowViewModel_for_port(config_port) -> FakeMainWindowViewModel:
    """Build a ``FakeMainWindowViewModel`` whose controller.config_port is a real adapter.

    The existing ``FakeMainWindowViewModel`` double takes a values dict and builds a
    ``FakeConfigView`` that only knows the keys in that dict. For the
    sync tests we need a config port that actually reads the persisted
    config dir, so wire the real adapter through the same shape.
    """
    main_vm = FakeMainWindowViewModel({})
    main_vm.controller.config_port = config_port
    return main_vm


__all__ = (
    "DEFAULT_PROOFREAD_OPTIONS",
    "MODE_DOCUMENT",
    "MODE_IMAGE",
    "MODE_LAYOUT",
    "MODE_MD_TO_DOCUMENT",
    "MODE_MD_TO_SPREADSHEET",
    "MODE_SPREADSHEET",
    "PROOFREAD_OPTION_KEYS",
    "SENSITIVE_WORD",
    "SYMBOL_CORRECTION",
    "SYMBOL_PAIRING",
    "TYPOS_RULE",
    "ActionAreaViewModel",
    "FakeMainWindowViewModel",
    "FakeMainWindowViewModel_for_port",
    "_t",
    "_write_minimal_base_config_tree",
    "pytest",
    "pytestmark",
    "vm",
)
