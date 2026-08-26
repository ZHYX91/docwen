"""Model-state tests for ConversionPanelViewModel.

These tests validate ViewModel state transitions, signals, and business logic.
No QApplication required.
"""

import pytest
from tests.support.gui_vm_fakes import FakeMainWindowViewModel

from docwen_gui.view_models.conversion_panel_vm import (
    BUTTON_COLORS,
    COMPRESSIBLE_FORMATS,
    SENSITIVE_WORD,
    SYMBOL_CORRECTION,
    SYMBOL_PAIRING,
    TYPOS_RULE,
    VALIDATION_OPTION_KEYS,
    ConversionPanelViewModel,
)

pytestmark = pytest.mark.unit


@pytest.fixture
def vm() -> ConversionPanelViewModel:
    return ConversionPanelViewModel(FakeMainWindowViewModel())  # type: ignore[arg-type]


__all__ = (
    "BUTTON_COLORS",
    "COMPRESSIBLE_FORMATS",
    "SENSITIVE_WORD",
    "SYMBOL_CORRECTION",
    "SYMBOL_PAIRING",
    "TYPOS_RULE",
    "VALIDATION_OPTION_KEYS",
    "ConversionPanelViewModel",
    "FakeMainWindowViewModel",
    "pytest",
    "pytestmark",
    "vm",
)
