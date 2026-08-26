"""Tests for MainWindow semantic slot layout.

Validates the three-slot layout documented in ``docs/specs/gui-behavior.md``:

- left: batch list (batch mode visible)
- center: input area + action area + info area + bottom bar
- right: QStackedWidget with template selector + conversion panel
- No QSplitter
- conversion_panel lives in right stack, not center column
- info_area lives in center column, not a separate right column
"""

from __future__ import annotations

import pytest
from PySide6.QtWidgets import QGridLayout, QStackedWidget, QWidget

from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.gui


# ── Fixtures ────────────────────────────────────────────────────────────


@pytest.fixture
def vm() -> MainWindowViewModel:
    return MainWindowViewModel(controller=None)


@pytest.fixture
def window(vm: MainWindowViewModel, qapp):
    from docwen_gui.main_window import MainWindow

    w = MainWindow(view_model=vm)
    w.setup_ui()
    yield w
    w.close()


# ── Helpers ─────────────────────────────────────────────────────────────


def _find_by_name(widget, obj_name: str) -> QWidget | None:
    return widget.findChild(QWidget, obj_name)


# ── Layout structure ────────────────────────────────────────────────────


class TestLayoutStructure:
    def test_root_uses_grid_layout(self, window) -> None:
        """Root layout of the central container must be a grid (no splitter)."""
        central = _find_by_name(window, "centralContainer")
        assert central is not None, "centralContainer not found"
        root_layout = central.layout()
        assert isinstance(root_layout, QGridLayout), f"Expected QGridLayout, got {type(root_layout).__name__}"

    def test_right_panel_frame_exists(self, window) -> None:
        """Right panel must have a dedicated frame with a QStackedWidget."""
        right_frame = _find_by_name(window, "rightPanelFrame")
        assert right_frame is not None, "rightPanelFrame not found"

    def test_right_panel_has_stacked_widget(self, window) -> None:
        """Right panel must contain a QStackedWidget for template/conversion."""
        right_frame = _find_by_name(window, "rightPanelFrame")
        assert right_frame is not None
        stack = right_frame.findChild(QStackedWidget)
        assert stack is not None, "No QStackedWidget in right panel"
        assert stack.count() >= 2, f"Stack should have ≥2 widgets, has {stack.count()}"

    def test_conversion_panel_in_right_stack_not_center(self, window) -> None:
        """Conversion panel must live in the right stack, not the center column."""
        right_frame = _find_by_name(window, "rightPanelFrame")
        assert right_frame is not None
        cp = window.conversion_panel
        # Walk ancestors from conversion_panel upward
        parent = cp.parentWidget()
        found_right = False
        while parent is not window and parent is not None:
            if parent is right_frame:
                found_right = True
                break
            parent = parent.parentWidget()
        assert found_right, "conversion_panel should be inside right panel, not center column"

    def test_info_area_in_center_column(self, window) -> None:
        """Info area must live in the center column, not a separate grid column."""
        center_col = _find_by_name(window, "mainWindowCenterColumn")
        assert center_col is not None, "mainWindowCenterColumn not found"
        ia = window.info_area
        parent = ia.parentWidget()
        found_center = False
        while parent is not window and parent is not None:
            if parent is center_col:
                found_center = True
                break
            parent = parent.parentWidget()
        assert found_center, "info_area should be a descendant of mainWindowCenterColumn"

    def test_grid_no_widget_at_old_info_area_column(self, window) -> None:
        """The old column 3 position must no longer hold a widget."""
        central = _find_by_name(window, "centralContainer")
        assert central is not None
        grid = central.layout()
        assert isinstance(grid, QGridLayout)
        old_ia_item = grid.itemAtPosition(0, 3)
        # Either no item at all, or a spacer (not a widget)
        if old_ia_item is not None:
            assert old_ia_item.widget() is None, "Column 3 should not hold the old info_area widget"


# ── Public API preserved ────────────────────────────────────────────────


class TestPublicApiPreserved:
    """All public properties used by existing tests must survive the refactor."""

    def test_input_area_accessible(self, window) -> None:
        assert window.input_area is not None

    def test_batch_list_accessible(self, window) -> None:
        assert window.batch_list is not None

    def test_conversion_panel_accessible(self, window) -> None:
        assert window.conversion_panel is not None

    def test_action_area_accessible(self, window) -> None:
        assert window.action_area is not None

    def test_info_area_accessible(self, window) -> None:
        assert window.info_area is not None

    def test_settings_button_accessible(self, window) -> None:
        assert window.settings_button is not None


# ── No splitter ─────────────────────────────────────────────────────────


class TestNoSplitter:
    def test_no_qsplitter_in_window(self, window) -> None:
        from PySide6.QtWidgets import QSplitter

        splitters = window.findChildren(QSplitter)
        assert len(splitters) == 0, "MainWindow must not use QSplitter"
