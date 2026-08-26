"""Tests for MainWindowViewModel projection integration.

Validates that the ViewModel holds selected-file source state and publishes a
``ui_projection`` derived from ``context → capabilities → projection`` as
documented in ``docs/architecture.md`` and ``docs/specs/gui-behavior.md``.

The VM is the single source of truth; it recomputes the projection whenever
files, mode, or the selected file change, and emits ``ui_projection_changed``.
No QApplication is required — the VM is exercised directly with a real
``Signal`` but no widgets attached.
"""

from __future__ import annotations

import pytest

from docwen_core.models.file_ref import FileRef
from docwen_gui.view_models.interaction import (
    MainWindowUiProjection,
    RightPanelSlot,
)
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.unit


# ── Fixtures ────────────────────────────────────────────────────────────


@pytest.fixture
def vm() -> MainWindowViewModel:
    return MainWindowViewModel(controller=None)


def _file_ref(path: str, category: str, fmt: str) -> FileRef:
    """Build the real Qt-free source-state model used by the ViewModel."""
    return FileRef(path=path, category=category, format=fmt)


# ── Initial projection (no file) ────────────────────────────────────────


class TestInitialProjection:
    def test_no_file_projection_hides_right_panel(self, vm: MainWindowViewModel) -> None:
        proj = vm.ui_projection
        assert isinstance(proj, MainWindowUiProjection)
        assert proj.right_panel_visible is False
        assert proj.right_panel_slot == RightPanelSlot.NONE

    def test_single_mode_hides_left_panel(self, vm: MainWindowViewModel) -> None:
        assert vm.ui_projection.left_panel_visible is False

    def test_batch_mode_shows_left_panel_before_files_exist(self, vm: MainWindowViewModel) -> None:
        vm.set_mode("batch")
        assert vm.ui_projection.left_panel_visible is True


# ── Selected-file source state ──────────────────────────────────────────


class TestSelectedFileState:
    def test_set_selected_file_updates_projection(self, vm: MainWindowViewModel) -> None:
        ref = _file_ref("/x.docx", "document", "docx")
        vm.set_selected_file(ref)
        proj = vm.ui_projection
        assert proj.right_panel_visible is True
        assert proj.right_panel_slot == RightPanelSlot.CONVERSION
        assert proj.conversion_context is not None
        assert proj.conversion_context.category == "document"

    def test_markdown_routes_to_template_slot(self, vm: MainWindowViewModel) -> None:
        ref = _file_ref("/note.md", "markdown", "md")
        vm.set_selected_file(ref)
        proj = vm.ui_projection
        assert proj.right_panel_slot == RightPanelSlot.TEMPLATE
        assert proj.template_context is not None

    def test_canonical_markdown_txt_routes_to_template_slot(self, vm: MainWindowViewModel) -> None:
        ref = _file_ref("/note.txt", "markdown", "txt")
        vm.set_selected_file(ref)
        proj = vm.ui_projection
        assert proj.right_panel_slot == RightPanelSlot.TEMPLATE
        assert proj.template_context is not None
        assert proj.template_context.file_path == "/note.txt"

    def test_unsupported_category_hides_right_panel(self, vm: MainWindowViewModel) -> None:
        ref = _file_ref("/p.pptx", "presentation", "pptx")
        vm.set_selected_file(ref)
        proj = vm.ui_projection
        assert proj.right_panel_visible is False
        assert proj.right_panel_slot == RightPanelSlot.NONE
        assert proj.center_action_visible is True

    def test_clear_selected_file_hides_right_panel(self, vm: MainWindowViewModel) -> None:
        ref = _file_ref("/x.docx", "document", "docx")
        vm.set_selected_file(ref)
        assert vm.ui_projection.right_panel_visible is True
        vm.clear_selected_file()
        proj = vm.ui_projection
        assert proj.right_panel_visible is False
        assert proj.right_panel_slot == RightPanelSlot.NONE


class TestFileListSignalContract:
    def test_files_changed_emits_real_file_refs(self, vm: MainWindowViewModel, tmp_path) -> None:
        received: list[list[FileRef]] = []
        vm.files_changed.connect(received.append)
        first = tmp_path / "x.docx"
        second = tmp_path / "y.xlsx"
        first.write_text("document-shaped filename with text content", encoding="utf-8")
        second.write_text("spreadsheet-shaped filename with text content", encoding="utf-8")

        vm.add_files([str(first), str(second)])

        assert len(received) == 1
        assert all(isinstance(ref, FileRef) for ref in received[0])
        assert received[0] == vm.files


# ── Mode-change recomputes projection ───────────────────────────────────


class TestModeRecompute:
    def test_mode_change_recomputes_left_visibility(self, vm: MainWindowViewModel) -> None:
        assert vm.ui_projection.left_panel_visible is False
        vm.set_mode("batch")
        assert vm.ui_projection.left_panel_visible is True
        vm.set_mode("single")
        assert vm.ui_projection.left_panel_visible is False


# ── Signal emission ─────────────────────────────────────────────────────


class TestProjectionSignal:
    def test_set_selected_file_emits_projection_changed(self, vm: MainWindowViewModel) -> None:
        received: list[MainWindowUiProjection] = []
        vm.ui_projection_changed.connect(received.append)
        ref = _file_ref("/x.docx", "document", "docx")
        vm.set_selected_file(ref)
        assert len(received) >= 1
        assert received[-1].right_panel_slot == RightPanelSlot.CONVERSION

    def test_mode_change_emits_projection_changed(self, vm: MainWindowViewModel) -> None:
        received: list[MainWindowUiProjection] = []
        vm.ui_projection_changed.connect(received.append)
        vm.set_mode("batch")
        assert any(p.left_panel_visible for p in received)

    def test_clear_selected_file_emits_projection_changed(self, vm: MainWindowViewModel) -> None:
        ref = _file_ref("/x.docx", "document", "docx")
        vm.set_selected_file(ref)
        received: list[MainWindowUiProjection] = []
        vm.ui_projection_changed.connect(received.append)
        vm.clear_selected_file()
        assert any(p.right_panel_visible is False for p in received)


# ── Deterministic selection advancement ─────────────────────────────────


class TestSelectionAdvancing:
    def test_removing_selected_file_clears_selection(self, vm: MainWindowViewModel) -> None:
        ref = _file_ref("/x.docx", "document", "docx")
        vm.set_selected_file(ref)
        assert vm.ui_projection.right_panel_visible is True
        # When the selected file is removed, selection must clear (no stale
        # projection pointing at a vanished file).
        vm.clear_selected_file()
        assert vm.selected_file is None
        assert vm.ui_projection.right_panel_visible is False
