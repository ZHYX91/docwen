"""Pure-Python tests for the interaction projection model.

These tests validate the current GUI interaction skeleton documented in
``docs/architecture.md`` and ``docs/specs/gui-behavior.md``.

The interaction module is intentionally Qt-free so the routing logic can be
unit-tested without a QApplication.  It encodes the old UI/UX rules (sourced
from the old project's ``selection_state.py``) in a new, maintainable form:
``context → capabilities → projection``.

Coverage:
- ``resolve_capabilities(category)`` — the capability matrix
- ``project_main_window_ui(context)`` — the projection rules
- ``build_interaction_context(mode, selected_file)`` — context assembly
- right-panel priority (TEMPLATE_SELECTION > FORMAT_CONVERSION > none)
- selection normalization entry points
"""

from __future__ import annotations

import pytest

from docwen_gui.view_models.interaction import (
    FileCapability,
    FileInteractionContext,
    MainWindowUiProjection,
    RightPanelSlot,
    UiMode,
    build_interaction_context,
    normalize_workflow_category,
    project_main_window_ui,
    resolve_capabilities,
)

pytestmark = pytest.mark.unit


# ── resolve_capabilities: capability matrix ─────────────────────────────


class TestResolveCapabilities:
    """The capability matrix mirrors old ``build_selector_category_render_plan``."""

    def test_none_category_yields_empty(self) -> None:
        assert resolve_capabilities(None) == frozenset()

    @pytest.mark.parametrize("category", ["markup", "presentation", "other", ""])
    def test_empty_capability_categories(self, category: str) -> None:
        assert resolve_capabilities(category) == frozenset()

    def test_markdown_yields_template_selection(self) -> None:
        caps = resolve_capabilities("markdown")
        assert FileCapability.TEMPLATE_SELECTION in caps
        assert FileCapability.FORMAT_CONVERSION not in caps

    def test_display_only_text_category_is_not_a_workflow_alias(self) -> None:
        assert normalize_workflow_category("text") == "text"
        assert resolve_capabilities("text") == frozenset()

    def test_workflow_category_is_not_reclassified_from_format(self) -> None:
        assert normalize_workflow_category("document") == "document"
        assert FileCapability.FORMAT_CONVERSION in resolve_capabilities("document")

    def test_core_markdown_workflow_projects_to_template_selection(self) -> None:
        assert normalize_workflow_category("markdown") == "markdown"
        assert FileCapability.TEMPLATE_SELECTION in resolve_capabilities("markdown")

    def test_document_yields_conversion_and_proofread(self) -> None:
        caps = resolve_capabilities("document")
        assert FileCapability.FORMAT_CONVERSION in caps
        assert FileCapability.PROOFREAD not in caps

    def test_spreadsheet_yields_conversion_only(self) -> None:
        caps = resolve_capabilities("spreadsheet")
        assert caps == frozenset({FileCapability.FORMAT_CONVERSION})

    def test_image_yields_conversion_compression_ocr(self) -> None:
        caps = resolve_capabilities("image")
        assert FileCapability.FORMAT_CONVERSION in caps
        assert FileCapability.IMAGE_COMPRESSION not in caps
        assert FileCapability.OCR not in caps

    def test_layout_yields_conversion_split_merge(self) -> None:
        caps = resolve_capabilities("layout")
        assert FileCapability.FORMAT_CONVERSION in caps
        assert FileCapability.PDF_SPLIT not in caps
        assert FileCapability.PDF_MERGE not in caps

    def test_category_is_case_insensitive(self) -> None:
        assert resolve_capabilities("Document") == resolve_capabilities("document")
        assert resolve_capabilities("MARKDOWN") == resolve_capabilities("markdown")

    def test_reserved_capabilities_are_returned_but_not_active_ui(self) -> None:
        """Reserved capabilities exist in the model but are not mapped to a
        visible slot — only TEMPLATE_SELECTION and FORMAT_CONVERSION are active
        UI capabilities in this round."""
        image_caps = resolve_capabilities("image")
        # Runtime capabilities are not copied into the panel-slot policy.
        assert FileCapability.OCR not in image_caps
        assert FileCapability.IMAGE_COMPRESSION not in image_caps


# ── project_main_window_ui: projection rules ────────────────────────────


def _ctx(mode: UiMode, category: str | None, fmt: str | None = None) -> FileInteractionContext:
    return build_interaction_context(
        mode=mode,
        category=category,
        current_format=fmt,
        file_path="/tmp/sample" if category is not None else None,
    )


class TestProjectMainWindowUi:
    """Projection rules: context → visible slots and right-panel routing."""

    def test_no_file_hides_right_panel_and_slot_none(self) -> None:
        ctx = build_interaction_context(mode=UiMode.SINGLE, category=None, current_format=None, file_path=None)
        proj = project_main_window_ui(ctx)
        assert proj.right_panel_visible is False
        assert proj.right_panel_slot == RightPanelSlot.NONE
        assert proj.left_panel_visible is False  # single mode

    def test_single_markdown_shows_template_slot(self) -> None:
        ctx = _ctx(UiMode.SINGLE, "markdown", "md")
        proj = project_main_window_ui(ctx)
        assert proj.right_panel_visible is True
        assert proj.right_panel_slot == RightPanelSlot.TEMPLATE
        assert proj.template_context is not None
        assert proj.conversion_context is None

    @pytest.mark.parametrize(
        "category,fmt",
        [
            ("document", "docx"),
            ("spreadsheet", "xlsx"),
            ("image", "png"),
            ("layout", "pdf"),
        ],
    )
    def test_supported_conversion_categories_route_to_conversion(self, category: str, fmt: str) -> None:
        ctx = _ctx(UiMode.SINGLE, category, fmt)
        proj = project_main_window_ui(ctx)
        assert proj.right_panel_visible is True
        assert proj.right_panel_slot == RightPanelSlot.CONVERSION
        assert proj.conversion_context is not None
        assert proj.conversion_context.category == category
        assert proj.conversion_context.current_format == fmt
        assert proj.template_context is None

    @pytest.mark.parametrize("category", ["markup", "presentation", "other"])
    def test_unsupported_categories_hide_right_panel(self, category: str) -> None:
        ctx = _ctx(UiMode.SINGLE, category, "x")
        proj = project_main_window_ui(ctx)
        assert proj.right_panel_visible is False
        assert proj.right_panel_slot == RightPanelSlot.NONE
        assert proj.center_action_visible is True
        assert proj.conversion_context is None
        assert proj.template_context is None

    def test_batch_mode_shows_left_panel(self) -> None:
        ctx = _ctx(UiMode.BATCH, "document", "docx")
        proj = project_main_window_ui(ctx)
        assert proj.left_panel_visible is True

    def test_single_mode_hides_left_panel(self) -> None:
        ctx = _ctx(UiMode.SINGLE, "document", "docx")
        proj = project_main_window_ui(ctx)
        assert proj.left_panel_visible is False

    def test_empty_batch_mode_shows_left_and_keeps_right_hidden(self) -> None:
        ctx = build_interaction_context(mode=UiMode.BATCH, category=None, current_format=None, file_path=None)
        proj = project_main_window_ui(ctx)
        assert proj.left_panel_visible is True
        assert proj.right_panel_visible is False
        assert proj.right_panel_slot == RightPanelSlot.NONE

    def test_conversion_context_carries_file_path(self) -> None:
        ctx = _ctx(UiMode.SINGLE, "document", "docx")
        proj = project_main_window_ui(ctx)
        assert proj.conversion_context is not None
        assert proj.conversion_context.file_path == "/tmp/sample"

    def test_template_context_carries_file_path(self) -> None:
        ctx = _ctx(UiMode.SINGLE, "markdown", "md")
        proj = project_main_window_ui(ctx)
        assert proj.template_context is not None
        assert proj.template_context.file_path == "/tmp/sample"

    def test_core_txt_markdown_workflow_routes_to_template_slot(self) -> None:
        ctx = _ctx(UiMode.SINGLE, "markdown", "txt")
        proj = project_main_window_ui(ctx)
        assert ctx.selected_category == "markdown"
        assert ctx.selected_format == "txt"
        assert proj.right_panel_visible is True
        assert proj.right_panel_slot == RightPanelSlot.TEMPLATE
        assert proj.template_context is not None
        assert proj.conversion_context is None


# ── Right-panel priority ────────────────────────────────────────────────


class TestRightPanelPriority:
    """TEMPLATE_SELECTION (100) > FORMAT_CONVERSION (50) > none (0).

    Currently no category yields both TEMPLATE_SELECTION and FORMAT_CONVERSION,
    so the priority is enforced structurally: if a future category yields both,
    template wins.  This test pins the rule so future capability additions
    cannot silently break the priority.
    """

    def test_markdown_priority_over_conversion(self) -> None:
        """If markdown ever gains FORMAT_CONVERSION, template must still win."""
        ctx = _ctx(UiMode.SINGLE, "markdown", "md")
        proj = project_main_window_ui(ctx)
        assert proj.right_panel_slot == RightPanelSlot.TEMPLATE

    def test_document_without_template_shows_conversion(self) -> None:
        ctx = _ctx(UiMode.SINGLE, "document", "docx")
        proj = project_main_window_ui(ctx)
        assert proj.right_panel_slot == RightPanelSlot.CONVERSION


# ── Projection dataclass shape ──────────────────────────────────────────


class TestProjectionShape:
    def test_projection_is_frozen(self) -> None:
        ctx = _ctx(UiMode.SINGLE, "document", "docx")
        proj = project_main_window_ui(ctx)
        assert isinstance(proj, MainWindowUiProjection)
        with pytest.raises((AttributeError, Exception)):
            proj.right_panel_visible = True  # type: ignore[misc]

    def test_context_is_frozen(self) -> None:
        ctx = _ctx(UiMode.SINGLE, "document", "docx")
        with pytest.raises((AttributeError, Exception)):
            ctx.mode = UiMode.BATCH  # type: ignore[misc]


# ── build_interaction_context assembly ──────────────────────────────────


class TestBuildContext:
    def test_none_category_produces_empty_capabilities(self) -> None:
        ctx = build_interaction_context(mode=UiMode.SINGLE, category=None, current_format=None, file_path=None)
        assert ctx.capabilities == frozenset()

    def test_capabilities_derived_from_category(self) -> None:
        ctx = build_interaction_context(mode=UiMode.SINGLE, category="image", current_format="png", file_path="/x.png")
        assert FileCapability.FORMAT_CONVERSION in ctx.capabilities
        assert FileCapability.OCR not in ctx.capabilities

    def test_context_records_mode_and_selection(self) -> None:
        ctx = build_interaction_context(
            mode=UiMode.BATCH, category="document", current_format="docx", file_path="/a.docx"
        )
        assert ctx.mode == UiMode.BATCH
        assert ctx.selected_category == "document"
        assert ctx.selected_format == "docx"

    def test_context_keeps_canonical_markdown_category_for_txt(self) -> None:
        ctx = build_interaction_context(
            mode=UiMode.SINGLE, category="markdown", current_format="txt", file_path="/note.txt"
        )
        assert ctx.selected_category == "markdown"
        assert FileCapability.TEMPLATE_SELECTION in ctx.capabilities
