"""Long-term GUI interaction skeleton — pure Python, Qt-free.

This module encodes the old UI/UX routing rules (sourced from the old
project's ``selection_state.py``) in a new, maintainable form:

::

    FileRef / format category
            ↓
    FileInteractionContext
            ↓
    CapabilityResolver   (resolve_capabilities)
            ↓
    MainWindowUiProjection   (project_main_window_ui)
            ↓
    MainWindow render bindings
            ↓
    Widgets

Design principles (see ``docs/architecture.md`` and ``docs/specs/gui-behavior.md``):

- The old ``category route`` is a *specification source*, not the new core
  model.  The new core is ``context + capabilities + projection``.
- ``right_panel_slot`` is a UI-projection field, never core business state.
- State-derivation logic is pure Python and testable without Qt.
- Only capabilities explicitly mapped to a visible slot may appear in the UI.
- ``markup`` / ``presentation`` / ``other`` never map to the conversion panel.

This module imports nothing from PySide6 so it can be unit-tested in
isolation.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum

# ── Enums ───────────────────────────────────────────────────────────────


class UiMode(StrEnum):
    """Top-level input mode for the main window."""

    SINGLE = "single"
    BATCH = "batch"


class FileCapability(StrEnum):
    """What the currently-selected file can do in the GUI.

    Split into *active UI capabilities* (mapped to a visible slot this round)
    and *reserved capabilities* (allowed in the model but not exposed as
    clickable UI yet).  Only ``TEMPLATE_SELECTION`` and ``FORMAT_CONVERSION``
    are active UI capabilities; the rest are long-term reservations.
    """

    # ── Active UI capabilities ──────────────────────────────────────────
    TEMPLATE_SELECTION = "template_selection"
    FORMAT_CONVERSION = "format_conversion"

    # ── Reserved capabilities (present in model, not UI this round) ─────
    IMAGE_COMPRESSION = "image_compression"
    PDF_SPLIT = "pdf_split"
    PDF_MERGE = "pdf_merge"
    OCR = "ocr"
    PROOFREAD = "proofread"


class RightPanelSlot(StrEnum):
    """Which panel the right slot currently shows."""

    NONE = "none"
    TEMPLATE = "template"
    CONVERSION = "conversion"


# ── Panel-slot policy ───────────────────────────────────────────────────
#
# Project capability categories into stable panel slots instead of coupling
# presentation to route strings:
#
#   text sources      → TEMPLATE_SELECTION → template_action
#   document sources  → FORMAT_CONVERSION  → conversion_action
#   old route=spreadsheet   → FORMAT_CONVERSION            → conversion_action
#   old route=image         → FORMAT_CONVERSION + reserved → conversion_action
#   old route=layout        → FORMAT_CONVERSION + reserved → conversion_action
#   old route=other         → (empty)                      → action_only (none)
#
# This table decides only which panel shell owns the right slot. Runtime owns
# business capabilities and route/action availability; those facts must never
# be copied into this presentation policy.

_PANEL_SLOT_CAPABILITY_MATRIX: dict[str, frozenset[FileCapability]] = {
    "markdown": frozenset({FileCapability.TEMPLATE_SELECTION}),
    "document": frozenset({FileCapability.FORMAT_CONVERSION}),
    "spreadsheet": frozenset({FileCapability.FORMAT_CONVERSION}),
    "image": frozenset({FileCapability.FORMAT_CONVERSION}),
    "layout": frozenset({FileCapability.FORMAT_CONVERSION}),
    # markup / presentation / other carry no active UI capability this round.
    "markup": frozenset(),
    "presentation": frozenset(),
    "other": frozenset(),
}

# Right-panel priority — a file may carry several capabilities; the slot is
# chosen by the highest-priority *active UI* capability present.  Higher
# number wins.  Reserved capabilities never own a slot.
_SLOT_PRIORITY: dict[FileCapability, int] = {
    FileCapability.TEMPLATE_SELECTION: 100,
    FileCapability.FORMAT_CONVERSION: 50,
}


def normalize_workflow_category(category: str | None) -> str | None:
    """Normalize a Core-owned workflow category for GUI projection.

    Format detection and workflow classification belong to Core.  The GUI
    therefore never reclassifies a file from its format or suffix.  Display
    groups such as the left panel's ``text`` bucket do not belong at this
    Core-owned workflow-category boundary.
    """
    if category is None:
        return None
    normalized = str(category).strip().lower()
    if not normalized:
        return None
    return normalized


def resolve_capabilities(category: str | None) -> frozenset[FileCapability]:
    """Return the capability set for *category*.

    The canonical category vocabulary comes from
    ``docwen_core.formats.categories``: ``document``, ``spreadsheet``,
    ``image``, ``layout``, ``markdown``, ``markup``, ``presentation``,
    ``other``.  Unknown, display-only, or ``None`` categories yield an empty
    set instead of being silently rewritten into another workflow.

    """
    normalized = normalize_workflow_category(category)
    if normalized is None:
        return frozenset()
    return _PANEL_SLOT_CAPABILITY_MATRIX.get(normalized, frozenset())


# ── Dataclasses ─────────────────────────────────────────────────────────


@dataclass(frozen=True)
class ConversionContext:
    """Context handed to the conversion panel when the right slot is conversion."""

    category: str
    current_format: str | None
    file_path: str


@dataclass(frozen=True)
class TemplateContext:
    """Context handed to the template selector when the right slot is template."""

    file_path: str


@dataclass(frozen=True)
class FileInteractionContext:
    """Core input to the projection — the current interaction snapshot.

    ``selected_file`` is typed loosely here (``object | None``) so this module
    stays Qt-free; at the call site it holds a ``FileRef | None``.  The
    projector reads ``selected_category``/``selected_format``/``mode`` and
    ``selected_file_path`` only — it never inspects ``selected_file``'s shape.
    """

    mode: UiMode
    selected_file: object | None
    selected_category: str | None
    selected_format: str | None
    selected_file_path: str | None
    capabilities: frozenset[FileCapability]


@dataclass(frozen=True)
class MainWindowUiProjection:
    """Render-only output — what the MainWindow should show right now.

    The MainWindow binds this to widget visibility and the right-panel stack;
    it must not re-derive business state from file suffixes or categories.
    """

    left_panel_visible: bool
    right_panel_visible: bool
    right_panel_slot: RightPanelSlot
    center_action_visible: bool
    info_area_visible: bool
    conversion_context: ConversionContext | None
    template_context: TemplateContext | None


# ── Context assembly ────────────────────────────────────────────────────


def build_interaction_context(
    *,
    mode: UiMode,
    category: str | None,
    current_format: str | None,
    file_path: str | None,
    selected_file: object | None = None,
) -> FileInteractionContext:
    """Assemble a :class:`FileInteractionContext` from raw selection inputs.

    When ``file_path``/``category`` are ``None`` the context represents the
    no-file state (empty capabilities).  ``selected_file`` is optional; callers
    that have a real ``FileRef`` may pass it through so the VM can retain the
    source reference, but the projector only reads category/format/mode.
    """
    normalized_category = normalize_workflow_category(category)
    capabilities = resolve_capabilities(normalized_category)
    return FileInteractionContext(
        mode=mode,
        selected_file=selected_file,
        selected_category=normalized_category,
        selected_format=current_format,
        selected_file_path=file_path,
        capabilities=capabilities,
    )


# ── Projection ──────────────────────────────────────────────────────────


def _pick_right_slot(caps: frozenset[FileCapability]) -> RightPanelSlot:
    """Choose the right-panel slot by active-UI capability priority."""
    best: FileCapability | None = None
    best_priority = -1
    for cap in caps:
        priority = _SLOT_PRIORITY.get(cap)
        if priority is None:
            continue  # reserved capability — never owns a slot
        if priority > best_priority:
            best = cap
            best_priority = priority
    if best is None:
        return RightPanelSlot.NONE
    if best is FileCapability.TEMPLATE_SELECTION:
        return RightPanelSlot.TEMPLATE
    if best is FileCapability.FORMAT_CONVERSION:
        return RightPanelSlot.CONVERSION
    return RightPanelSlot.NONE


def project_main_window_ui(context: FileInteractionContext) -> MainWindowUiProjection:
    """Project an interaction context into render-only UI state.

    Rules:
    - ``left_panel_visible`` is true only in batch mode.
    - No selected file → right panel hidden, slot ``none``.
    - ``TEMPLATE_SELECTION`` → right visible, slot ``template``.
    - ``FORMAT_CONVERSION`` → right visible, slot ``conversion``.
    - No active UI capability → right hidden, slot ``none``.
    - ``info_area`` and the action area stay visible (the action area carries
      the no-files / "other" semantics from the old ``action_only`` mode).
    """
    has_file = context.selected_category is not None and context.selected_file is not None
    # A context built via build_interaction_context with a non-None file_path
    # but no FileRef still counts as "has a file" for routing purposes — the
    # projector keys off category presence, not the FileRef object.
    if not has_file and context.selected_category is not None:
        has_file = True

    left_visible = context.mode == UiMode.BATCH

    if not has_file:
        return MainWindowUiProjection(
            left_panel_visible=left_visible,
            right_panel_visible=False,
            right_panel_slot=RightPanelSlot.NONE,
            center_action_visible=False,
            info_area_visible=True,
            conversion_context=None,
            template_context=None,
        )

    slot = _pick_right_slot(context.capabilities)
    file_path = _extract_file_path(context)

    if slot == RightPanelSlot.TEMPLATE:
        return MainWindowUiProjection(
            left_panel_visible=left_visible,
            right_panel_visible=True,
            right_panel_slot=RightPanelSlot.TEMPLATE,
            center_action_visible=True,
            info_area_visible=True,
            conversion_context=None,
            template_context=TemplateContext(file_path=file_path) if file_path else None,
        )

    if slot == RightPanelSlot.CONVERSION:
        category = context.selected_category or "other"
        return MainWindowUiProjection(
            left_panel_visible=left_visible,
            right_panel_visible=True,
            right_panel_slot=RightPanelSlot.CONVERSION,
            center_action_visible=True,
            info_area_visible=True,
            conversion_context=ConversionContext(
                category=category,
                current_format=context.selected_format,
                file_path=file_path,
            ),
            template_context=None,
        )

    # No active UI capability — right hidden, action area still visible
    # (mirrors old ``action_only`` mode for "other" files).
    return MainWindowUiProjection(
        left_panel_visible=left_visible,
        right_panel_visible=False,
        right_panel_slot=RightPanelSlot.NONE,
        center_action_visible=True,
        info_area_visible=True,
        conversion_context=None,
        template_context=None,
    )


def _extract_file_path(context: FileInteractionContext) -> str:
    """Best-effort file-path extraction for panel contexts.

    Prefers the explicit ``selected_file_path`` on the context (the routing
    input); falls back to reading ``path`` off the selected file object if it
    looks like a ``FileRef``.  Returns ``""`` when no path is available —
    panels treat an empty path as a reset/no-file state.
    """
    if context.selected_file_path:
        return context.selected_file_path
    candidate = getattr(context.selected_file, "path", None)
    if isinstance(candidate, str) and candidate:
        return candidate
    return ""


__all__ = [
    "ConversionContext",
    "FileCapability",
    "FileInteractionContext",
    "MainWindowUiProjection",
    "RightPanelSlot",
    "TemplateContext",
    "UiMode",
    "build_interaction_context",
    "normalize_workflow_category",
    "project_main_window_ui",
    "resolve_capabilities",
]
