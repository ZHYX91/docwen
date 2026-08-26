"""ActionAreaViewModel — observable state for the ActionArea widget.

This is the single source of truth for the ActionArea's observable state.
Widgets bind to its signals and properties; user actions flow through method
calls that delegate to ``MainWindowViewModel``.

Widgets never call runtime/plugins directly — they go through this ViewModel
which delegates to the parent ``MainWindowViewModel``.

Supports 7 ``setup_for_*`` modes:
  1. ``setup_for_document_file`` — Document → MD
  2. ``setup_for_spreadsheet_file`` — Spreadsheet → MD
  3. ``setup_for_image_file`` — Image → MD (OCR enabled by default)
  4. ``setup_for_layout_file`` — Layout → MD
  5. ``setup_for_other_file`` — Other → MD
  6. ``setup_for_md_to_document`` — MD → Document
  7. ``setup_for_md_to_spreadsheet`` — MD → Spreadsheet
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QObject, Signal

from docwen_cli.options.to_markdown import build_to_markdown_options
from docwen_core.formats.categories import CATEGORY_DOCUMENT, CATEGORY_SPREADSHEET, get_category
from docwen_gui.i18n import t as _t

from ._optimization_filter import (
    OptimizationChoice,
    OptimizationChoicesResult,
    OptimizationSource,
    discover_optimization_choices,
)
from ._runtime_route_filter import (
    RuntimeRouteChoicesResult,
    RuntimeRouteSource,
    discover_runtime_route_choices,
)

if TYPE_CHECKING:
    from .main_window_vm import MainWindowViewModel

logger = logging.getLogger(__name__)

# ── Proofread option keys ────────────────────────────────────────────────

SYMBOL_PAIRING = "symbol_pairing"
SYMBOL_CORRECTION = "symbol_correction"
TYPOS_RULE = "typos_rule"
SENSITIVE_WORD = "sensitive_word"

PROOFREAD_OPTION_KEYS = [SYMBOL_PAIRING, SYMBOL_CORRECTION, TYPOS_RULE, SENSITIVE_WORD]

# Canonical proofread defaults.
DEFAULT_PROOFREAD_OPTIONS: dict[str, bool] = {
    SYMBOL_PAIRING: True,
    SYMBOL_CORRECTION: True,
    TYPOS_RULE: True,
    SENSITIVE_WORD: False,
}

# ── Mode constants ───────────────────────────────────────────────────────

MODE_DOCUMENT = "document"
MODE_SPREADSHEET = "spreadsheet"
MODE_IMAGE = "image"
MODE_LAYOUT = "layout"
MODE_OTHER = "other"
MODE_MD_TO_DOCUMENT = "docx"  # file_type for md_to_document
MODE_MD_TO_SPREADSHEET = "md_to_spreadsheet"
MODE_AGGREGATE = "aggregate"  # merge/aggregate mode (many-to-one)

# ── Default options per file type ────────────────────────────────────────

_DEFAULT_FILE_TO_MD: dict[str, Any] = {
    "extract_image": True,
    "extract_ocr": False,
    "remove_numbering": True,
    "add_numbering": False,
    "numbering_scheme": "gongwen_standard",
    "action_name": None,
}

_NON_FILE_TO_MD_TYPES = {MODE_MD_TO_DOCUMENT, MODE_MD_TO_SPREADSHEET, MODE_AGGREGATE}
_PRESENTATION_FILE_TO_MD_TYPES = {"ppt", "pptx"}
_MARKUP_FILE_TO_MD_TYPES = {"html", "htm", "mhtml", "mht", "enex", "epub"}


class ActionAreaViewModel(QObject):
    """Observable state for the ActionArea widget.

    Holds state for the current setup mode (one of 7), including
    button labels, options, and the cancel stack.

    Signals:
        state_changed: emitted when the mode or visible state changes.
        conversion_requested: emitted when user clicks the action button.
        cancel_requested: emitted when user clicks cancel.
    """

    # ── Signals ──────────────────────────────────────────────────────────

    state_changed = Signal()
    """Emitted when file_type, visibility, or options change — widgets rebind."""

    conversion_requested = Signal(str, str, object)
    """Emitted when user clicks the primary action button:
    (target_format, file_path, options_dict)."""

    cancel_requested = Signal()
    """Emitted when user clicks the Cancel button."""

    # ── Construction ──────────────────────────────────────────────────────

    def __init__(
        self,
        main_vm: MainWindowViewModel | None = None,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._main_vm = main_vm

        # Core state
        self._visible: bool = False
        self._cancel_visible: bool = False
        self._file_type: str | None = None
        self._file_path: str | None = None
        self._mode: str = "single"

        # File→MD options
        self._extract_image: bool = True
        self._extract_ocr: bool = False
        self._optimize_for_type: str | None = None
        self._optimization_choices_result = OptimizationChoicesResult(status="ready", choices=())
        self._optimization_sources: tuple[OptimizationSource, ...] = ()
        # File→MD (docx→md) numbering options
        self._doc_remove_numbering: bool = True
        self._doc_add_numbering: bool = False
        self._doc_numbering_scheme: str = "gongwen_standard"

        # MD→Document (md→docx) numbering options
        self._md_remove_numbering: bool = True
        self._md_add_numbering: bool = False
        self._md_numbering_scheme: str = "hierarchical_standard"
        self._show_numbering: bool = True
        self._show_optimize: bool = False

        # MD→Document options
        self._target_format: str = "docx"
        self._available_target_formats: list[str] = []
        self._target_route_choices_result = RuntimeRouteChoicesResult(status="empty", choices=())
        self._last_document_format: str = "docx"
        self._last_spreadsheet_format: str = "xlsx"

        # Proofread options
        self._proofread_options: dict[str, bool] = dict(DEFAULT_PROOFREAD_OPTIONS)
        self._show_proofread: bool = False

        # Aggregate actions and optimization actions are mutually exclusive.
        # Keep both states explicit from construction so every setup transition
        # can clear the inactive branch without getattr-based shadow state.
        self._aggregate_action_name: str = ""
        self._aggregate_file_list: list[str] = []

        logger.info("ActionAreaViewModel initialized")

    # ── Core properties ──────────────────────────────────────────────────

    @property
    def visible(self) -> bool:
        """Whether the ActionArea is visible."""
        return self._visible

    @property
    def cancel_visible(self) -> bool:
        """Whether the cancel stack page is shown."""
        return self._cancel_visible

    @property
    def file_type(self) -> str | None:
        """Current file type / mode identifier."""
        return self._file_type

    @property
    def file_path(self) -> str | None:
        """Current file path."""
        return self._file_path

    @property
    def mode(self) -> str:
        """Current UI mode: ``"single"`` or ``"batch"``."""
        return self._mode

    # ── File→MD option properties ────────────────────────────────────────

    @property
    def extract_image(self) -> bool:
        """Whether to extract images during conversion."""
        return self._extract_image

    @extract_image.setter
    def extract_image(self, value: bool) -> None:
        if self._extract_image != value:
            self._extract_image = value
            self.state_changed.emit()

    @property
    def extract_ocr(self) -> bool:
        """Whether to enable OCR during conversion."""
        return self._extract_ocr

    @extract_ocr.setter
    def extract_ocr(self, value: bool) -> None:
        if self._extract_ocr != value:
            self._extract_ocr = value
            self.state_changed.emit()

    @property
    def optimize_for_type(self) -> str | None:
        """Selected optimization type, or None."""
        return self._optimize_for_type

    @optimize_for_type.setter
    def optimize_for_type(self, value: str | None) -> None:
        normalized = str(value).strip() if value else None
        if normalized is not None and self._optimization_choices_result.get(normalized) is None:
            normalized = None
        if self._optimize_for_type != normalized:
            self._optimize_for_type = normalized
            self.state_changed.emit()

    @property
    def optimization_choices_result(self) -> OptimizationChoicesResult:
        """Runtime-backed choices and their explicit ready/failed state."""

        return self._optimization_choices_result

    @property
    def optimization_choices(self) -> tuple[OptimizationChoice, ...]:
        return self._optimization_choices_result.choices

    @property
    def optimization_route_options(self) -> tuple[str, ...]:
        choice = self._optimization_choices_result.get(self._optimize_for_type)
        return choice.route_options if choice is not None else ()

    @property
    def doc_remove_numbering(self) -> bool:
        """Whether to remove existing numbering (docx→md)."""
        return self._doc_remove_numbering

    @doc_remove_numbering.setter
    def doc_remove_numbering(self, value: bool) -> None:
        if self._doc_remove_numbering != value:
            self._doc_remove_numbering = value
            self.state_changed.emit()

    @property
    def doc_add_numbering(self) -> bool:
        """Whether to add new numbering (docx→md)."""
        return self._doc_add_numbering

    @doc_add_numbering.setter
    def doc_add_numbering(self, value: bool) -> None:
        if self._doc_add_numbering != value:
            self._doc_add_numbering = value
            self.state_changed.emit()

    @property
    def doc_numbering_scheme(self) -> str:
        """Selected numbering scheme ID (docx→md)."""
        return self._doc_numbering_scheme

    @doc_numbering_scheme.setter
    def doc_numbering_scheme(self, value: str) -> None:
        if self._doc_numbering_scheme != value:
            self._doc_numbering_scheme = value
            self.state_changed.emit()

    @property
    def md_remove_numbering(self) -> bool:
        """Whether to remove existing numbering (md→docx)."""
        return self._md_remove_numbering

    @md_remove_numbering.setter
    def md_remove_numbering(self, value: bool) -> None:
        if self._md_remove_numbering != value:
            self._md_remove_numbering = value
            self.state_changed.emit()

    @property
    def md_add_numbering(self) -> bool:
        """Whether to add new numbering (md→docx)."""
        return self._md_add_numbering

    @md_add_numbering.setter
    def md_add_numbering(self, value: bool) -> None:
        if self._md_add_numbering != value:
            self._md_add_numbering = value
            self.state_changed.emit()

    @property
    def md_numbering_scheme(self) -> str:
        """Selected numbering scheme ID (md→docx)."""
        return self._md_numbering_scheme

    @md_numbering_scheme.setter
    def md_numbering_scheme(self, value: str) -> None:
        if self._md_numbering_scheme != value:
            self._md_numbering_scheme = value
            self.state_changed.emit()

    @property
    def md_heading_numbering_render_mode(self) -> str:
        """Return the current Settings-owned render mode for MD→DOCX.

        This preference no longer has an operation-panel control.  Read it
        from the config port on demand so a Settings change made while this
        panel is already open is reflected by the very next request rather
        than by whichever value happened to be copied during ``setup``.
        """

        value = str(
            self._read_md_to_doc_default(
                "text.heading_numbering_render_mode",
                "text",
            )
            or "text"
        )
        return value if value in {"text", "word_native"} else "text"

    @property
    def show_numbering(self) -> bool:
        """Whether numbering options are visible."""
        return self._show_numbering

    @property
    def show_optimize(self) -> bool:
        """Whether optimization options are visible."""
        return self._show_optimize

    @property
    def action_name(self) -> str:
        """Resolved action name for the current conversion request.

        The combobox stores a public resource ID.  Runtime owns the separate
        internal action name used by the request.
        """
        choice = self._optimization_choices_result.get(self._optimize_for_type)
        return choice.action_name if choice is not None else ""

    # ── MD→Document/Spreadsheet properties ───────────────────────────────

    @property
    def target_format(self) -> str:
        """Selected target format for MD→Document/Spreadsheet."""
        return self._target_format

    @target_format.setter
    def target_format(self, value: str) -> None:
        if self._target_format != value:
            self._target_format = value
            self.state_changed.emit()

    @property
    def available_target_formats(self) -> list[str]:
        """Available target format options for the current mode."""
        return list(self._available_target_formats)

    @property
    def target_route_choices_result(self) -> RuntimeRouteChoicesResult:
        """Canonical Runtime routes behind the MD generation target picker."""

        return self._target_route_choices_result

    @property
    def target_route_options(self) -> tuple[str, ...]:
        """Canonical options supported by the selected MD generation route."""

        choice = self._target_route_choices_result.get(self._target_format)
        return choice.options if choice is not None else ()

    @property
    def last_document_format(self) -> str:
        """Last used document format (persisted to config)."""
        return self._last_document_format

    @last_document_format.setter
    def last_document_format(self, value: str) -> None:
        if self._last_document_format != value:
            self._last_document_format = value
            self.state_changed.emit()

    @property
    def last_spreadsheet_format(self) -> str:
        """Last used spreadsheet format (persisted to config)."""
        return self._last_spreadsheet_format

    @last_spreadsheet_format.setter
    def last_spreadsheet_format(self, value: str) -> None:
        if self._last_spreadsheet_format != value:
            self._last_spreadsheet_format = value
            self.state_changed.emit()

    # ── Proofread option properties ──────────────────────────────────────

    @property
    def proofread_options(self) -> dict[str, bool]:
        """Current proofread option states."""
        return dict(self._proofread_options)

    def set_proofread_option(self, key: str, checked: bool) -> None:
        """Set a proofread option value."""
        if key not in PROOFREAD_OPTION_KEYS:
            raise ValueError(f"Unknown proofread option: {key!r}")
        if self._proofread_options.get(key) != checked:
            self._proofread_options[key] = checked
            self.state_changed.emit()

    @property
    def show_proofread(self) -> bool:
        """Whether proofread options are visible (MD→Document mode)."""
        return self._show_proofread

    # ── Button label helpers ────────────────────────────────────────────

    def get_button_label(self) -> str:
        """Get the primary action button label for the current mode."""
        if self._is_file_to_md_mode():
            return _t("action_area.document.export_markdown", "Convert to MD")
        return _t("action_area.generate", "Generate")

    def get_button_tooltip(self) -> str:
        """Get the primary action button tooltip for the current mode."""
        key_map = {
            MODE_DOCUMENT: "action_area.document.export_markdown_tooltip",
            MODE_SPREADSHEET: "action_area.spreadsheet.export_markdown_tooltip",
            MODE_IMAGE: "action_area.image.export_markdown_tooltip",
            MODE_LAYOUT: "action_area.layout.export_markdown_tooltip",
            MODE_MD_TO_DOCUMENT: "action_area.md_to_document.generate_tooltip",
            MODE_MD_TO_SPREADSHEET: "action_area.md_to_spreadsheet.generate_tooltip",
        }
        key = key_map.get(self._file_type or "", "action_area.document.export_markdown_tooltip")
        return _t(key, "Convert file to Markdown")

    # ── 7 setup_for_* methods ───────────────────────────────────────────

    def setup_for_document_file(
        self,
        file_path: str,
        detected_format: str = "docx",
        *,
        source_inputs: tuple[OptimizationSource, ...] | None = None,
    ) -> None:
        """Mode 1: Set up for Document → MD conversion."""
        self._set_file_to_md_common(
            file_type=MODE_DOCUMENT,
            file_path=file_path,
            show_numbering=True,
            extract_ocr=False,
            source_inputs=source_inputs
            or (OptimizationSource(detected_format=detected_format, source_category=MODE_DOCUMENT),),
        )
        logger.info("ActionArea set to document→MD mode")

    def setup_for_spreadsheet_file(
        self,
        file_path: str,
        detected_format: str = "xlsx",
        *,
        source_inputs: tuple[OptimizationSource, ...] | None = None,
    ) -> None:
        """Mode 2: Set up for Spreadsheet → MD conversion."""
        self._set_file_to_md_common(
            file_type=MODE_SPREADSHEET,
            file_path=file_path,
            show_numbering=False,
            extract_ocr=False,
            source_inputs=source_inputs
            or (OptimizationSource(detected_format=detected_format, source_category=MODE_SPREADSHEET),),
        )
        logger.info("ActionArea set to spreadsheet→MD mode")

    def setup_for_image_file(
        self,
        file_path: str,
        detected_format: str = "image",
        *,
        source_inputs: tuple[OptimizationSource, ...] | None = None,
    ) -> None:
        """Mode 3: Set up for Image → MD conversion (OCR enabled by default)."""
        self._set_file_to_md_common(
            file_type=MODE_IMAGE,
            file_path=file_path,
            show_numbering=False,
            extract_ocr=True,
            source_inputs=source_inputs
            or (OptimizationSource(detected_format=detected_format, source_category=MODE_IMAGE),),
        )
        logger.info("ActionArea set to image→MD mode")

    def setup_for_layout_file(
        self,
        file_path: str,
        detected_format: str = "pdf",
        *,
        source_inputs: tuple[OptimizationSource, ...] | None = None,
    ) -> None:
        """Mode 4: Set up for Layout → MD conversion."""
        self._set_file_to_md_common(
            file_type=MODE_LAYOUT,
            file_path=file_path,
            show_numbering=False,
            extract_ocr=False,
            source_inputs=source_inputs
            or (OptimizationSource(detected_format=detected_format, source_category=MODE_LAYOUT),),
        )
        logger.info("ActionArea set to layout→MD mode")

    def setup_for_other_file(
        self,
        file_path: str,
        detected_format: str,
        *,
        source_category: str = MODE_OTHER,
        source_inputs: tuple[OptimizationSource, ...] | None = None,
    ) -> None:
        """Mode 5: Set up for Other → MD conversion."""
        self._set_file_to_md_common(
            file_type=detected_format,
            file_path=file_path,
            show_numbering=False,
            extract_ocr=False,
            source_inputs=source_inputs
            or (OptimizationSource(detected_format=detected_format, source_category=source_category),),
        )
        logger.info("ActionArea set to %s→MD mode", detected_format)

    def setup_for_md_to_document(self, file_path: str) -> None:
        """Mode 6: Set up for MD → Document conversion.

        Includes numbering options and proofread grid.
        """
        self._clear_mutually_exclusive_action_state()
        remove_numbering = self._read_md_to_doc_default(
            "text.remove_numbering",
            True,
        )
        add_numbering = self._read_md_to_doc_default(
            "text.add_numbering",
            False,
        )
        numbering_scheme = self._read_md_to_doc_default(
            "text.numbering_scheme",
            "hierarchical_standard",
        )
        self._file_type = MODE_MD_TO_DOCUMENT
        self._file_path = file_path
        self._show_proofread = True
        self._show_numbering = True
        self._show_optimize = False
        self._extract_image = True
        self._extract_ocr = False
        self._md_remove_numbering = bool(remove_numbering)
        self._md_add_numbering = bool(add_numbering)
        self._md_numbering_scheme = str(numbering_scheme or "hierarchical_standard")
        self._set_markdown_target_routes(
            lambda target: get_category(target) == CATEGORY_DOCUMENT or target == "pdf",
            preferred=self._last_document_format,
        )
        self._proofread_options = dict(DEFAULT_PROOFREAD_OPTIONS)
        self._visible = True
        self._cancel_visible = False
        self.state_changed.emit()
        logger.info("ActionArea set to MD→Document mode")

    def setup_for_md_to_spreadsheet(self, file_path: str) -> None:
        """Mode 7: Set up for MD → Spreadsheet conversion."""
        self._clear_mutually_exclusive_action_state()
        self._file_type = MODE_MD_TO_SPREADSHEET
        self._file_path = file_path
        self._show_proofread = False
        self._show_numbering = False
        self._show_optimize = False
        self._extract_image = True
        self._extract_ocr = False
        self._set_markdown_target_routes(
            lambda target: get_category(target) == CATEGORY_SPREADSHEET,
            preferred=self._last_spreadsheet_format,
        )
        self._visible = True
        self._cancel_visible = False
        self.state_changed.emit()
        logger.info("ActionArea set to MD→Spreadsheet mode")

    # ── Internal helpers ─────────────────────────────────────────────────

    def _set_markdown_target_routes(self, predicate: Any, *, preferred: str) -> None:
        """Populate a generation picker from Runtime's canonical Markdown routes."""

        main_vm = self._main_vm
        controller = getattr(main_vm, "controller", None) if main_vm is not None else None
        discovered = discover_runtime_route_choices(
            controller,
            sources=(RuntimeRouteSource(detected_format="md", source_category="markdown"),),
            operation="conversion",
        )
        if discovered.status == "failed":
            self._target_route_choices_result = discovered
            logger.warning(
                "Runtime route discovery failed for Markdown generation; target picker disabled",
                exc_info=discovered.error,
            )
        else:
            selected = {choice.target for choice in discovered.choices if predicate(choice.target)}
            self._target_route_choices_result = discovered.select_targets(selected)
        self._available_target_formats = [target.upper() for target in self._target_route_choices_result.targets]
        preferred_target = str(preferred or "").strip().lower()
        self._target_format = (
            preferred_target
            if self._target_route_choices_result.get(preferred_target) is not None
            else (self._target_route_choices_result.targets[0] if self._target_route_choices_result.targets else "")
        )

    def _set_file_to_md_common(
        self,
        file_type: str,
        file_path: str,
        *,
        show_numbering: bool,
        extract_ocr: bool,
        source_inputs: tuple[OptimizationSource, ...],
    ) -> None:
        """Common setup for all file→MD modes."""
        self._clear_mutually_exclusive_action_state()
        self._file_type = file_type
        self._file_path = file_path
        self._show_proofread = False
        self._show_numbering = show_numbering
        self._optimization_sources = source_inputs
        main_vm = self._main_vm
        controller = getattr(main_vm, "controller", None) if main_vm is not None else None
        from docwen_gui.i18n import get_locale

        self._optimization_choices_result = discover_optimization_choices(
            controller,
            locale=get_locale(),
            sources=source_inputs,
        )
        self._show_optimize = self._optimization_choices_result.status == "failed" or bool(
            self._optimization_choices_result.choices
        )
        if self._optimization_choices_result.status == "failed":
            logger.warning(
                "Optimization discovery unavailable for %s: %s",
                source_inputs,
                self._optimization_choices_result.error,
            )
        section = file_type if file_type in {MODE_DOCUMENT, MODE_SPREADSHEET, MODE_IMAGE, MODE_LAYOUT} else "other"
        self._extract_image = bool(self._read_file_to_md_default(section, "to_md_keep_images", True))
        self._extract_ocr = bool(self._read_file_to_md_default(section, "to_md_enable_ocr", extract_ocr))
        if file_type == MODE_DOCUMENT:
            self._doc_remove_numbering = bool(self._read_file_to_md_default("document", "to_md_remove_numbering", True))
            self._doc_add_numbering = bool(self._read_file_to_md_default("document", "to_md_add_numbering", False))
            self._doc_numbering_scheme = str(
                self._read_file_to_md_default(
                    "document",
                    "to_md_default_scheme",
                    "hierarchical_standard",
                )
                or "hierarchical_standard"
            )
        else:
            self._doc_remove_numbering = True
            self._doc_add_numbering = False
            self._doc_numbering_scheme = "gongwen_standard"
        if self._show_optimize:
            enable_optimization = bool(self._read_file_to_md_default(section, "to_md_enable_optimization", False))
            optimization_type = self._read_file_to_md_default(section, "to_md_optimization_type", None)
            candidate_id = str(optimization_type).strip() if enable_optimization and optimization_type else None
            self._optimize_for_type = (
                candidate_id if self._optimization_choices_result.get(candidate_id) is not None else None
            )
        else:
            self._optimize_for_type = None
        self._visible = True
        self._cancel_visible = False
        self.state_changed.emit()

    def _clear_mutually_exclusive_action_state(self) -> None:
        """Clear action state that belongs to a different setup branch."""
        self._optimize_for_type = None
        self._optimization_choices_result = OptimizationChoicesResult(status="ready", choices=())
        self._optimization_sources = ()
        self._target_route_choices_result = RuntimeRouteChoicesResult(status="empty", choices=())
        self._aggregate_action_name = ""
        self._aggregate_file_list = []

    def _read_md_to_doc_default(self, key: str, default: Any) -> Any:
        """Read an MD→DOCX default from config, falling back safely."""
        return self._read_config_default(key, default)

    def _read_file_to_md_default(self, section: str, key: str, default: Any) -> Any:
        """Read a file→MD default from the new registry-driven config (no ``defaults.`` wrapper)."""
        return self._read_config_default(f"{section}.{key}", default)

    def _read_export_file_to_md_default(self, key: str, default: Any) -> Any:
        """Read the single Export-owned file→Markdown default."""
        return self._read_config_default(f"export.{key}", default)

    def _file_to_md_section(self) -> str:
        """Return the config section for the current file→MD mode.

        This helper is only valid inside the file→Markdown branch guarded by
        :meth:`_is_file_to_md_mode`.  Keeping that invariant explicit prevents
        non-file modes from being mistaken for the ``other`` config section.
        """
        file_type = self._file_type
        if file_type is None or file_type in _NON_FILE_TO_MD_TYPES:
            raise RuntimeError("file-to-Markdown config requested outside a file-to-Markdown mode")
        if file_type in {MODE_DOCUMENT, MODE_SPREADSHEET, MODE_IMAGE, MODE_LAYOUT}:
            return file_type
        return MODE_OTHER

    def _is_file_to_md_mode(self) -> bool:
        """Return whether the current mode converts a source file to Markdown.

        ``setup_for_other_file`` preserves the concrete format (for example
        ``epub`` or ``pptx``) as ``file_type`` so the runtime request can keep
        source-specific route metadata.
        """
        return self._file_type is not None and self._file_type not in _NON_FILE_TO_MD_TYPES

    def _file_to_md_image_mode_option(self) -> str | None:
        section = self._file_to_md_section()
        if section not in {MODE_DOCUMENT, MODE_SPREADSHEET, MODE_IMAGE, MODE_LAYOUT} and (
            self._file_type not in _MARKUP_FILE_TO_MD_TYPES and self._file_type not in _PRESENTATION_FILE_TO_MD_TYPES
        ):
            return None
        value = self._read_export_file_to_md_default("to_md_image_extraction_mode", None)
        return str(value).strip() if value else None

    def _file_to_md_ocr_placement_option(self) -> str | None:
        section = self._file_to_md_section()
        if section == MODE_LAYOUT:
            return None
        if (
            section not in {MODE_DOCUMENT, MODE_SPREADSHEET, MODE_IMAGE}
            and self._file_type not in _MARKUP_FILE_TO_MD_TYPES
            and self._file_type not in _PRESENTATION_FILE_TO_MD_TYPES
        ):
            return None
        value = self._read_export_file_to_md_default("to_md_ocr_placement_mode", None)
        return str(value).strip() if value else None

    def _file_to_md_ocr_language_option(self) -> str | None:
        if not self._is_file_to_md_mode():
            return None
        value = self._read_config_default("image.ocr_language", None)
        return str(value).strip() if value else None

    def _file_to_md_image_link_style_option(self) -> str | None:
        if self._file_type not in {
            MODE_DOCUMENT,
            MODE_SPREADSHEET,
            MODE_IMAGE,
            MODE_LAYOUT,
            *_PRESENTATION_FILE_TO_MD_TYPES,
        }:
            return None
        value = self._read_config_default("link.format.image_link_style", None)
        return str(value).strip() if value else None

    def _file_to_md_table_merge_strategy_option(self) -> str | None:
        if self._file_type not in {MODE_DOCUMENT, MODE_SPREADSHEET}:
            return None
        section = "document" if self._file_type == MODE_DOCUMENT else "spreadsheet"
        value = self._read_file_to_md_default(section, "to_md_table_merge_export_strategy", None)
        return str(value).strip() if value else None

    def _read_config_default(self, key: str, default: Any) -> Any:
        """Read a config value from the controller chain, falling back safely."""
        main_vm = self._main_vm
        if main_vm is None:
            return default
        controller = getattr(main_vm, "controller", None)
        if controller is None:
            return default
        cfg_port = getattr(controller, "config_port", None)
        if cfg_port is None:
            return default
        try:
            return cfg_port.get(key, default)
        except Exception:
            logger.warning("Config read failed; using default (key=%s, stage=action-area)", key, exc_info=True)
            return default

    def numbering_scheme_config(self) -> object:
        """Return the injected numbering scheme snapshot for GUI projection."""

        return self._read_config_default("numbering.add", {})

    # ── Command methods ──────────────────────────────────────────────────

    def show(self) -> None:
        """Show the ActionArea."""
        self._visible = True
        self._cancel_visible = False
        self.state_changed.emit()

    def hide(self) -> None:
        """Hide the ActionArea."""
        self._visible = False
        self._cancel_visible = False
        self.state_changed.emit()

    def show_cancel(self) -> None:
        """Switch to the cancel page."""
        self._cancel_visible = True
        self.state_changed.emit()

    def hide_cancel(self) -> None:
        """Switch back to the action page."""
        self._cancel_visible = False
        self.state_changed.emit()

    def request_cancel(self) -> None:
        """Handle cancel button click."""
        logger.info("Cancel requested")
        self.cancel_requested.emit()

    def request_conversion(
        self,
        target_format: str | None = None,
        file_path: str | None = None,
        options: dict[str, Any] | None = None,
    ) -> None:
        """Request a conversion action through the ViewModel.

        Args:
            target_format: Target format (uses self._target_format if None).
            file_path: Source file (uses self._file_path if None).
            options: Conversion options (built from current state if None).
        """
        fp = file_path or self._file_path
        if not fp:
            logger.warning("Conversion requested without a file path")
            return

        fmt = target_format or self._target_format
        if options is None:
            options = self.collect_options()

        logger.info("Conversion requested: %s -> %s", fp, fmt)
        self.conversion_requested.emit(fmt, fp, options)

    def collect_options(self) -> dict[str, Any]:
        """Build the full options dict from current ViewModel state."""
        options: dict[str, Any] = {}

        # File→MD options
        if self._is_file_to_md_mode():
            if self._show_numbering and self._file_type == MODE_DOCUMENT:
                options["remove_numbering"] = self._doc_remove_numbering
                options["add_numbering"] = self._doc_add_numbering
                options["numbering_scheme"] = self._doc_numbering_scheme
            options.update(
                build_to_markdown_options(
                    keep_images=self._extract_image,
                    enable_ocr=self._extract_ocr,
                    image_mode=self._file_to_md_image_mode_option(),
                    ocr_placement=self._file_to_md_ocr_placement_option(),
                    ocr_language=self._file_to_md_ocr_language_option(),
                    image_link_style=self._file_to_md_image_link_style_option(),
                    table_merge_strategy=self._file_to_md_table_merge_strategy_option(),
                )
            )
            # NOTE: optimize_for_type is no longer written to options.
            # It is surfaced as action_name via the ViewModel property
            # and consumed in main_window._handle_conversion_requested().

        # MD→Document options
        if self._file_type == MODE_MD_TO_DOCUMENT:
            options.update(self._proofread_options)
            options["remove_numbering"] = self._md_remove_numbering
            options["add_numbering"] = self._md_add_numbering
            options["numbering_scheme"] = self._md_numbering_scheme
            options["heading_numbering_render_mode"] = self.md_heading_numbering_render_mode

        # MD→Spreadsheet: only pass options for xlsx
        if self._file_type == MODE_MD_TO_SPREADSHEET and self._target_format != "xlsx":
            return {}

        return options

    # ── Aggregate mode methods ───────────────────────────────────────────

    def setup_for_aggregate(self, action_name: str, file_list: list[str]) -> None:
        """Set up the ActionArea for an aggregate (merge) operation.

        Aggregate operations combine multiple input files into a single
        output (merge-pdfs, merge-tables, merge-images-to-tiff).

        Args:
            action_name: One of ``"merge_pdfs"``, ``"merge_tables"``,
                ``"merge_images_to_tiff"``.
            file_list: The files to merge (at least 2).
        """
        self._clear_mutually_exclusive_action_state()
        self._file_type = MODE_AGGREGATE
        self._file_path = file_list[0] if file_list else None
        self._mode = "aggregate"
        self._show_proofread = False
        self._show_numbering = False
        self._show_optimize = False
        self._visible = True
        self._cancel_visible = False
        # Store the action name and file list for request context
        self._aggregate_action_name = action_name
        self._aggregate_file_list = list(file_list)
        self.state_changed.emit()
        logger.info("ActionArea set to aggregate mode: %s (%d files)", action_name, len(file_list))

    def collect_aggregate_request_context(
        self,
    ) -> tuple[str, list[str], dict[str, Any]] | None:
        """Build aggregate request context for constructing a ConversionRequest.

        Returns:
            ``(action_name, file_list, options)`` when in aggregate mode,
            or ``None`` if the current mode is not ``MODE_AGGREGATE``.
        """
        if self._file_type != MODE_AGGREGATE:
            return None
        action_name = self._aggregate_action_name
        file_list = self._aggregate_file_list
        options = self.collect_options()
        return (action_name, file_list, options)

    @property
    def is_aggregate_mode(self) -> bool:
        """Return True if the ActionArea is in aggregate mode."""
        return self._file_type == MODE_AGGREGATE

    def set_file_to_md_option(self, key: str, value: Any) -> None:
        """Set a file→MD option by key.

        Valid keys: ``"extract_image"``, ``"extract_ocr"``, ``"optimize_for_type"``,
        ``"remove_numbering"``, ``"add_numbering"``, ``"numbering_scheme"``.
        """
        if key == "extract_image":
            self.extract_image = bool(value)
        elif key == "extract_ocr":
            self.extract_ocr = bool(value)
        elif key == "optimize_for_type":
            self.optimize_for_type = value if value else None
        elif key == "remove_numbering":
            self.doc_remove_numbering = bool(value)
        elif key == "add_numbering":
            self.doc_add_numbering = bool(value)
        elif key == "numbering_scheme":
            self.doc_numbering_scheme = str(value)
        else:
            raise KeyError(f"Unknown file_to_md option: {key!r}")

    def set_md_to_doc_option(self, key: str, value: Any) -> None:
        """Set an MD→Document option by key.

        Valid keys: ``"remove_numbering"``, ``"add_numbering"`` and
        ``"numbering_scheme"``.  Heading-number rendering is Settings-owned
        and intentionally has no transient operation-panel mutation path.
        """
        if key == "remove_numbering":
            self.md_remove_numbering = bool(value)
        elif key == "add_numbering":
            self.md_add_numbering = bool(value)
        elif key == "numbering_scheme":
            self.md_numbering_scheme = str(value)
        else:
            raise KeyError(f"Unknown md_to_doc option: {key!r}")

    def save_last_document_format(self, fmt: str) -> None:
        """Persist the last used document format."""
        self._last_document_format = fmt

    def save_last_spreadsheet_format(self, fmt: str) -> None:
        """Persist the last used spreadsheet format."""
        self._last_spreadsheet_format = fmt

    def set_mode(self, mode: str) -> None:
        """Set the UI mode (single/batch)."""
        if mode not in ("single", "batch"):
            raise ValueError(f"Invalid mode: {mode!r}")
        self._mode = mode

    def reset(self) -> None:
        """Reset all state to defaults and hide."""
        self._visible = False
        self._cancel_visible = False
        self._file_type = None
        self._file_path = None
        self._mode = "single"
        self._extract_image = True
        self._extract_ocr = False
        self._optimize_for_type = None
        self._optimization_choices_result = OptimizationChoicesResult(status="ready", choices=())
        self._optimization_sources = ()
        self._target_route_choices_result = RuntimeRouteChoicesResult(status="empty", choices=())
        self._doc_remove_numbering = True
        self._doc_add_numbering = False
        self._doc_numbering_scheme = "gongwen_standard"
        self._md_remove_numbering = True
        self._md_add_numbering = False
        self._md_numbering_scheme = "hierarchical_standard"
        self._show_numbering = True
        self._show_optimize = False
        self._target_format = "docx"
        self._available_target_formats = []
        self._proofread_options = dict(DEFAULT_PROOFREAD_OPTIONS)
        self._show_proofread = False
        # Clear aggregate-specific state
        self._aggregate_action_name = ""
        self._aggregate_file_list = []
        self.state_changed.emit()
        logger.info("ActionAreaViewModel reset to default state")


__all__ = [
    "DEFAULT_PROOFREAD_OPTIONS",
    "MODE_AGGREGATE",
    "MODE_DOCUMENT",
    "MODE_IMAGE",
    "MODE_LAYOUT",
    "MODE_MD_TO_DOCUMENT",
    "MODE_MD_TO_SPREADSHEET",
    "MODE_OTHER",
    "MODE_SPREADSHEET",
    "PROOFREAD_OPTION_KEYS",
    "SENSITIVE_WORD",
    "SYMBOL_CORRECTION",
    "SYMBOL_PAIRING",
    "TYPOS_RULE",
    "ActionAreaViewModel",
]
