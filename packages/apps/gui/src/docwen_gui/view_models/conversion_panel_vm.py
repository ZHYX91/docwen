"""ConversionPanelViewModel — observable state for the ConversionPanel widget.

This is the single source of truth for the ConversionPanel's observable state.
Widgets bind to its signals and properties; user actions flow through method
calls that delegate to ``MainWindowViewModel``.

Widgets never call runtime/plugins directly — they go through this ViewModel
which delegates to the parent ``MainWindowViewModel``.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QObject, Signal

from docwen_core.formats.categories import CATEGORY_DOCUMENT, CATEGORY_IMAGE, CATEGORY_SPREADSHEET, get_category
from docwen_core.page_selection import PageSelectionError, parse_pdf_split_input
from docwen_gui.format_presentation import (
    FormatChoice,
    format_choice,
    normalize_format,
)
from docwen_gui.spreadsheet_protection import SpreadsheetProtectionInfo

from ._runtime_route_filter import (
    RuntimeRouteChoicesResult,
    RuntimeRouteSource,
    discover_composed_action_route_choices,
    discover_runtime_route_choices,
)
from .spreadsheet_analysis import SpreadsheetAnalysis

if TYPE_CHECKING:
    from .main_window_vm import MainWindowViewModel

logger = logging.getLogger(__name__)

COMPRESSIBLE_FORMATS: list[str] = ["JPG", "JPEG", "WEBP"]

# ── Validation option keys ───────────────────────────────────────────────

SYMBOL_PAIRING = "symbol_pairing"
SYMBOL_CORRECTION = "symbol_correction"
TYPOS_RULE = "typos_rule"
SENSITIVE_WORD = "sensitive_word"

VALIDATION_OPTION_KEYS = [SYMBOL_PAIRING, SYMBOL_CORRECTION, TYPOS_RULE, SENSITIVE_WORD]


class ConversionPanelViewModel(QObject):
    """Observable state for the ConversionPanel widget.

    Holds all state for the 4 category layouts (document/spreadsheet/image/layout)
    and their conversion/saveas/extra sections.

    Signals:
        state_changed: emitted when category/format changes (widgets rebind).
        conversion_requested: emitted when user clicks a convert/export/render button.
        named_action_requested: emitted for proofread/merge/split actions.
    """

    # ── Signals ──────────────────────────────────────────────────────────

    state_changed = Signal()
    """Emitted when the panel category or format changes — widgets rebind UI."""

    conversion_requested = Signal(str, str, object)
    """Emitted when user clicks convert: (target_format, file_path, options_dict)."""

    named_action_requested = Signal(str, str, object)
    """Emitted for special actions: (action_name, file_path, options_dict).
    Actions: "validate", "merge_tables", "merge_images_to_tiff", "merge_pdfs", "split_pdf".
    """

    # ── Construction ──────────────────────────────────────────────────────

    def __init__(
        self,
        main_vm: MainWindowViewModel | None = None,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._main_vm = main_vm

        # Core file state
        self._file_category: str | None = None
        self._current_format: str = ""
        self._current_file_path: str | None = None
        self._file_list: list[str] = []
        self._aggregate_counts: dict[str, int] = {}
        self._ui_mode: str = "single"
        self._route_choices_result = RuntimeRouteChoicesResult(status="empty", choices=())
        self._spreadsheet_analysis = SpreadsheetAnalysis(self)
        self._spreadsheet_analysis.changed.connect(self.state_changed)
        self._spreadsheet_sources: tuple[str, ...] = ()

        # Image section state
        self._compress_mode: str = self._normalize_compress_mode(self._read_image_default("compress_mode", "lossless"))
        self._size_limit: int = self._read_int_default("image.size_limit", 200)
        self._size_unit: str = self._normalize_size_unit(self._read_image_default("size_unit", "KB"))
        self._tiff_mode: str = self._normalize_tiff_mode(self._read_image_default("tiff_mode", "smart"))
        self._pdf_quality: str = self._normalize_pdf_quality(self._read_image_default("pdf_quality", "original"))

        # Spreadsheet section state
        self._merge_mode: int = self._normalize_merge_mode(self._read_int_default("spreadsheet.merge_mode", 3))
        self._reference_table_name: str = ""

        # Layout section state
        self._layout_overrides: set[str] = set()
        self._document_format = "DOCX"
        self._split_mode = "every_page"
        self._render_format: str = "TIF"
        self._render_dpi: int = self._normalize_render_dpi(self._read_int_default("layout.render_dpi", 300))
        self._page_input: str = ""
        self._pdf_total_pages: int = 0
        self._pdf_file_name: str = ""
        self.refresh_layout_defaults()

        # Validation options
        self._validation_options: dict[str, bool] = {
            SYMBOL_PAIRING: True,
            SYMBOL_CORRECTION: True,
            TYPOS_RULE: True,
            SENSITIVE_WORD: False,
        }

        logger.info("ConversionPanelViewModel initialized")

    # ── Properties ────────────────────────────────────────────────────────

    @property
    def file_category(self) -> str | None:
        """Current file category: ``"document"``, ``"spreadsheet"``, ``"image"``, ``"layout"``, or None."""
        return self._file_category

    @property
    def current_format(self) -> str:
        """Current file format string (lowercase)."""
        return self._current_format

    @property
    def current_file_path(self) -> str | None:
        """Current file path, or None."""
        return self._current_file_path

    @property
    def file_list(self) -> list[str]:
        """List of file paths in the current context."""
        return list(self._file_list)

    @property
    def ui_mode(self) -> str:
        """Current UI mode: ``"single"`` or ``"batch"``."""
        return self._ui_mode

    @property
    def has_files(self) -> bool:
        """Whether files are selected."""
        return self._current_file_path is not None

    @property
    def route_choices_result(self) -> RuntimeRouteChoicesResult:
        """Canonical conversion routes for the current concrete source."""

        return self._route_choices_result

    @property
    def spreadsheet_protection_info(self) -> tuple[SpreadsheetProtectionInfo, ...]:
        """Read-only protection state for the selected XLSX inputs."""

        return self._spreadsheet_analysis.results

    @property
    def spreadsheet_analysis_pending(self) -> bool:
        return self._spreadsheet_analysis.pending

    @property
    def has_xlsx_inputs(self) -> bool:
        return bool(self._spreadsheet_sources)

    def close(self) -> None:
        self._spreadsheet_analysis.close()

    @property
    def spreadsheet_protected_files(self) -> tuple[str, ...]:
        """Selected inputs for which workbook or sheet protection was detected."""

        return tuple(info.path for info in self.spreadsheet_protection_info if info.is_protected)

    @property
    def spreadsheet_unknown_files(self) -> tuple[str, ...]:
        """Selected inputs whose protection state could not be inspected."""

        return tuple(info.path for info in self.spreadsheet_protection_info if info.is_unknown)

    @property
    def spreadsheet_password_required(self) -> bool:
        """Whether an encrypted workbook requires a password before delivery."""

        return any(info.requires_password for info in self.spreadsheet_protection_info)

    def set_aggregate_counts(self, counts: dict[str, int]) -> None:
        if counts != self._aggregate_counts:
            self._aggregate_counts = dict(counts)
            self.state_changed.emit()

    def aggregate_count(self, action: str) -> int:
        return self._aggregate_counts.get(action, 0)

    # ── Image section properties ──────────────────────────────────────────

    @property
    def compress_mode(self) -> str:
        """Image compression mode: ``"lossless"`` or ``"limit_size"``."""
        return self._compress_mode

    @compress_mode.setter
    def compress_mode(self, value: str) -> None:
        if value not in ("lossless", "limit_size"):
            raise ValueError(f"Invalid compress_mode: {value!r}")
        if self._compress_mode != value:
            self._compress_mode = value
            self.state_changed.emit()

    @property
    def size_limit(self) -> int:
        """Size limit value for limit_size compression mode (1..10240 KB, 1..100 MB)."""
        return self._size_limit

    @size_limit.setter
    def size_limit(self, value: int) -> None:
        if self._size_limit != value:
            self._size_limit = value
            self.state_changed.emit()

    @property
    def size_unit(self) -> str:
        """Size unit: ``"KB"`` or ``"MB"``."""
        return self._size_unit

    @size_unit.setter
    def size_unit(self, value: str) -> None:
        if value not in ("KB", "MB"):
            raise ValueError(f"Invalid size_unit: {value!r}")
        if self._size_unit != value:
            self._size_unit = value
            self.state_changed.emit()

    @property
    def tiff_mode(self) -> str:
        """TIFF merge mode: ``"smart"`` (preserve transparency) or ``"RGB"`` (no transparency)."""
        return self._tiff_mode

    @tiff_mode.setter
    def tiff_mode(self, value: str) -> None:
        value = self._normalize_tiff_mode(value)
        if value not in ("smart", "RGB"):
            raise ValueError(f"Invalid tiff_mode: {value!r}")
        if self._tiff_mode != value:
            self._tiff_mode = value
            self.state_changed.emit()

    @property
    def pdf_quality(self) -> str:
        """Image→PDF quality: ``"original"``, ``"a4"``, or ``"a3"``."""
        return self._pdf_quality

    @pdf_quality.setter
    def pdf_quality(self, value: str) -> None:
        value = self._normalize_pdf_quality(value)
        if value not in ("original", "a4", "a3"):
            raise ValueError(f"Invalid pdf_quality: {value!r}")
        if self._pdf_quality != value:
            self._pdf_quality = value
            self.state_changed.emit()

    # ── Spreadsheet section properties ────────────────────────────────────

    @property
    def merge_mode(self) -> int:
        """Merge mode: 1=row, 2=column, 3=cell."""
        return self._merge_mode

    @merge_mode.setter
    def merge_mode(self, value: int) -> None:
        if value not in (1, 2, 3):
            raise ValueError(f"Invalid merge_mode: {value}")
        if self._merge_mode != value:
            self._merge_mode = value
            self.state_changed.emit()

    @property
    def reference_table_name(self) -> str:
        """Name of the reference/base table for spreadsheet merge."""
        return self._reference_table_name

    @reference_table_name.setter
    def reference_table_name(self, value: str) -> None:
        if self._reference_table_name != value:
            self._reference_table_name = value
            self.state_changed.emit()

    # ── Layout section properties ─────────────────────────────────────────

    @property
    def document_format(self) -> str:
        return self._document_format

    @document_format.setter
    def document_format(self, value: str) -> None:
        self._layout_overrides.add("document_format")
        if self._document_format != value:
            self._document_format = value
            self.state_changed.emit()

    @property
    def split_mode(self) -> str:
        return self._split_mode

    @split_mode.setter
    def split_mode(self, value: str) -> None:
        if value not in {"every_page", "odd_even", "custom"}:
            raise ValueError(f"Invalid split mode: {value!r}")
        self._layout_overrides.add("split_mode")
        if self._split_mode != value:
            self._split_mode = value
            self.state_changed.emit()

    def refresh_layout_defaults(self, *, emit: bool = True) -> None:
        """Refresh committed defaults while retaining explicit operation choices."""
        for key, fallback in (("document_format", "DOCX"), ("render_format", "TIF"), ("split_mode", "every_page")):
            if key not in self._layout_overrides:
                value = str(self._read_config_default(f"layout.{key}", fallback))
                if key != "split_mode":
                    value = value.upper()
                allowed = {"every_page", "odd_even", "custom"} if key == "split_mode" else None
                if allowed is not None and value not in allowed:
                    value = fallback
                setattr(self, f"_{key}", value)
        if "render_dpi" not in self._layout_overrides:
            self._render_dpi = self._normalize_render_dpi(self._read_int_default("layout.render_dpi", 300))
        if emit:
            self.state_changed.emit()

    def split_expression(self) -> str:
        """Adapt the UI mode to the existing bounded page-selection protocol."""
        return {"every_page": "*", "odd_even": "#"}.get(self._split_mode, self._page_input.strip())

    @property
    def render_format(self) -> str:
        """Layout render format: ``"TIF"`` or ``"JPG"``."""
        return self._render_format

    @render_format.setter
    def render_format(self, value: str) -> None:
        if value not in ("PNG", "TIF", "JPG"):
            raise ValueError(f"Invalid render_format: {value!r}")
        self._layout_overrides.add("render_format")
        if self._render_format != value:
            self._render_format = value
            self.state_changed.emit()

    @property
    def render_dpi(self) -> int:
        """Layout render DPI: 150, 300, or 600."""
        return self._render_dpi

    @render_dpi.setter
    def render_dpi(self, value: int) -> None:
        if value not in (150, 300, 600):
            raise ValueError(f"Invalid render_dpi: {value}")
        self._layout_overrides.add("render_dpi")
        if self._render_dpi != value:
            self._render_dpi = value
            self.state_changed.emit()

    @property
    def page_input(self) -> str:
        """Page range input text."""
        return self._page_input

    @page_input.setter
    def page_input(self, value: str) -> None:
        if self._page_input != value:
            self._page_input = value
            self.state_changed.emit()

    @property
    def pdf_page_count_unavailable(self) -> bool:
        """Native PDF split needs readable metadata; bridge inputs obtain it after conversion."""
        return self._current_format == "pdf" and self._pdf_total_pages == 0

    @property
    def pdf_total_pages(self) -> int:
        """Total pages of the selected PDF."""
        return self._pdf_total_pages

    @pdf_total_pages.setter
    def pdf_total_pages(self, value: int) -> None:
        self.set_pdf_info(value, self._pdf_file_name)

    @property
    def pdf_file_name(self) -> str:
        """File name of the selected PDF."""
        return self._pdf_file_name

    @pdf_file_name.setter
    def pdf_file_name(self, value: str) -> None:
        if self._pdf_file_name != value:
            self._pdf_file_name = value
            self.state_changed.emit()

    def set_pdf_info(self, total_pages: int, file_name: str) -> None:
        """Set selected-PDF metadata used by layout split controls."""
        if not isinstance(total_pages, int) or isinstance(total_pages, bool) or total_pages < 0:
            raise ValueError("PDF page count must be a non-negative integer.")
        # Metadata may describe a source larger than the split budget. Keep
        # that fact for other routes; split validation independently refuses it.
        normalized_pages = total_pages
        normalized_name = str(file_name or "")
        if self._pdf_total_pages == normalized_pages and self._pdf_file_name == normalized_name:
            return
        self._pdf_total_pages = normalized_pages
        self._pdf_file_name = normalized_name
        self.state_changed.emit()

    # ── Validation option properties ──────────────────────────────────────

    @property
    def validation_options(self) -> dict[str, bool]:
        """Current proofread validation option states."""
        return dict(self._validation_options)

    def set_validation_option(self, key: str, checked: bool) -> None:
        """Set a validation option value."""
        if key not in VALIDATION_OPTION_KEYS:
            raise ValueError(f"Unknown validation option: {key!r}")
        if self._validation_options.get(key) != checked:
            self._validation_options[key] = checked
            self.state_changed.emit()

    @property
    def is_any_validation_option_checked(self) -> bool:
        """Whether at least one validation option is enabled."""
        return any(self._validation_options.values())

    # ── Command methods (called by widgets → may delegate to main_vm) ─────

    def set_file_info(
        self,
        category: str,
        current_format: str,
        file_path: str | None = None,
        file_list: list[str] | None = None,
        ui_mode: str = "single",
        source_formats: dict[str, str] | None = None,
    ) -> None:
        """Set the current file context — called when file selection changes.

        Args:
            category: File category (document/spreadsheet/image/layout).
            current_format: Current format string (lowercase).
            file_path: Single file path or None.
            file_list: List of file paths (for batch mode context).
            ui_mode: ``"single"`` or ``"batch"``.
        """
        logger.info(
            "Setting file info: category=%s, format=%s, mode=%s",
            category,
            current_format,
            ui_mode,
        )
        if (file_path, category, ui_mode) != (self._current_file_path, self._file_category, self._ui_mode):
            self._layout_overrides.clear()
            self.refresh_layout_defaults(emit=False)
        if file_path != self._current_file_path:
            self._page_input = ""
            self._pdf_total_pages = 0
            self._pdf_file_name = ""
        self._file_category = category
        self._current_format = current_format.lower() if current_format else ""
        self._current_file_path = file_path
        self._reference_table_name = Path(file_path).name if category == "spreadsheet" and file_path else ""
        self._file_list = list(file_list or [])
        self._ui_mode = ui_mode
        formats = (
            source_formats
            if ui_mode == "batch" and source_formats is not None
            else ({file_path: self._current_format} if file_path else {})
        )
        self._spreadsheet_sources = tuple(path for path, fmt in formats.items() if fmt == "xlsx")
        main_vm = self._main_vm
        controller = getattr(main_vm, "controller", None) if main_vm is not None else None
        self._route_choices_result = discover_runtime_route_choices(
            controller,
            sources=tuple(
                RuntimeRouteSource(detected_format=fmt, source_category=category)
                for fmt in dict.fromkeys(formats.values() or [self._current_format])
            ),
            operation="conversion",
        )
        if self._route_choices_result.status == "failed":
            logger.warning(
                "Runtime route discovery failed; conversion panel disabled (stage=set-file-info, format=%s)",
                self._current_format,
            )
        self._spreadsheet_analysis.select(self._spreadsheet_sources)
        self.state_changed.emit()

    def reset(self) -> None:
        """Reset the panel to its default (no file) state."""
        self._file_category = None
        self._current_format = ""
        self._current_file_path = None
        self._file_list = []
        self._ui_mode = "single"
        self._route_choices_result = RuntimeRouteChoicesResult(status="empty", choices=())
        self._spreadsheet_sources = ()
        self._spreadsheet_analysis.select(())
        self._compress_mode = self._normalize_compress_mode(self._read_image_default("compress_mode", "lossless"))
        self._size_limit = self._read_int_default("image.size_limit", 200)
        self._size_unit = self._normalize_size_unit(self._read_image_default("size_unit", "KB"))
        self._tiff_mode = self._normalize_tiff_mode(self._read_image_default("tiff_mode", "smart"))
        self._pdf_quality = self._normalize_pdf_quality(self._read_image_default("pdf_quality", "original"))
        self._merge_mode = self._normalize_merge_mode(self._read_int_default("spreadsheet.merge_mode", 3))
        self._reference_table_name = ""
        self._layout_overrides.clear()
        self.refresh_layout_defaults(emit=False)
        self._render_dpi = self._normalize_render_dpi(self._read_int_default("layout.render_dpi", 300))
        self._page_input = ""
        self._pdf_total_pages = 0
        self._pdf_file_name = ""
        self._validation_options = {
            SYMBOL_PAIRING: True,
            SYMBOL_CORRECTION: True,
            TYPOS_RULE: True,
            SENSITIVE_WORD: False,
        }
        self.state_changed.emit()
        logger.info("ConversionPanelViewModel reset to default state")

    # ── Config defaults ──────────────────────────────────────────────────

    def _read_image_default(self, key: str, default: str) -> str:
        value = self._read_config_default(f"image.{key}", default)
        return str(value or default)

    def _read_int_default(self, key: str, default: int) -> int:
        value = self._read_config_default(key, default)
        if isinstance(value, bool):
            return default
        if isinstance(value, int):
            return value
        return default

    def _read_config_default(self, key: str, default: object) -> object:
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
            logger.warning("Config read failed; using default (key=%s, stage=conversion-panel)", key, exc_info=True)
            return default

    @staticmethod
    def _normalize_pdf_quality(value: str) -> str:
        aliases = {"fit_a4": "a4", "fit_a3": "a3"}
        return aliases.get(str(value).strip().lower(), str(value).strip().lower())

    @staticmethod
    def _normalize_tiff_mode(value: str) -> str:
        normalized = str(value).strip()
        return "RGB" if normalized.lower() == "rgb" else normalized

    @staticmethod
    def _normalize_compress_mode(value: str) -> str:
        normalized = str(value).strip().lower()
        return normalized if normalized in {"lossless", "limit_size"} else "lossless"

    @staticmethod
    def _normalize_size_unit(value: str) -> str:
        normalized = str(value).strip().upper()
        return normalized if normalized in {"KB", "MB"} else "KB"

    @staticmethod
    def _normalize_merge_mode(value: int) -> int:
        return value if value in (1, 2, 3) else 3

    @staticmethod
    def _normalize_render_dpi(value: int) -> int:
        return value if value in (150, 300, 600) else 300

    # ── Action methods ───────────────────────────────────────────────────

    def request_conversion(
        self,
        target_format: str,
        file_path: str | None = None,
        options: dict[str, Any] | None = None,
    ) -> None:
        """Request a format conversion action.

        Args:
            target_format: Target format (lowercase).
            file_path: File to convert (defaults to current_file_path).
            options: Additional conversion options.
        """
        fp = file_path or self._current_file_path
        if not fp:
            logger.warning("Conversion requested without a file path")
            return
        normalized_target = str(target_format or "").strip().lower()
        route = self._route_choices_result.get(normalized_target)
        if route is None:
            route = next(
                (
                    choice
                    for choice in self._route_choices_result.choices
                    if normalize_format(choice.target) == normalize_format(normalized_target)
                ),
                None,
            )
        if route is None:
            logger.warning("Conversion target is not backed by a current runtime route: %s", normalized_target)
            return
        if self._file_category == CATEGORY_IMAGE and get_category(normalized_target) == CATEGORY_IMAGE:
            choice = format_choice(
                normalized_target,
                current_format=self._current_format,
                compression_mode=self._compress_mode,
            )
            if not choice.enabled:
                logger.warning("Image conversion target rejected by presentation contract: %s", normalized_target)
                return
        opts = options or {}
        logger.info("Conversion requested: %s -> %s", fp, target_format)
        self.conversion_requested.emit(route.target, fp, opts)

    def request_named_action(
        self,
        action_name: str,
        file_path: str | None = None,
        options: dict[str, Any] | None = None,
    ) -> None:
        """Request a named action (validate, merge_tables, split_pdf, etc.).

        Args:
            action_name: Action identifier (e.g., "validate", "merge_tables").
            file_path: Target file path (defaults to current_file_path).
            options: Action-specific options.
        """
        fp = file_path or self._current_file_path
        if not fp:
            logger.warning("Named action %r requested without a file path", action_name)
            return
        opts = options or {}
        logger.info("Named action requested: %s on %s", action_name, fp)
        self.named_action_requested.emit(action_name, fp, opts)

    # ── Validation helpers ────────────────────────────────────────────────

    @staticmethod
    def validate_size_input(value: str, unit: str) -> bool:
        """Validate size input for image compression.

        Args:
            value: The size string.
            unit: ``"KB"`` (1..10240) or ``"MB"`` (1..100).

        Returns:
            True if valid.
        """
        try:
            size = int(value)
            return (1 <= size <= 10240) if unit == "KB" else (1 <= size <= 100)
        except ValueError:
            return False

    @staticmethod
    def validate_page_input(input_text: str, total_pages: int = 0) -> bool:
        """Validate the complete selection before bounded expansion."""
        try:
            parse_pdf_split_input(input_text, total_pages=total_pages)
        except PageSelectionError:
            return False
        return True

    @staticmethod
    def parse_split_input(input_text: str, total_pages: int = 0) -> tuple[str, list[int] | None]:
        """Return one validated split mode and its complete custom selection."""
        selection = parse_pdf_split_input(input_text, total_pages=total_pages)
        return selection.mode, list(selection.pages) if selection.mode == "custom" else None

    @staticmethod
    def parse_page_ranges(input_text: str) -> list[int]:
        """Parse a bounded custom expression using the shared input dialect."""
        selection = parse_pdf_split_input(input_text)
        if selection.mode != "custom":
            raise PageSelectionError("invalid_mode", "A custom page expression is required.")
        return list(selection.pages)

    @staticmethod
    def normalize_format(fmt: str) -> str:
        """Normalize format aliases (jpg→jpeg, tif→tiff, heif→heic)."""
        return normalize_format(fmt)

    def get_conversion_format_choices(self) -> list[FormatChoice]:
        """Return runtime-backed targets with explicit UI availability reasons."""

        category = self._file_category or ""
        choices: list[FormatChoice] = []
        for route_choice in self._route_choices_result.choices:
            target = route_choice.target
            include = target == "pdf" if category == "layout" else get_category(target) == category
            if include:
                choices.append(
                    format_choice(
                        target,
                        current_format=self._current_format,
                        compression_mode=self._compress_mode,
                    )
                )
        return choices

    def get_conversion_formats(self) -> list[str]:
        """Get the list of conversion target formats for the current category."""
        return [choice.display_name for choice in self.get_conversion_format_choices() if choice.enabled]

    def get_layout_export_formats(self) -> list[str]:
        """Get layout→document export formats backed by actual layout routes."""
        if self._file_category != "layout":
            return []
        return [
            self._display_format(choice.target)
            for choice in self._route_choices_result.choices
            if get_category(choice.target) == CATEGORY_DOCUMENT
        ]

    def get_layout_render_formats(self) -> list[str]:
        """Get layout→image render formats backed by actual layout routes."""
        if self._file_category != "layout":
            return []
        return [
            self._display_format(choice.target)
            for choice in self._route_choices_result.choices
            if get_category(choice.target) == CATEGORY_IMAGE
        ]

    @property
    def supports_layout_pdf_operations(self) -> bool:
        """Whether both canonical PDF merge and split action routes are available."""
        if self._file_category != "layout":
            return False
        main_vm = self._main_vm
        controller = getattr(main_vm, "controller", None) if main_vm is not None else None
        source = (
            RuntimeRouteSource(
                detected_format=self._current_format,
                source_category=self._file_category,
            ),
        )
        return all(
            discover_composed_action_route_choices(
                controller,
                sources=source,
                target="pdf",
                action_name=action,
            ).get("pdf")
            is not None
            for action in ("merge_pdfs", "split_pdf")
        )

    def get_saveas_formats(self) -> list[str]:
        """Get the list of save-as formats for the current category."""
        if self._file_category not in {CATEGORY_DOCUMENT, CATEGORY_SPREADSHEET, CATEGORY_IMAGE}:
            return []
        return ["PDF"] if self._route_choices_result.get("pdf") is not None else []

    @staticmethod
    def _display_format(target: str) -> str:
        normalized = str(target or "").strip().lower()
        return "WebP" if normalized == "webp" else normalized.upper()


__all__ = [
    "COMPRESSIBLE_FORMATS",
    "SENSITIVE_WORD",
    "SYMBOL_CORRECTION",
    "SYMBOL_PAIRING",
    "TYPOS_RULE",
    "VALIDATION_OPTION_KEYS",
    "ConversionPanelViewModel",
]
