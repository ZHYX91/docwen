"""ConversionPanelViewModel — observable state for the ConversionPanel widget.

This is the single source of truth for the ConversionPanel's observable state.
Widgets bind to its signals and properties; user actions flow through method
calls that delegate to ``MainWindowViewModel``.

Widgets never call runtime/plugins directly — they go through this ViewModel
which delegates to the parent ``MainWindowViewModel``.
"""

from __future__ import annotations

import logging
import re
from typing import TYPE_CHECKING, Any

from PySide6.QtCore import QObject, Signal

from docwen_core.formats.categories import CATEGORY_DOCUMENT, CATEGORY_IMAGE, CATEGORY_SPREADSHEET, get_category
from docwen_gui.i18n import t as _t

from ._runtime_route_filter import (
    RuntimeRouteChoicesResult,
    RuntimeRouteSource,
    discover_runtime_route_choices,
)

if TYPE_CHECKING:
    from .main_window_vm import MainWindowViewModel

logger = logging.getLogger(__name__)

# ── Pure presentation constants ──────────────────────────────────────────
COMPRESSIBLE_FORMATS: list[str] = ["JPG", "JPEG", "WEBP"]

BUTTON_COLORS: dict[str, str] = {
    "DOCX": "primary",
    "DOC": "info",
    "ODT": "success",
    "RTF": "warning",
    "WPS": "info",
    "XLSX": "primary",
    "XLS": "info",
    "ODS": "success",
    "CSV": "warning",
    "TSV": "warning",
    "ET": "info",
    "PNG": "primary",
    "JPG": "primary",
    "BMP": "info",
    "GIF": "success",
    "TIF": "warning",
    "WebP": "danger",
    "PDF": "danger",
    "OFD": "success",
}

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
        self._ui_mode: str = "single"
        self._route_choices_result = RuntimeRouteChoicesResult(status="empty", choices=())

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
        self._render_format: str = "TIF"
        self._render_dpi: int = self._normalize_render_dpi(self._read_int_default("layout.render_dpi", 300))
        self._page_input: str = ""
        self._pdf_total_pages: int = 0
        self._pdf_file_name: str = ""

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
    def render_format(self) -> str:
        """Layout render format: ``"TIF"`` or ``"JPG"``."""
        return self._render_format

    @render_format.setter
    def render_format(self, value: str) -> None:
        if value not in ("PNG", "TIF", "JPG"):
            raise ValueError(f"Invalid render_format: {value!r}")
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
    def pdf_total_pages(self) -> int:
        """Total pages of the selected PDF."""
        return self._pdf_total_pages

    @pdf_total_pages.setter
    def pdf_total_pages(self, value: int) -> None:
        if self._pdf_total_pages != value:
            self._pdf_total_pages = value
            self.state_changed.emit()

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
        normalized_pages = max(0, int(total_pages))
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
        if file_path != self._current_file_path:
            self._page_input = ""
            self._pdf_total_pages = 0
            self._pdf_file_name = ""
        self._file_category = category
        self._current_format = current_format.lower() if current_format else ""
        self._current_file_path = file_path
        self._file_list = list(file_list or [])
        self._ui_mode = ui_mode
        main_vm = self._main_vm
        controller = getattr(main_vm, "controller", None) if main_vm is not None else None
        self._route_choices_result = discover_runtime_route_choices(
            controller,
            sources=(
                RuntimeRouteSource(
                    detected_format=self._current_format,
                    source_category=str(category or "").strip().lower(),
                ),
            ),
            operation="conversion",
        )
        if self._route_choices_result.status == "failed":
            logger.warning(
                "Runtime route discovery failed; conversion panel disabled (stage=set-file-info, format=%s)",
                self._current_format,
            )
        self.state_changed.emit()

    def reset(self) -> None:
        """Reset the panel to its default (no file) state."""
        self._file_category = None
        self._current_format = ""
        self._current_file_path = None
        self._file_list = []
        self._ui_mode = "single"
        self._route_choices_result = RuntimeRouteChoicesResult(status="empty", choices=())
        self._compress_mode = self._normalize_compress_mode(self._read_image_default("compress_mode", "lossless"))
        self._size_limit = self._read_int_default("image.size_limit", 200)
        self._size_unit = self._normalize_size_unit(self._read_image_default("size_unit", "KB"))
        self._tiff_mode = self._normalize_tiff_mode(self._read_image_default("tiff_mode", "smart"))
        self._pdf_quality = self._normalize_pdf_quality(self._read_image_default("pdf_quality", "original"))
        self._merge_mode = self._normalize_merge_mode(self._read_int_default("spreadsheet.merge_mode", 3))
        self._reference_table_name = ""
        self._render_format = "TIF"
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
        opts = options or {}
        logger.info("Conversion requested: %s -> %s", fp, target_format)
        self.conversion_requested.emit(target_format, fp, opts)

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
        """Validate page range input for PDF split.

        Args:
            input_text: Page range expression.
            total_pages: Total pages in the PDF (0 if unknown).

        Returns:
            True if the input is a valid page range expression.
        """
        text = input_text.strip()
        if not text:
            return False

        # Wildcards
        if text in ("*", "#"):
            return not 0 < total_pages <= 1

        # Check for valid characters
        if not re.match(r"^[\d,，;；、\-~－至\s]+$", text):
            return False

        try:
            pages = ConversionPanelViewModel.parse_page_ranges(text)
        except ValueError:
            return False

        if not pages:
            return False

        if total_pages > 0:
            max_page = max(pages)
            if max_page > total_pages:
                valid_pages = [p for p in pages if p <= total_pages]
                if not valid_pages:
                    return False
            elif set(pages) == set(range(1, total_pages + 1)):
                return False

        return True

    @staticmethod
    def parse_split_input(input_text: str) -> tuple[str, list[int] | None]:
        """Parse split input into mode and page list.

        Returns:
            Tuple of (mode, pages_or_None).
            mode is one of: ``"every_page"``, ``"odd_even"``, ``"custom"``.
        """
        text = input_text.strip()
        if text == "*":
            return ("every_page", None)
        if text == "#":
            return ("odd_even", None)
        return ("custom", ConversionPanelViewModel.parse_page_ranges(text))

    @staticmethod
    def parse_page_ranges(input_text: str) -> list[int]:
        """Parse a custom page range string into a sorted list of page numbers.

        Supports comma/semicolon/Chinese comma separators and range notation
        with ``-``, ``~``, ``－``, ``至``.
        """
        pages: set[int] = set()
        text = re.sub(r"\s+", "", input_text)
        text = re.sub(r"[，;；、]", ",", text)
        text = re.sub(r"[~－至]", "-", text)
        for part in text.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                rng = part.split("-")
                if len(rng) == 2:
                    try:
                        start, end = int(rng[0]), int(rng[1])
                        if start > end:
                            start, end = end, start
                        pages.update(range(start, end + 1))
                    except ValueError:
                        raise ValueError(
                            _t("conversion_panel.layout.invalid_page_range", "Invalid page range: {part}").format(
                                part=part
                            )
                        ) from None
            else:
                try:
                    page = int(part)
                    if page > 0:
                        pages.add(page)
                except ValueError:
                    raise ValueError(
                        _t("conversion_panel.layout.invalid_page", "Invalid page: {part}").format(part=part)
                    ) from None
        return sorted(pages)

    @staticmethod
    def normalize_format(fmt: str) -> str:
        """Normalize format aliases (jpg→jpeg, tif→tiff, heif→heic)."""
        fmt = fmt.lower()
        eq = {
            "jpg": "jpeg",
            "jpeg": "jpeg",
            "tif": "tiff",
            "tiff": "tiff",
            "heif": "heic",
            "heic": "heic",
        }
        return eq.get(fmt, fmt)

    def get_conversion_formats(self) -> list[str]:
        """Get the list of conversion target formats for the current category."""
        category = self._file_category or ""
        current = self.normalize_format(self._current_format.strip())
        formats: list[str] = []
        for choice in self._route_choices_result.choices:
            target = choice.target
            include = target == "pdf" if category == "layout" else get_category(target) == category
            if not include:
                continue
            allow_current_image_compress = (
                category == CATEGORY_IMAGE
                and self._compress_mode == "limit_size"
                and current.upper() in COMPRESSIBLE_FORMATS
            )
            if not allow_current_image_compress and self.normalize_format(target) == current:
                continue
            formats.append(self._display_format(target))
        return formats

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
            discover_runtime_route_choices(
                controller,
                sources=source,
                operation="action",
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
    "BUTTON_COLORS",
    "COMPRESSIBLE_FORMATS",
    "SENSITIVE_WORD",
    "SYMBOL_CORRECTION",
    "SYMBOL_PAIRING",
    "TYPOS_RULE",
    "VALIDATION_OPTION_KEYS",
    "ConversionPanelViewModel",
]
