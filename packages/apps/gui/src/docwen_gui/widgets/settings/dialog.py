"""Settings main dialog.

Replicates the user-visible behavior of the old ``SettingsDialog``:
- 700x800 initial, 360x320 minimum, modal
- Grouped scrolling sidebar, with a page selector at narrow widths
- Bottom bar: Reset Tab + Reset All + Ok/Cancel/Apply (QDialogButtonBox)
- Dirty tracking with auto-signal monitoring
- Unsaved-close confirmation (danger, default=no)
- Change summary label (field-level, max 10 lines in tooltip)
- Status message with auto-hide timer
"""

from __future__ import annotations

import contextlib
import logging
from collections.abc import Callable, Mapping
from typing import NamedTuple
from typing import cast as _cast

from PySide6.QtCore import QSignalBlocker, Qt, QTimer, Signal
from PySide6.QtGui import QCloseEvent, QIcon
from PySide6.QtWidgets import (
    QAbstractButton,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from docwen_gui.i18n import t
from docwen_gui.styles.design_tokens import Sizing, Spacing
from docwen_gui.styles.ui_scale import dp, set_metric
from docwen_gui.widgets.value_controls import ScrollSafeComboBox

from ...styles.theme_semantics import apply_theme_class
from ...view_models.settings_vm import SettingsViewModel
from ...view_models.template_vm import TemplateViewModel
from .base_tab import BaseSettingsTab
from .navigation import SettingsNavigationKeyboard, SettingsSidebar

logger = logging.getLogger(__name__)

# ── Dialog geometry ────────────────────────────────────────────────────────
DEFAULT_WIDTH = 700
DEFAULT_HEIGHT = 800
MIN_WIDTH = 360
MIN_HEIGHT = 320
DIALOG_PADDING = Spacing.CARD_PADDING
ACTION_BUTTON_MIN_HEIGHT = Sizing.CONTROL_HEIGHT
RESET_TAB_BUTTON_MIN_WIDTH = 116
RESET_ALL_BUTTON_MIN_WIDTH = 104
STATUS_DISPLAY_MS = 3000

_CHANGE_FIELD_LABEL_KEYS: dict[str, str] = {
    "gui.language": "settings.general.language_label",
    "gui.theme": "settings.general.theme_label",
    "gui.font_size_preset": "settings.general.font_label",
    "gui.scale_percent": "settings.general.scale_label",
    "gui.md_default_template": "settings.text.output_format",
    "output.output_mode": "settings.output.output_mode_label",
    "output.custom_path": "settings.output.custom_path_label",
    "output.create_date_subfolder": "settings.output.date_folder.create_label",
    "output.date_folder_format": "settings.output.date_folder.format_label",
    "conversion_defaults.layout.split_mode": "conversion_panel.layout.split_mode",
    "conversion_defaults.layout.document_format": "settings.layout.document_format",
    "conversion_defaults.layout.render_format": "conversion_panel.layout.image_format",
    "conversion_defaults.layout.render_dpi": "settings.layout.render_dpi_label",
    "export.ocr_language": "settings.ocr.language_label",
    "export.recognize_tables": "action_area.recognize_tables",
    "formatting.mermaid_mode": "settings.formatting.mermaid_mode_label",
    "logging.file_prefix": "settings.logging.file_prefix_label",
    "logging.directory_mode": "settings.logging.directory_mode_label",
    "logging.directory": "settings.logging.custom_directory_label",
    "conversion_defaults.image.compress_mode": "settings.image.compress_mode_label",
    "conversion_defaults.image.size_limit": "settings.image.size_limit_label",
    "conversion_defaults.image.size_unit": "settings.image.size_unit_label",
    "conversion_defaults.document.to_md_table_merge_export_strategy": "settings.table_export.merge_strategy_label",
    "conversion_defaults.spreadsheet.to_md_table_merge_export_strategy": "settings.table_export.merge_strategy_label",
    "conversion_defaults.spreadsheet.merge_mode": "settings.spreadsheet.default_merge_mode_label",
}

_CHANGE_VALUE_LABEL_KEYS: dict[str, dict[object, str]] = {
    "gui.theme": {
        "light": "settings.general.themes.light",
        "dark": "settings.general.themes.dark",
        "system": "settings.general.themes.system",
    },
    "output.output_mode": {
        "source": "settings.output.output_modes.source",
        "custom": "settings.output.output_modes.custom",
    },
    "conversion_defaults.layout.split_mode": {
        "every_page": "conversion_panel.layout.split_every_page",
        "odd_even": "conversion_panel.layout.split_odd_even",
        "custom": "conversion_panel.layout.split_custom",
    },
    "export.ocr_language": {
        "auto": "settings.ocr.language_auto",
        "chinese": "settings.ocr.language_chinese",
        "chinese_cht": "settings.ocr.language_chinese_cht",
        "english": "settings.ocr.language_english",
        "japanese": "settings.ocr.language_japanese",
        "korean": "settings.ocr.language_korean",
        "latin": "settings.ocr.language_latin",
        "cyrillic": "settings.ocr.language_cyrillic",
    },
    "formatting.mermaid_mode": {
        "code": "settings.formatting.mermaid_code",
        "image": "settings.formatting.mermaid_image",
    },
    "logging.directory_mode": {
        "user": "settings.logging.dir_modes.user",
        "temp": "settings.logging.dir_modes.temp",
        "custom": "settings.logging.dir_modes.custom",
    },
    "conversion_defaults.image.compress_mode": {
        "lossless": "settings.image.compress_lossless",
        "limit_size": "settings.image.compress_limit_size",
    },
    "conversion_defaults.document.to_md_table_merge_export_strategy": {
        "fill": "settings.table_export.strategies.fill",
        "empty": "settings.table_export.strategies.empty",
    },
    "conversion_defaults.spreadsheet.to_md_table_merge_export_strategy": {
        "fill": "settings.table_export.strategies.fill",
        "empty": "settings.table_export.strategies.empty",
    },
    "conversion_defaults.spreadsheet.merge_mode": {
        1: "settings.spreadsheet.merge_modes.by_row",
        2: "settings.spreadsheet.merge_modes.by_column",
        3: "settings.spreadsheet.merge_modes.by_cell",
    },
}


def _safe_settings_error_detail(error: Exception) -> str:
    """Return a bounded, plain-text detail for a failed settings page."""

    error_type = type(error).__name__
    try:
        message = " ".join(str(error).split())
    except Exception:
        message = ""
    if not message:
        return error_type
    return f"{error_type}: {message[:300]}"


def _read_initial_theme() -> str:
    """Read the requested theme mode before the dialog opens."""
    try:
        from docwen_gui.styles.theme_manager import ThemeManager

        return ThemeManager.get_instance().get_requested_theme()
    except Exception:
        return "light"


def _read_initial_opacity(parent: QWidget | None) -> float:
    """Read the current window opacity before the dialog opens."""
    if parent is None:
        return 1.0
    try:
        main_window = parent
        return float(main_window.windowOpacity())
    except Exception:
        return 1.0


# ── Tab order ──────────────────────────────────────────────────────────────
#
class SettingsNavigationGroup(NamedTuple):
    key: str
    page_keys: tuple[str, ...]


# These built-in groups own navigation order, not page construction or config.
# Empty groups add neither a heading nor a placeholder page.
NAVIGATION_GROUPS = (
    SettingsNavigationGroup("general", ("general", "output", "templates", "software", "logging")),
    SettingsNavigationGroup("inputs", ("text", "document", "spreadsheet", "image", "layout", "other")),
    SettingsNavigationGroup("content", ("formatting", "markdown_input", "link", "export", "proofread")),
    SettingsNavigationGroup("optimizations", ("gongwen", "invoice")),
)
TAB_KEYS = [key for group in NAVIGATION_GROUPS for key in group.page_keys]


def _navigation_group_title(key: str) -> str:
    return {
        "general": t("settings.navigation.general"),
        "inputs": t("settings.navigation.inputs"),
        "content": t("settings.navigation.content"),
        "optimizations": t("settings.navigation.optimizations"),
    }[key]


# ── Tab display names ──────────────────────────────────────────────────────
TAB_NAMES: dict[str, str] = {
    "templates": t("settings.templates.title"),
    "general": t("settings.tabs.general"),
    "text": t("settings.tabs.text"),
    "document": t("settings.tabs.document"),
    "spreadsheet": t("settings.tabs.spreadsheet"),
    "image": t("settings.tabs.image"),
    "layout": t("settings.tabs.layout"),
    "other": t("settings.tabs.other"),
    "proofread": t("settings.tabs.proofread"),
    "formatting": t("settings.tabs.formatting"),
    "markdown_input": t("settings.tabs.markdown_input"),
    "link": t("settings.tabs.link"),
    "export": t("settings.tabs.export"),
    "software": t("settings.tabs.software"),
    "output": t("settings.tabs.output"),
    "logging": t("settings.tabs.logging"),
    "gongwen": t("settings.gongwen.title"),
    "invoice": t("settings.invoice.title"),
}

_TabFactory = Callable[[SettingsViewModel], object]


class _TabSpec(NamedTuple):
    module_name: str
    class_name: str
    factory: _TabFactory


# Keep each import inside a small factory: imports remain lazy, PyInstaller can
# discover them statically, and _add_all_tabs can isolate a single broken page.
def _build_general_tab(view_model: SettingsViewModel) -> QWidget:
    from .general_tab import GeneralTab

    return GeneralTab(view_model)


def _build_text_tab(view_model: SettingsViewModel) -> QWidget:
    from .text_tab import TextTab

    return TextTab(view_model)


def _build_proofread_tab(view_model: SettingsViewModel) -> QWidget:
    from .proofread_tab import ProofreadTab

    return ProofreadTab(view_model)


def _build_document_tab(view_model: SettingsViewModel) -> QWidget:
    from .document_tab import DocumentTab

    return DocumentTab(view_model)


def _build_spreadsheet_tab(view_model: SettingsViewModel) -> QWidget:
    from .spreadsheet_tab import SpreadsheetTab

    return SpreadsheetTab(view_model)


def _build_image_tab(view_model: SettingsViewModel) -> QWidget:
    from .image_tab import ImageTab

    return ImageTab(view_model)


def _build_layout_tab(view_model: SettingsViewModel) -> QWidget:
    from .layout_tab import LayoutTab

    return LayoutTab(view_model)


def _build_link_tab(view_model: SettingsViewModel) -> QWidget:
    from .link_tab import LinkTab

    return LinkTab(view_model)


def _build_formatting_tab(view_model: SettingsViewModel) -> QWidget:
    from .formatting_tab import FormattingTab

    return FormattingTab(view_model)


def _build_markdown_input_tab(view_model: SettingsViewModel) -> QWidget:
    from .formatting_tab import MarkdownInputTab

    return MarkdownInputTab(view_model)


def _build_software_tab(view_model: SettingsViewModel) -> QWidget:
    from .software_tab import SoftwareTab

    return SoftwareTab(view_model)


def _build_output_tab(view_model: SettingsViewModel) -> QWidget:
    from .output_tab import OutputTab

    return OutputTab(view_model)


def _build_export_tab(view_model: SettingsViewModel) -> QWidget:
    from .export_tab import ExportTab

    return ExportTab(view_model)


def _build_logging_tab(view_model: SettingsViewModel) -> QWidget:
    from .logging_tab import LoggingTab

    return LoggingTab(view_model)


def _build_other_tab(view_model: SettingsViewModel) -> QWidget:
    from .other_tab import OtherTab

    return OtherTab(view_model)


def _build_gongwen_tab(view_model: SettingsViewModel) -> QWidget:
    from .gongwen_tab import GongwenTab

    return GongwenTab(view_model)


def _build_invoice_tab(view_model: SettingsViewModel) -> QWidget:
    from .invoice_tab import InvoiceTab

    return InvoiceTab(view_model)


_TAB_SPECS: dict[str, _TabSpec] = {
    "general": _TabSpec("general_tab", "GeneralTab", _build_general_tab),
    "text": _TabSpec("text_tab", "TextTab", _build_text_tab),
    "proofread": _TabSpec("proofread_tab", "ProofreadTab", _build_proofread_tab),
    "document": _TabSpec("document_tab", "DocumentTab", _build_document_tab),
    "spreadsheet": _TabSpec("spreadsheet_tab", "SpreadsheetTab", _build_spreadsheet_tab),
    "image": _TabSpec("image_tab", "ImageTab", _build_image_tab),
    "layout": _TabSpec("layout_tab", "LayoutTab", _build_layout_tab),
    "link": _TabSpec("link_tab", "LinkTab", _build_link_tab),
    "formatting": _TabSpec("formatting_tab", "FormattingTab", _build_formatting_tab),
    "markdown_input": _TabSpec("formatting_tab", "MarkdownInputTab", _build_markdown_input_tab),
    "software": _TabSpec("software_tab", "SoftwareTab", _build_software_tab),
    "output": _TabSpec("output_tab", "OutputTab", _build_output_tab),
    "export": _TabSpec("export_tab", "ExportTab", _build_export_tab),
    "logging": _TabSpec("logging_tab", "LoggingTab", _build_logging_tab),
    "other": _TabSpec("other_tab", "OtherTab", _build_other_tab),
    "gongwen": _TabSpec("gongwen_tab", "GongwenTab", _build_gongwen_tab),
    "invoice": _TabSpec("invoice_tab", "InvoiceTab", _build_invoice_tab),
}

# ── Reset group mapping (registry group to reset per tab) ──────────────────
RESET_GROUPS: dict[str, str] = {
    "general": "general",
    "text": "text",
    "proofread": "proofread",
    "document": "document",
    "spreadsheet": "spreadsheet",
    "image": "image",
    "layout": "layout",
    "link": "link",
    "formatting": "formatting",
    "markdown_input": "markdown_input",
    "software": "software",
    "output": "output",
    "export": "export",
    "logging": "logging",
    "other": "other",
    "gongwen": "gongwen",
    "invoice": "invoice",
}


def _show_confirm(parent: QWidget, title: str, message: str, danger: bool = True) -> bool:
    """Show a simple confirmation dialog.

    In production this should use a proper themed dialog.
    For now we use QDialog to avoid deep-importing old feedback modules.
    """
    from PySide6.QtWidgets import QMessageBox

    mb = QMessageBox(parent)
    mb.setWindowTitle(title)
    mb.setText(message)
    mb.setIcon(QMessageBox.Icon.Warning if danger else QMessageBox.Icon.Question)
    mb.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
    mb.button(QMessageBox.StandardButton.Yes).setText(t("common.ok"))
    mb.button(QMessageBox.StandardButton.No).setText(t("common.cancel"))
    mb.setDefaultButton(QMessageBox.StandardButton.No)
    return mb.exec() == QMessageBox.StandardButton.Yes


class SettingsDialog(QDialog):
    """Settings dialog with task-oriented navigation and dirty tracking.

    Owns a ``SettingsViewModel`` as its state source of truth.
    Delegates all Apply/Reset operations through the ViewModel.

    Args:
        parent: Parent widget (typically MainWindow).
        view_model: Pre-configured SettingsViewModel.
    """

    settings_source_changed = Signal()

    def __init__(
        self,
        parent: QWidget | None = None,
        view_model: SettingsViewModel | None = None,
        template_view_model: TemplateViewModel | None = None,
    ) -> None:
        title = t("settings.title")
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)

        self._vm = view_model or SettingsViewModel()
        self._template_vm = template_view_model or TemplateViewModel(self)
        # Capture the persisted baseline so Cancel can roll back all edits
        self._vm.begin_session()
        # Save initial visual state for Cancel rollback
        self._initial_theme = _read_initial_theme()
        from docwen_gui.styles.theme_manager import ThemeManager

        appearance = ThemeManager.get_instance()
        self._initial_font = appearance.get_font_size_preset()
        self._initial_scale = appearance.get_ui_scale()
        self._initial_opacity = _read_initial_opacity(parent)
        self._tabs: dict[str, QWidget] = {}
        self._tab_widget: QTabWidget = _cast(QTabWidget, None)
        self._navigation: SettingsSidebar | None = None
        self._navigation_keyboard: SettingsNavigationKeyboard | None = None
        self._status_timer: QTimer = _cast(QTimer, None)
        self._cancel_close_in_progress = False
        self._close_cleanup_done = False

        self._build_ui()
        self._wire_view_model()
        appearance.appearance_changed.connect(self._queue_navigation_layout)
        self._queue_navigation_layout()

    # ── UI construction ─────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.setObjectName("settingsDialog")
        self.resize(DEFAULT_WIDTH, DEFAULT_HEIGHT)
        set_metric(self, "setMinimumSize", MIN_WIDTH, MIN_HEIGHT)
        screen = self.screen()
        if screen is not None:
            available = screen.availableGeometry()
            self.resize(min(DEFAULT_WIDTH, available.width() - 32), min(DEFAULT_HEIGHT, available.height() - 64))

        layout = QVBoxLayout(self)
        set_metric(layout, "setContentsMargins", DIALOG_PADDING, DIALOG_PADDING, DIALOG_PADDING, DIALOG_PADDING)
        set_metric(layout, "setSpacing", DIALOG_PADDING)
        self._compact_navigation = ScrollSafeComboBox(self)
        self._compact_navigation.setObjectName("settingsPageSelector")
        self._compact_navigation.setAccessibleName(t("settings.title"))
        self._compact_navigation.hide()
        layout.addWidget(self._compact_navigation)

        # ── Tab widget (hidden tab bar) ─────────────────────────────────
        self._tab_widget = QTabWidget(self)
        self._tab_widget.setDocumentMode(True)
        self._tab_widget.setUsesScrollButtons(False)
        self._tab_widget.setElideMode(Qt.TextElideMode.ElideRight)
        self._tab_widget.tabBar().hide()
        self._compact_navigation.currentIndexChanged.connect(self._on_compact_page_changed)

        # ── Sidebar + content row ───────────────────────────────────────
        content_row = QHBoxLayout()
        content_row.setContentsMargins(0, 0, 0, 0)
        set_metric(content_row, "setSpacing", 8)

        self._build_navigation(content_row)
        content_row.addWidget(self._tab_widget, 1)
        layout.addLayout(content_row, 1)

        # ── Tabs ────────────────────────────────────────────────────────
        self._add_all_tabs()

        # ── Activate the optimal tab (if specified by caller) ───────────
        self._activate_initial_tab()

        # ── Bottom action bar ───────────────────────────────────────────
        action_row = QHBoxLayout()
        set_metric(action_row, "setContentsMargins", 0, 4, 0, 0)
        set_metric(action_row, "setSpacing", 8)

        reset_tab_btn = QPushButton(t("settings.reset.tab_button"), self)
        reset_tab_btn.setObjectName("settingsResetTabButton")
        set_metric(reset_tab_btn, "setMinimumHeight", ACTION_BUTTON_MIN_HEIGHT)
        set_metric(reset_tab_btn, "setMinimumWidth", RESET_TAB_BUTTON_MIN_WIDTH)
        apply_theme_class(reset_tab_btn, "secondary")
        reset_tab_btn.clicked.connect(self._on_reset_tab)
        self._tab_widget.currentChanged.connect(
            lambda _: reset_tab_btn.setEnabled(self.current_section() != "templates")
        )
        reset_tab_btn.setEnabled(self.current_section() != "templates")

        reset_all_btn = QPushButton(t("settings.reset.all_button"), self)
        reset_all_btn.setObjectName("settingsResetAllButton")
        set_metric(reset_all_btn, "setMinimumHeight", ACTION_BUTTON_MIN_HEIGHT)
        set_metric(reset_all_btn, "setMinimumWidth", RESET_ALL_BUTTON_MIN_WIDTH)
        apply_theme_class(reset_all_btn, "secondary")
        reset_all_btn.clicked.connect(self._on_reset_all)

        action_row.addWidget(reset_tab_btn)
        action_row.addWidget(reset_all_btn)
        action_row.addStretch(1)

        # Ok / Cancel / Apply via QDialogButtonBox
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
            | QDialogButtonBox.StandardButton.Apply,
            parent=self,
        )
        ok_btn = button_box.button(QDialogButtonBox.StandardButton.Ok)
        if ok_btn:
            ok_btn.setText(t("common.ok"))
            ok_btn.setObjectName("settingsOkButton")
            set_metric(ok_btn, "setMinimumHeight", ACTION_BUTTON_MIN_HEIGHT)
            apply_theme_class(ok_btn, "primary")
        cancel_btn = button_box.button(QDialogButtonBox.StandardButton.Cancel)
        if cancel_btn:
            cancel_btn.setText(t("common.cancel"))
            cancel_btn.setObjectName("settingsCancelButton")
            set_metric(cancel_btn, "setMinimumHeight", ACTION_BUTTON_MIN_HEIGHT)
            apply_theme_class(cancel_btn, "secondary")
        apply_btn = button_box.button(QDialogButtonBox.StandardButton.Apply)
        if apply_btn:
            apply_btn.setText(t("common.apply"))
            apply_btn.setObjectName("settingsApplyButton")
            set_metric(apply_btn, "setMinimumHeight", ACTION_BUTTON_MIN_HEIGHT)
            apply_theme_class(apply_btn, "secondary")

        button_box.accepted.connect(self._on_ok)
        for button in button_box.buttons():
            set_metric(button, "setMinimumWidth", Sizing.BUTTON_MIN_WIDTH)
        button_box.rejected.connect(self._on_cancel)
        if apply_btn:
            apply_btn.clicked.connect(self._on_apply)

        # ── Status label ────────────────────────────────────────────────
        self._status_label = QLabel("", self)
        self._status_label.setObjectName("settingsStatusLabel")
        self._status_label.setVisible(False)
        self._status_label.setWordWrap(True)
        layout.addWidget(self._status_label)

        # ── Changes summary label ───────────────────────────────────────
        self._changes_label = QLabel("", self)
        self._changes_label.setObjectName("settingsChangesLabel")
        self._changes_label.setVisible(False)
        self._changes_label.setWordWrap(True)
        layout.addWidget(self._changes_label)

        # Keep confirmation buttons anchored while feedback appears above them.
        layout.addLayout(action_row)
        layout.addWidget(button_box)

        self._tab_widget.currentChanged.connect(self._on_tab_changed)
        self._on_tab_changed(self._tab_widget.currentIndex())

    def _queue_navigation_layout(self) -> None:
        QTimer.singleShot(0, self, self._sync_navigation_layout)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._sync_navigation_layout()

    def _sync_navigation_layout(self) -> None:
        if not hasattr(self, "_compact_navigation"):
            return
        # Use real button size hints so translated text and large fonts fit.
        if self._navigation is not None:
            self._navigation.refresh_width()
        sidebar_width = self._navigation.width() if self._navigation is not None else 0
        compact = self.width() < sidebar_width + dp(480)
        self._compact_navigation.setVisible(compact)
        if self._navigation is not None:
            self._navigation.setVisible(not compact)
            if not compact:
                self._navigation.ensure_current_visible()
        else:
            self._tab_widget.tabBar().setVisible(not compact)

    # ── Preview restore ────────────────────────────────────────────────────

    def _restore_preview_state(self) -> None:
        """Restore only visual preview state that actually diverged.

        Reapplying the active Qt theme rebuilds the application palette/QSS and
        repolishes every visible top-level widget.  Avoid that global work on a
        routine Cancel when the user never changed the theme.  The same rule
        applies to opacity so a no-op close stays on the fast path.
        """
        from docwen_gui.styles.theme_manager import ThemeManager

        appearance = ThemeManager.get_instance()
        if appearance.get_ui_scale() != self._initial_scale:
            appearance.apply_ui_scale(self._initial_scale)
        if appearance.get_font_size_preset() != self._initial_font:
            appearance.apply_font_size_preset(self._initial_font)
        if self._initial_theme is not None and _read_initial_theme() != self._initial_theme:
            _apply_theme_no_persist(self._initial_theme)

        parent = self.parentWidget()
        if self._initial_opacity is not None and parent is not None:
            try:
                current_opacity = float(parent.windowOpacity())
            except Exception:
                current_opacity = None
            if current_opacity is None or abs(current_opacity - self._initial_opacity) > 1e-6:
                _apply_opacity_no_persist(self._initial_opacity, parent)

    def _build_navigation(self, content_row: QHBoxLayout) -> None:
        nav = SettingsSidebar(self)
        nav.page_requested.connect(self._select_navigation_page)
        self._navigation = nav
        self._navigation_keyboard = SettingsNavigationKeyboard(nav)
        self._navigation_keyboard.page_requested.connect(self._select_navigation_page)
        content_row.addWidget(nav, 0)

    def _add_nav_item(self, key: str, title: str) -> None:
        assert self._navigation is not None and self._navigation_keyboard is not None
        try:
            icon = self._load_tab_icon(key)
        except Exception:
            logger.exception("Unable to load settings navigation icon: %s", key)
            icon = None
        button = self._navigation.add_page(key, title, icon)
        self._navigation_keyboard.add(key, button, title)

    @staticmethod
    def _load_tab_icon(tab_key: str) -> QIcon | None:
        """Load per-tab SVG icon if available; falls back to None."""
        icon_name_map = {
            "general": "general.svg",
            "text": "text.svg",
            "proofread": "proofread.svg",
            "document": "document.svg",
            "spreadsheet": "spreadsheet.svg",
            "image": "image.svg",
            "layout": "layout.svg",
            "link": "link.svg",
            "formatting": "formatting.svg",
            "markdown_input": "markdown_input.svg",
            "templates": "templates.svg",
            "output": "output.svg",
            "export": "export.svg",
            "software": "software.svg",
            "logging": "logging.svg",
            "other": "other.svg",
            "gongwen": "gongwen.svg",
            "invoice": "invoice.svg",
        }
        icon_name = icon_name_map.get(tab_key)
        if icon_name is None:
            return None
        try:
            from docwen_gui.resources import load_svg_icon

            icon = load_svg_icon(icon_name)
            return icon if isinstance(icon, QIcon) and not icon.isNull() else None
        except Exception:
            return None

    def _update_nav_item_title(self, key: str, title: str) -> None:
        item = self._navigation.page_button(key) if self._navigation is not None else None
        if item is not None:
            item.setText(title)
            item.setToolTip(title)
            item.setAccessibleName(title)

    # ── Tab management ──────────────────────────────────────────────────────

    def _add_all_tabs(self) -> None:
        assert self._navigation is not None
        with QSignalBlocker(self._compact_navigation):
            for group in NAVIGATION_GROUPS:
                if not group.page_keys:
                    continue
                group_title = _navigation_group_title(group.key)
                self._navigation.add_group(group.key, group_title)
                for key in group.page_keys:
                    title = TAB_NAMES.get(key, key)
                    try:
                        tab = self._build_tab(key)
                    except Exception as error:
                        logger.exception("Failed to build Settings tab: %s", key)
                        tab = self._build_failed_tab(key, title, error)
                    self._tabs[key] = tab
                    tab.setProperty("settingsPageKey", key)
                    set_tab_title = getattr(tab, "set_tab_title", None)
                    if callable(set_tab_title):
                        set_tab_title(title)
                    tab_index = self._tab_widget.addTab(tab, title)
                    self._tab_widget.setTabToolTip(tab_index, title)
                    self._compact_navigation.addItem(f"{group_title} · {title}", key)
                    self._add_nav_item(key, title)
                    if isinstance(tab, BaseSettingsTab):
                        tab.section_requested.connect(self.activate_section)

        # Select first tab
        self._tab_widget.setCurrentIndex(0)

    def _build_tab(self, key: str) -> QWidget:
        """Construct one settings page through a PyInstaller-visible lazy import."""
        if key == "templates":
            from .templates_tab import TemplatesTab

            return TemplatesTab(view_model=self._template_vm)
        spec = _TAB_SPECS.get(key)
        if spec is None:
            raise KeyError(f"Unknown Settings tab: {key!r}")
        tab = spec.factory(self._vm)
        if not isinstance(tab, QWidget):
            raise TypeError(f"Settings tab factory {key!r} did not return a QWidget")
        return tab

    def focus_template(self, target: str, template_id: str | None) -> None:
        from .templates_tab import TemplatesTab

        page = self._tabs.get("templates")
        if isinstance(page, TemplatesTab):
            self.activate_section("templates")
            page.focus_template(target, template_id)

    def _build_failed_tab(self, key: str, title: str, error: Exception) -> QWidget:
        """Return a stable placeholder when one settings page cannot be built."""
        tab = QWidget(self._tab_widget)
        tab.setObjectName("settingsTabLoadErrorPage")
        tab.setProperty("failedTabKey", key)

        layout = QVBoxLayout(tab)
        set_metric(layout, "setContentsMargins", 16, 16, 16, 16)
        set_metric(layout, "setSpacing", 12)

        title_label = QLabel(title, tab)
        title_label.setObjectName("settingsTabTitle")
        title_label.setTextFormat(Qt.TextFormat.PlainText)
        layout.addWidget(title_label)

        message = QLabel(
            t(
                "settings.errors.tab_load_failed_message",
                "Failed to load {tab}\n\nError: {error}",
                tab=title,
                error=_safe_settings_error_detail(error),
            ),
            tab,
        )
        message.setObjectName("settingsTabLoadError")
        message.setProperty("settingsRole", "tabDescription")
        message.setTextFormat(Qt.TextFormat.PlainText)
        message.setWordWrap(True)
        layout.addWidget(message)
        layout.addStretch(1)
        return tab

    def _activate_initial_tab(self) -> None:
        """Activate the tab specified by the ViewModel's ``initial_tab_key``.

        If no initial tab was set this is a no-op — the dialog stays on
        the first (general) tab.
        """
        initial = self._vm.initial_tab_key
        if initial:
            self._select_navigation_page(initial)

    def _select_navigation_page(self, key: str) -> None:
        """Resolve navigation by page identity, including a page's error state."""
        tab = self._tabs.get(key)
        if tab is not None:
            self._tab_widget.setCurrentWidget(tab)
            if self._navigation is not None:
                self._navigation.select_page(key)

    def _on_compact_page_changed(self, index: int) -> None:
        key = self._compact_navigation.itemData(index)
        if isinstance(key, str):
            self._select_navigation_page(key)

    def activate_section(self, section: str) -> bool:
        """Select a public settings section, rejecting failed page placeholders."""

        tab = self._tabs.get(section)
        if tab is None or tab.property("failedTabKey") == section:
            return False
        self._select_navigation_page(section)
        return True

    def current_section(self) -> str | None:
        """Return the semantic key of the currently selected settings page."""

        tab = self._tab_widget.currentWidget()
        key = tab.property("settingsPageKey") if tab is not None else None
        return key if isinstance(key, str) else None

    def _wire_view_model(self) -> None:
        """Connect ViewModel signals."""
        self._vm.dirty_state_changed.connect(self._refresh_changes_summary)
        self._vm.config_changed.connect(lambda _section, _key, _value: self._refresh_changes_summary())
        self._vm.config_reloaded.connect(self._refresh_changes_summary)
        self._vm.status_changed.connect(self._show_status)

    # ── Tab changed ─────────────────────────────────────────────────────────

    def _on_tab_changed(self, index: int) -> None:
        """Sync navigation selection and set focus on first focusable widget."""
        del index
        key = self.current_section()
        if key is None:
            return
        with QSignalBlocker(self._compact_navigation):
            self._compact_navigation.setCurrentIndex(self._compact_navigation.findData(key))
        if self._navigation is not None:
            self._navigation.select_page(key)

        # Auto-focus first eligible widget in the new tab.
        tab = self._tab_widget.currentWidget()
        if tab is None:
            return
        candidate = self._find_focus_target(tab)
        if self._navigation is not None and self._navigation_keyboard is not None:
            self._navigation_keyboard.sync(key, candidate)
        if candidate is not None and not self._compact_navigation.hasFocus():
            candidate.setFocus(Qt.FocusReason.TabFocusReason)

    @staticmethod
    def _find_focus_target(tab: QWidget) -> QWidget | None:
        """Return the first eligible control, preserving the old Qt focus UX."""
        for widget in tab.findChildren(QWidget):
            if widget.objectName() == "settingsInfoButton":
                continue
            if not widget.isEnabled() or widget.focusPolicy() == Qt.FocusPolicy.NoFocus:
                continue
            if isinstance(widget, (QAbstractButton, QCheckBox, QComboBox, QLineEdit, QSpinBox, QDoubleSpinBox)):
                return widget
            if isinstance(widget, QScrollArea):
                continue
        return tab if tab.focusPolicy() != Qt.FocusPolicy.NoFocus else None

    # ── Action handlers ─────────────────────────────────────────────────────

    def _on_ok(self) -> None:
        """Apply changed settings and close.

        A clean dialog has nothing to persist.  In particular, do not send a
        full multi-file durable configuration transaction through the runtime
        merely because the user pressed OK after inspecting settings.
        """
        if not self._vm.is_dirty:
            self._close_cleanup_done = True
            self.accept()
            return

        applied = False
        try:
            applied = self._vm.ok_changes()
        except Exception:
            logger.exception("Unexpected Settings OK failure")
        finally:
            # A multi-file write can persist an early file before a later
            # failure. Consumers must always reconcile the injected source.
            self._commit_preview_state()
            self.settings_source_changed.emit()
        if applied:
            # ``accept()`` may produce a close event on some Qt/platform
            # combinations.  The settings have already been committed, so
            # that event must not run the discard path.
            self._close_cleanup_done = True
            self.accept()

    def _on_apply(self) -> None:
        """Apply settings without closing.  Updates the persisted baseline
        so that Cancel-after-Apply restores to this post-Apply state."""
        try:
            self._vm.apply_changes()
        except Exception:
            logger.exception("Unexpected Settings Apply failure")
        finally:
            self._commit_preview_state()
            self.settings_source_changed.emit()

    def _commit_preview_state(self) -> None:
        """Bind the visual Cancel target to the VM's persisted baseline."""
        gui = self._vm.persisted_config.gui
        self._initial_theme = gui.theme
        self._initial_font = gui.font_size_preset
        self._initial_scale = gui.scale_percent
        self._initial_opacity = gui.transparency_value if gui.transparency_enabled else 1.0

    def _apply_visual_config_as_preview_baseline(self) -> None:
        """Apply persisted visual settings after an immediate reset action."""
        gui = self._vm.persisted_config.gui
        _apply_theme_no_persist(gui.theme)
        from docwen_gui.styles.theme_manager import ThemeManager

        appearance = ThemeManager.get_instance()
        appearance.apply_ui_scale(gui.scale_percent)
        appearance.apply_font_size_preset(gui.font_size_preset)
        opacity = gui.transparency_value if gui.transparency_enabled else 1.0
        _apply_opacity_no_persist(opacity, self.parentWidget())
        self._commit_preview_state()

    @staticmethod
    def _reload_tab_after_reset(tab_key: str, tab: QWidget | None) -> None:
        if tab is None or not hasattr(tab, "reload_from_config"):
            return
        try:
            tab.reload_from_config()  # type: ignore[union-attr]
        except Exception:
            logger.exception("Unexpected Settings tab reload failure: %s", tab_key)

    def _on_cancel(self) -> None:
        """Route the Cancel button through the dialog rejection lifecycle."""

        self.reject()

    def _prepare_cancel_close(self, *, confirm_dirty: bool) -> bool:
        """Confirm and roll back one cancel-style close exactly once.

        ``reject()`` handles Escape, the Cancel button, and programmatic
        rejection. ``closeEvent()`` handles the window-manager close paths
        (title-bar X and Alt+F4).  Both enter here so visual previews and the
        ViewModel draft cannot be left behind by one of those alternate paths.
        """

        if self._close_cleanup_done:
            return True
        if self._cancel_close_in_progress:
            return False

        self._cancel_close_in_progress = True
        try:
            if (
                confirm_dirty
                and self._vm.is_dirty
                and not _show_confirm(
                    self,
                    t("editors.common.confirm_close"),
                    t("editors.common.unsaved_close_message"),
                    danger=True,
                )
            ):
                return False
            self._restore_preview_state()
            if self._vm.is_dirty:
                self._vm.cancel_changes()
            self._close_cleanup_done = True
            return True
        except Exception:
            # A failed rollback must not silently destroy the dialog and its
            # draft.  Keep it open so the user can retry or save instead.
            logger.exception("Unexpected Settings cancel rollback failure")
            return False
        finally:
            self._cancel_close_in_progress = False

    def reject(self) -> None:
        """Reject only after the shared dirty-confirmation rollback succeeds."""

        if self._prepare_cancel_close(confirm_dirty=True):
            super().reject()

    def closeEvent(self, event: QCloseEvent) -> None:
        """Apply the same rollback contract to title-bar and Alt+F4 closes."""

        # A never-shown dialog is normally being disposed after construction
        # or by a test fixture; there was no user close gesture to confirm.
        # It still receives the same rollback before destruction.
        confirm_dirty = self.isVisible()
        if not self._prepare_cancel_close(confirm_dirty=confirm_dirty):
            event.ignore()
            return
        super().closeEvent(event)

    def _on_reset_tab(self) -> None:
        """Reset the current tab to defaults."""
        tab_key = self.current_section()
        if tab_key is None:
            return
        if tab_key == "templates":
            return
        tab_name = TAB_NAMES.get(tab_key, tab_key)

        if not _show_confirm(
            self,
            t("settings.reset.tab_confirm_title", "Reset Current Tab"),
            t(
                "settings.reset.tab_confirm_message",
                "Reset '{tab_name}' settings to their defaults?",
                tab_name=tab_name,
            ),
            danger=True,
        ):
            return

        # Reset is an immediate persistent action after confirmation. The port
        # contract guarantees all-or-nothing persistence for handled failures.
        reset_completed = False
        try:
            try:
                reset_completed = bool(self._vm.reset_group(self._get_reset_group_for_tab(tab_key)))
            except Exception:
                logger.exception("Unexpected Settings tab reset failure: %s", tab_key)
            tab = self._tabs.get(tab_key)
            self._reload_tab_after_reset(tab_key, tab)
            if tab_key == "general" and reset_completed:
                self._apply_visual_config_as_preview_baseline()
        finally:
            self.settings_source_changed.emit()

    def _on_reset_all(self) -> None:
        """Reset configuration drafts; template files and catalog state are independent."""
        if not _show_confirm(
            self,
            t("settings.reset.all_confirm_title", "Reset All Settings"),
            t("settings.reset.all_confirm_message", "Reset ALL settings to their defaults? This cannot be undone."),
            danger=True,
        ):
            return

        reset_completed = False
        try:
            try:
                reset_completed = bool(self._vm.reset_all())
            except Exception:
                logger.exception("Unexpected Settings reset-all failure")
            for tab_key, tab in self._tabs.items():
                self._reload_tab_after_reset(tab_key, tab)
            if reset_completed:
                self._apply_visual_config_as_preview_baseline()
        finally:
            # Consumers reload the persistent source after every attempted reset.
            self.settings_source_changed.emit()

    # ── Status & summary ────────────────────────────────────────────────────

    def _show_status(self, message: str, is_error: bool) -> None:
        """Keep actionable errors visible; auto-hide successful status updates."""
        self._status_label.setText(message)
        self._status_label.setProperty("class", "danger" if is_error else "success")
        self._status_label.style().unpolish(self._status_label)
        self._status_label.style().polish(self._status_label)
        self._status_label.setVisible(True)

        if self._status_timer is None:
            self._status_timer = QTimer(self)
            self._status_timer.setSingleShot(True)
            self._status_timer.timeout.connect(lambda: self._status_label.setVisible(False))
        self._status_timer.stop()
        if not is_error:
            self._status_timer.start(STATUS_DISPLAY_MS)

    def _refresh_changes_summary(self, _dirty: bool | None = None) -> None:
        """Update the changes-summary label with field-level old→new info."""
        if not self._vm.is_dirty:
            self._changes_label.clear()
            self._changes_label.setToolTip("")
            self._changes_label.setVisible(False)
            return

        changes = self._vm.get_change_summary()
        tooltip_lines = [line for change in changes for line in _friendly_change_lines(change)]
        count = len(tooltip_lines)
        self._changes_label.setText(t("settings.changes.summary", "Unsaved changes ({count})", count=count))

        # Keep the visible count and tooltip at user-facing leaf granularity,
        # including dict-backed conversion-default categories.
        visible_lines = tooltip_lines[:10]
        if len(tooltip_lines) > 10:
            visible_lines.append(t("settings.changes.more", "... and {count} more", count=len(tooltip_lines) - 10))
        self._changes_label.setToolTip("\n".join(visible_lines))
        self._changes_label.setVisible(True)

    # ── Tab-section mapping ─────────────────────────────────────────────────

    @staticmethod
    def _get_section_for_tab(tab_key: str) -> str:
        mapping = {
            "general": "gui",
            "text": "text",
            "proofread": "proofread",
            "document": "conversion_defaults",
            "spreadsheet": "conversion_defaults",
            "image": "conversion_defaults",
            "layout": "conversion_defaults",
            "other": "conversion_defaults",
            "link": "link",
            "formatting": "formatting",
            "markdown_input": "formatting",
            "software": "software_priority",
            "output": "output",
            "export": "export",
            "logging": "logging",
        }
        return mapping.get(tab_key, "gui")

    @staticmethod
    def _get_reset_group_for_tab(tab_key: str) -> str:
        return RESET_GROUPS.get(tab_key, "general")

    # ── Public API ──────────────────────────────────────────────────────────

    @property
    def view_model(self) -> SettingsViewModel:
        return self._vm


# ── Theme / opacity no-persist helpers ─────────────────────────────────────


def _apply_theme_no_persist(theme: str) -> None:
    """Apply a theme live without persisting to config."""
    try:
        from docwen_gui.styles.theme_manager import ThemeManager

        ThemeManager.get_instance().apply_theme(theme)
    except Exception:
        pass


def _apply_opacity_no_persist(opacity: float, parent: QWidget | None) -> None:
    """Apply window opacity live without persisting to config."""
    if parent is None:
        return
    with contextlib.suppress(Exception):
        parent.setWindowOpacity(float(opacity))


def _abbreviate_value(val: object) -> str:
    """Truncate long values for display in change-summary tooltips."""
    s = str(val)
    if len(s) > 60:
        return s[:57] + "..."
    return s


def _friendly_change_lines(change: Mapping[str, object]) -> list[str]:
    """Render one VM change as one or more understandable leaf-level lines."""
    field = str(change.get("field", ""))
    old = change.get("old")
    new = change.get("new")
    if isinstance(old, Mapping) and isinstance(new, Mapping):
        lines: list[str] = []
        for key in sorted(set(old) | set(new), key=str):
            old_value = old.get(key)
            new_value = new.get(key)
            if old_value == new_value:
                continue
            child_field = f"{field}.{key}"
            lines.extend(_friendly_change_lines({"field": child_field, "old": old_value, "new": new_value}))
        if lines:
            return lines
    return [
        f"{_friendly_change_field(field)}: {_friendly_change_value(field, old)} → {_friendly_change_value(field, new)}"
    ]


def _friendly_change_field(field: str) -> str:
    """Return a localized field label, with an honest dotted-path fallback."""
    if field.startswith("formatting.markdown_extensions."):
        parts = field.split(".")
        if len(parts) == 4:
            direction, name = parts[2:]
            if direction in ("input", "output") and name in (
                "structural_tables",
                "captions_references",
                "extended_headings",
                "typed_endnotes",
            ):
                return f"{t(f'settings.markdown_extensions.{direction}')} / {t(f'settings.markdown_extensions.{name}')}"
    key = _CHANGE_FIELD_LABEL_KEYS.get(field)
    if key is None:
        return t("settings.changes.field_fallback", "Setting ({field})", field=field)
    return t(key).rstrip(":：")


def _friendly_change_value(field: str, value: object) -> str:
    """Project common settings values to localized labels without hiding unknown data."""
    if isinstance(value, bool):
        return t("settings.changes.enabled") if value else t("settings.changes.disabled")
    if value == "":
        return t("settings.changes.empty_value", "(empty)")
    value_key: str | None = None
    value_labels = _CHANGE_VALUE_LABEL_KEYS.get(field)
    if value_labels is not None:
        try:
            value_key = value_labels.get(value)
        except TypeError:
            value_key = None
    if value_key is not None:
        return t(value_key)
    return _abbreviate_value(value)
