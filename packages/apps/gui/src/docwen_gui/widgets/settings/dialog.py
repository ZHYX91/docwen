"""Settings main dialog.

Replicates the user-visible behavior of the old ``SettingsDialog``:
- 700x800 initial, 510x750 minimum, modal
- FluentNavigationInterface sidebar sized for the active locale + hidden QTabWidget tabBar
- 13 tabs: general, text, proofread, document, spreadsheet, image, layout,
  link, formatting, output, export, logging, other
- Bottom bar: Reset Tab + Reset All + Ok/Cancel/Apply (QDialogButtonBox)
- Dirty tracking with auto-signal monitoring
- Unsaved-close confirmation (danger, default=no)
- Change summary label (field-level, max 10 lines in tooltip)
- Status message with auto-hide timer
"""

from __future__ import annotations

import contextlib
import logging
from collections.abc import Callable
from typing import Any, NamedTuple
from typing import cast as _cast

from PySide6.QtCore import QSignalBlocker, Qt, QTimer, Signal
from PySide6.QtGui import QCloseEvent, QFontMetrics, QIcon
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

from ...styles.theme_semantics import apply_theme_class
from ...view_models.settings_vm import SettingsViewModel

logger = logging.getLogger(__name__)

# ── Dialog geometry ────────────────────────────────────────────────────────
DEFAULT_WIDTH = 700
DEFAULT_HEIGHT = 800
MIN_WIDTH = 510
MIN_HEIGHT = 750
DIALOG_PADDING = 15
ACTION_BUTTON_MIN_HEIGHT = 32
RESET_TAB_BUTTON_MIN_WIDTH = 116
RESET_ALL_BUTTON_MIN_WIDTH = 104
STATUS_DISPLAY_MS = 3000
NAVIGATION_MIN_WIDTH = 168
NAVIGATION_TEXT_CHROME_WIDTH = 72


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
    """Read the current active theme before the dialog opens."""
    try:
        from docwen_gui.styles.theme_manager import ThemeManager

        return ThemeManager.get_instance().get_current_theme()
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
# Keep one explicit navigation order so the sidebar, page stack, tests, and
# documentation expose the same settings surface.
TAB_KEYS = [
    "general",
    "text",
    "proofread",
    "document",
    "spreadsheet",
    "image",
    "layout",
    "link",
    "formatting",
    "output",
    "export",
    "logging",
    "other",
]

# ── Tab display names ──────────────────────────────────────────────────────
TAB_NAMES: dict[str, str] = {
    "general": t("settings.tabs.general"),
    "text": t("settings.tabs.text"),
    "proofread": t("settings.tabs.proofread"),
    "document": t("settings.tabs.document"),
    "spreadsheet": t("settings.tabs.spreadsheet"),
    "image": t("settings.tabs.image"),
    "layout": t("settings.tabs.layout"),
    "link": t("settings.tabs.link"),
    "formatting": t("settings.tabs.formatting"),
    "output": t("settings.tabs.output"),
    "export": t("settings.tabs.export"),
    "logging": t("settings.tabs.logging"),
    "other": t("settings.tabs.other"),
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
    "output": _TabSpec("output_tab", "OutputTab", _build_output_tab),
    "export": _TabSpec("export_tab", "ExportTab", _build_export_tab),
    "logging": _TabSpec("logging_tab", "LoggingTab", _build_logging_tab),
    "other": _TabSpec("other_tab", "OtherTab", _build_other_tab),
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
    "output": "output",
    "export": "export",
    "logging": "logging",
    "other": "other",
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
    mb.setDefaultButton(QMessageBox.StandardButton.No)
    return mb.exec() == QMessageBox.StandardButton.Yes


def _try_fluent_panel(navigation: Any, key: str) -> Any | None:
    """Safely access the FluentNavigationInterface panel widget for a key."""
    panel = getattr(navigation, "panel", None)
    if panel is None or not hasattr(panel, "widget"):
        return None
    try:
        return panel.widget(key)
    except Exception:
        return None


class SettingsDialog(QDialog):
    """Settings dialog with 13-tab navigation and dirty tracking.

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
    ) -> None:
        title = t("settings.title")
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)

        self._vm = view_model or SettingsViewModel()
        # Capture the persisted baseline so Cancel can roll back all edits
        self._vm.begin_session()
        # Save initial visual state for Cancel rollback
        self._initial_theme = _read_initial_theme()
        self._initial_opacity = _read_initial_opacity(parent)
        self._tabs: dict[str, QWidget] = {}
        self._tab_widget: QTabWidget = _cast(QTabWidget, None)
        self._navigation: Any = None
        self._status_timer: QTimer = _cast(QTimer, None)
        self._cancel_close_in_progress = False
        self._close_cleanup_done = False

        self._build_ui()
        self._wire_view_model()

    # ── UI construction ─────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        self.setObjectName("settingsDialog")
        self.resize(DEFAULT_WIDTH, DEFAULT_HEIGHT)
        self.setMinimumSize(MIN_WIDTH, MIN_HEIGHT)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(DIALOG_PADDING, DIALOG_PADDING, DIALOG_PADDING, DIALOG_PADDING)
        layout.setSpacing(DIALOG_PADDING)

        # ── Tab widget (hidden tab bar) ─────────────────────────────────
        self._tab_widget = QTabWidget(self)
        self._tab_widget.setDocumentMode(True)
        self._tab_widget.setUsesScrollButtons(False)
        self._tab_widget.setElideMode(Qt.TextElideMode.ElideRight)
        self._tab_widget.tabBar().hide()

        # ── Sidebar + content row ───────────────────────────────────────
        content_row = QHBoxLayout()
        content_row.setContentsMargins(0, 0, 0, 0)
        content_row.setSpacing(8)

        self._build_navigation(content_row)
        content_row.addWidget(self._tab_widget, 1)
        layout.addLayout(content_row, 1)

        # ── Tabs ────────────────────────────────────────────────────────
        self._add_all_tabs()

        # ── Inject template data into the ViewModel ────────────────────
        self._load_templates_into_vm()

        # ── Populate dynamic optimization type items (ViewModel-driven) ─
        self._populate_optimization_types()

        # ── Activate the optimal tab (if specified by caller) ───────────
        self._activate_initial_tab()

        # ── Bottom action bar ───────────────────────────────────────────
        action_row = QHBoxLayout()
        action_row.setContentsMargins(0, 4, 0, 0)
        action_row.setSpacing(8)

        reset_tab_btn = QPushButton(t("settings.reset.tab_button"), self)
        reset_tab_btn.setObjectName("settingsResetTabButton")
        reset_tab_btn.setMinimumHeight(ACTION_BUTTON_MIN_HEIGHT)
        reset_tab_btn.setMinimumWidth(RESET_TAB_BUTTON_MIN_WIDTH)
        apply_theme_class(reset_tab_btn, "secondary")
        reset_tab_btn.clicked.connect(self._on_reset_tab)

        reset_all_btn = QPushButton(t("settings.reset.all_button"), self)
        reset_all_btn.setObjectName("settingsResetAllButton")
        reset_all_btn.setMinimumHeight(ACTION_BUTTON_MIN_HEIGHT)
        reset_all_btn.setMinimumWidth(RESET_ALL_BUTTON_MIN_WIDTH)
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
            ok_btn.setMinimumHeight(ACTION_BUTTON_MIN_HEIGHT)
            apply_theme_class(ok_btn, "primary")
        cancel_btn = button_box.button(QDialogButtonBox.StandardButton.Cancel)
        if cancel_btn:
            cancel_btn.setText(t("common.cancel"))
            cancel_btn.setObjectName("settingsCancelButton")
            cancel_btn.setMinimumHeight(ACTION_BUTTON_MIN_HEIGHT)
            apply_theme_class(cancel_btn, "secondary")
        apply_btn = button_box.button(QDialogButtonBox.StandardButton.Apply)
        if apply_btn:
            apply_btn.setText(t("common.apply"))
            apply_btn.setObjectName("settingsApplyButton")
            apply_btn.setMinimumHeight(ACTION_BUTTON_MIN_HEIGHT)
            apply_theme_class(apply_btn, "secondary")

        button_box.accepted.connect(self._on_ok)
        button_box.rejected.connect(self._on_cancel)
        if apply_btn:
            apply_btn.clicked.connect(self._on_apply)

        action_row.addWidget(button_box)
        layout.addLayout(action_row)

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

        self._tab_widget.currentChanged.connect(self._on_tab_changed)

    # ── Preview restore ────────────────────────────────────────────────────

    def _restore_preview_state(self) -> None:
        """Restore theme and opacity to the baseline captured at dialog open
        or by the most recent successful Apply operation."""
        if self._initial_theme is not None:
            _apply_theme_no_persist(self._initial_theme)
        if self._initial_opacity is not None:
            _apply_opacity_no_persist(self._initial_opacity, self.parentWidget())

    def _build_navigation(self, content_row: QHBoxLayout) -> None:
        """Build the FluentNavigationInterface sidebar (or fallback)."""
        try:
            from qfluentwidgets import NavigationInterface, NavigationItemPosition

            nav = NavigationInterface(self, showMenuButton=False, showReturnButton=False, collapsible=True)
            nav.setObjectName("settingsFluentNavigation")
            text_width = max(QFontMetrics(nav.font()).horizontalAdvance(title) for title in TAB_NAMES.values())
            navigation_width = max(NAVIGATION_MIN_WIDTH, text_width + NAVIGATION_TEXT_CHROME_WIDTH)
            nav.setExpandWidth(navigation_width)
            nav.setMinimumExpandWidth(0)
            nav.setCollapsible(False)
            nav.setFixedWidth(navigation_width)
            self._navigation = nav
            self._nav_position = NavigationItemPosition.TOP
            content_row.addWidget(nav, 0)
        except Exception:
            logger.debug("qfluentwidgets NavigationInterface not available — using tab bar fallback")
            self._navigation = None
            self._activate_tab_bar_fallback()

    def _activate_tab_bar_fallback(self) -> None:
        """Expose a stable, scrollable navigation surface for every page."""
        navigation = self._navigation
        if navigation is not None:
            with contextlib.suppress(Exception):
                navigation.hide()
        self._navigation = None
        self._tab_widget.setUsesScrollButtons(True)
        self._tab_widget.tabBar().show()

    def _add_nav_item(self, key: str, index: int, title: str) -> None:
        if self._navigation is None:
            return
        try:
            from PySide6.QtWidgets import QStyle

            # Prefer the tab's dedicated SVG before the platform fallback.
            icon = self._load_tab_icon(key)
            if icon is None:
                icon = self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogDetailedView)

            self._navigation.addItem(
                key,
                icon,
                title,
                onClick=lambda _checked=False, idx=index: self._tab_widget.setCurrentIndex(idx),
                position=self._nav_position,
                tooltip=title,
            )
            if index == 0:
                self._navigation.setCurrentItem(key)
        except Exception as exc:
            logger.debug("Failed to add nav item %s: %s", key, exc)
            self._activate_tab_bar_fallback()

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
            "output": "output.svg",
            "export": "export.svg",
            "logging": "logging.svg",
            "other": "other.svg",
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
        item = _try_fluent_panel(self._navigation, key)
        if item is None:
            return
        if hasattr(item, "setText"):
            item.setText(title)
        if hasattr(item, "setToolTip"):
            item.setToolTip(title)

    # ── Tab management ──────────────────────────────────────────────────────

    def _add_all_tabs(self) -> None:
        for idx, key in enumerate(TAB_KEYS):
            title = TAB_NAMES.get(key, key)
            try:
                tab = self._build_tab(key)
            except Exception as error:
                logger.exception("Failed to build Settings tab: %s", key)
                tab = self._build_failed_tab(key, title, error)
            self._tabs[key] = tab
            set_tab_title = getattr(tab, "set_tab_title", None)
            if callable(set_tab_title):
                set_tab_title(title)
            tab_index = self._tab_widget.addTab(tab, title)
            self._tab_widget.setTabToolTip(tab_index, title)
            self._add_nav_item(key, idx, title)
            self._wire_tab_dirty_tracking(key, tab)

        # Select first tab
        self._tab_widget.setCurrentIndex(0)

    def _build_tab(self, key: str) -> QWidget:
        """Construct one settings page through a PyInstaller-visible lazy import."""
        spec = _TAB_SPECS.get(key)
        if spec is None:
            raise KeyError(f"Unknown Settings tab: {key!r}")
        tab = spec.factory(self._vm)
        if not isinstance(tab, QWidget):
            raise TypeError(f"Settings tab factory {key!r} did not return a QWidget")
        return tab

    def _build_failed_tab(self, key: str, title: str, error: Exception) -> QWidget:
        """Return a stable placeholder when one settings page cannot be built."""
        tab = QWidget(self._tab_widget)
        tab.setObjectName("settingsTabLoadErrorPage")
        tab.setProperty("failedTabKey", key)

        layout = QVBoxLayout(tab)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

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

    def _wire_tab_dirty_tracking(self, tab_key: str, tab: QWidget) -> None:
        """ViewModel handles dirty tracking internally via set_field().

        This method is kept for potential future per-tab UI hooks (e.g.
        adding a "*" to the tab title) but no longer duplicates dirty-state
        logic — the ViewModel is the single source of truth.
        """

    def _load_templates_into_vm(self) -> None:
        """Query TemplateRegistry and inject template lists into the ViewModel.

        Called during UI construction so TextTab's TabbedTemplateSelector
        is populated on dialog open.  Falls back silently if the template
        registry is unavailable (no templates found, templates dir missing,
        or runtime package not installed).
        """
        try:
            from docwen_runtime.templates import TemplateRegistry
        except ImportError:
            logger.debug("TemplateRegistry not available — template selector stays empty")
            return
        try:
            registry = TemplateRegistry.default()
            templates: dict[str, list[str]] = {"docx": [], "xlsx": []}
            for info in registry.list_templates():
                target = info.target
                if target in templates:
                    templates[target].append(info.name)
            self._vm.set_templates(templates)
        except Exception:
            logger.debug("Failed to load templates from registry", exc_info=True)

    def _populate_optimization_types(self) -> None:
        """Populate optimization policy controls from Runtime capabilities.

        A valid empty catalog leaves an empty disabled combo.  Discovery
        failure instead shows a disabled error item, while the remainder of
        the settings page stays usable.
        """
        from .base_tab import BaseSettingsTab

        config = self._vm.config.conversion_defaults
        for tab_key, source_category in (
            ("document", "document"),
            ("image", "image"),
            ("layout", "layout"),
        ):
            tab = self._tabs.get(tab_key)
            widgets = getattr(tab, "widgets", {}) if tab is not None else {}
            if not isinstance(widgets, dict):
                continue
            combo = widgets.get("to_md_optimization_type")
            enable_checkbox = widgets.get("to_md_enable_optimization")
            if not isinstance(combo, QComboBox):
                continue

            result = self._vm.get_optimization_choices_result(source_category=source_category)
            configured = getattr(config, tab_key, {})
            configured_id = configured.get("to_md_optimization_type") if isinstance(configured, dict) else None
            blocker = QSignalBlocker(combo)
            combo.clear()
            if result.status == "failed":
                combo.addItem(
                    t("main_window.runtime_unavailable", "Runtime is unavailable; conversion cannot start."),
                    None,
                )
                error_text = str(result.error or "")
                combo.setToolTip(error_text)
                if isinstance(enable_checkbox, QCheckBox):
                    enable_checkbox.setToolTip(error_text)
            else:
                combo.setToolTip("")
                for choice in result.choices:
                    combo.addItem(choice.label, choice.id)
                if configured_id is not None:
                    BaseSettingsTab.set_combo_data(combo, configured_id)
            del blocker

            has_choices = bool(result.choices)
            combo.setEnabled(
                has_choices and (not isinstance(enable_checkbox, QCheckBox) or enable_checkbox.isChecked())
            )
            if isinstance(enable_checkbox, QCheckBox):
                enable_checkbox.setEnabled(has_choices)
                enable_checkbox.toggled.connect(
                    lambda checked, target=combo, available=has_choices: target.setEnabled(available and checked)
                )

    def _activate_initial_tab(self) -> None:
        """Activate the tab specified by the ViewModel's ``initial_tab_key``.

        If no initial tab was set this is a no-op — the dialog stays on
        the first (general) tab.
        """
        initial = self._vm.initial_tab_key
        if initial and initial in TAB_KEYS:
            idx = TAB_KEYS.index(initial)
            if self._tab_widget is not None:
                self._tab_widget.setCurrentIndex(idx)
                if self._navigation is not None:
                    with contextlib.suppress(Exception):
                        self._navigation.setCurrentItem(initial)

    def activate_section(self, section: str) -> bool:
        """Select a public settings section, rejecting failed page placeholders."""

        if section not in TAB_KEYS:
            return False
        tab = self._tabs.get(section)
        if tab is None or tab.property("failedTabKey") == section:
            return False
        self._tab_widget.setCurrentWidget(tab)
        if self._navigation is not None:
            with contextlib.suppress(Exception):
                self._navigation.setCurrentItem(section)
        return True

    def current_section(self) -> str | None:
        """Return the semantic key of the currently selected settings page."""

        index = self._tab_widget.currentIndex()
        return TAB_KEYS[index] if 0 <= index < len(TAB_KEYS) else None

    def _wire_view_model(self) -> None:
        """Connect ViewModel signals."""
        self._vm.dirty_state_changed.connect(self._refresh_changes_summary)
        self._vm.status_changed.connect(self._show_status)

    # ── Tab changed ─────────────────────────────────────────────────────────

    def _on_tab_changed(self, index: int) -> None:
        """Sync navigation selection and set focus on first focusable widget."""
        if 0 <= index < len(TAB_KEYS):
            key = TAB_KEYS[index]
            if self._navigation is not None:
                with contextlib.suppress(Exception):
                    self._navigation.setCurrentItem(key)

        # Auto-focus first eligible widget in the new tab.
        tab = self._tab_widget.currentWidget()
        if tab is None:
            return
        candidate = self._find_focus_target(tab)
        if candidate is not None:
            candidate.setFocus(Qt.FocusReason.TabFocusReason)

    @staticmethod
    def _find_focus_target(tab: QWidget) -> QWidget | None:
        """Return the first eligible control, preserving the old Qt focus UX."""
        for widget in tab.findChildren(QWidget):
            if not widget.isEnabled() or widget.focusPolicy() == Qt.FocusPolicy.NoFocus:
                continue
            if isinstance(widget, (QAbstractButton, QCheckBox, QComboBox, QLineEdit, QSpinBox, QDoubleSpinBox)):
                return widget
            if isinstance(widget, QScrollArea):
                continue
        return tab if tab.focusPolicy() != Qt.FocusPolicy.NoFocus else None

    # ── Action handlers ─────────────────────────────────────────────────────

    def _on_ok(self) -> None:
        """Apply settings and close."""
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
        self._initial_opacity = gui.transparency_value if gui.transparency_enabled else 1.0

    def _apply_visual_config_as_preview_baseline(self) -> None:
        """Apply persisted visual settings after an immediate reset action."""
        gui = self._vm.persisted_config.gui
        _apply_theme_no_persist(gui.theme)
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
                    t("common.error"),
                    t("settings.unsaved_changes.message", "You have unsaved changes. Close without saving?"),
                    danger=True,
                )
            ):
                return False
            self._restore_preview_state()
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
        index = self._tab_widget.currentIndex()
        if not (0 <= index < len(TAB_KEYS)):
            return
        tab_key = TAB_KEYS[index]
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
        """Reset all tabs to defaults."""
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
        """Display a status message with auto-hide."""
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
        self._status_timer.start(STATUS_DISPLAY_MS)

    def _refresh_changes_summary(self, _dirty: bool | None = None) -> None:
        """Update the changes-summary label with field-level old→new info."""
        if not self._vm.is_dirty:
            self._changes_label.clear()
            self._changes_label.setToolTip("")
            self._changes_label.setVisible(False)
            return

        changes = self._vm.get_change_summary()
        count = len(changes)
        self._changes_label.setText(t("settings.changes.summary", "Unsaved changes ({count})", count=count))

        # Build tooltip as field-level old → new (max 10 lines)
        tooltip_lines: list[str] = []
        for c in changes[:10]:
            old_str = _abbreviate_value(c.get("old"))
            new_str = _abbreviate_value(c.get("new"))
            tooltip_lines.append(f"{c['field']}: {old_str} → {new_str}")
        if len(changes) > 10:
            tooltip_lines.append(t("settings.changes.more", "... and {count} more", count=len(changes) - 10))
        self._changes_label.setToolTip("\n".join(tooltip_lines))
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
