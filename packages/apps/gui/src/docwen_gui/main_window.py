"""Main window and composition root for the DocWen GUI.

- sub-widgets are instantiated explicitly here;
- runtime work is delegated through ``ApplicationController``;
- runtime events cross threads only through ``TaskEventBridge``.
"""

from __future__ import annotations

import contextlib
import logging
import os
import time
import uuid
from collections.abc import Callable, Sequence
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, ClassVar, Literal
from typing import cast as _cast

from PySide6.QtCore import QEvent, QObject, QPoint, QRect, QSize, Qt, QThread, QTimer, Signal, Slot
from PySide6.QtGui import (
    QAction,
    QActionGroup,
    QCloseEvent,
    QKeySequence,
    QMoveEvent,
    QResizeEvent,
    QShortcut,
    QShowEvent,
)
from PySide6.QtWidgets import (
    QAbstractSpinBox,
    QApplication,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QPlainTextEdit,
    QStackedWidget,
    QSystemTrayIcon,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from docwen_gui import path_actions
from docwen_gui.file_admission_i18n import render_file_inspection_message
from docwen_gui.font_utils import FONT_SIZE_PRESETS, normalize_font_size_preset
from docwen_gui.i18n import t as _t
from docwen_gui.resources import load_svg_icon
from docwen_gui.styles.theme_manager import ThemeManager
from docwen_gui.view_models._optimization_filter import OptimizationSource
from docwen_gui.view_models._runtime_route_filter import (
    RuntimeRouteChoice,
    RuntimeRouteSource,
    discover_runtime_route_choices,
)
from docwen_gui.window_behavior import (
    DEFAULT_WINDOW_BEHAVIOR,
    PanelWidthPlan,
    WindowBehaviorPolicy,
    load_window_behavior_policy,
    plan_panel_widths,
)
from docwen_gui.window_geometry import (
    DEFAULT_MIN_HEIGHT,
    DEFAULT_MIN_WIDTH,
    DEFAULT_WINDOW_HEIGHT,
    RecoveredWindowGeometry,
    ScreenRect,
    WindowGeometryPolicy,
    WindowRect,
    build_canonical_geometry_values,
    center_window_geometry,
    load_window_geometry_policy,
    load_window_scale_factor,
    recover_window_geometry,
)
from docwen_runtime.path_io import filesystem_path

if TYPE_CHECKING:
    from docwen_application.controller import ApplicationController
    from docwen_core.models import FileInspection
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy
    from docwen_core.models.result import ConversionResult
    from docwen_gui.qt_bridge.task_event_bridge import TaskEventBridge

    from .view_models.action_area_vm import ActionAreaViewModel
    from .view_models.batch_list_vm import BatchListViewModel
    from .view_models.conversion_panel_vm import ConversionPanelViewModel
    from .view_models.info_area_vm import InfoAreaViewModel
    from .view_models.input_area_vm import InputAreaViewModel
    from .view_models.interaction import MainWindowUiProjection
    from .view_models.main_window_vm import MainWindowViewModel
    from .widgets.action_area import ActionArea
    from .widgets.batch_list import BatchList
    from .widgets.conversion_panel import ConversionPanel
    from .widgets.info_area import InfoArea
    from .widgets.input_area import InputArea
    from .widgets.template_selector_tabbed import TabbedTemplateSelector

# ── Window geometry constants ──────────────────────────────────────────
DEFAULT_CENTER_WIDTH = DEFAULT_WINDOW_BEHAVIOR.center_panel_width
DEFAULT_HEIGHT = DEFAULT_WINDOW_HEIGHT
MIN_WIDTH = DEFAULT_MIN_WIDTH
MIN_HEIGHT = DEFAULT_MIN_HEIGHT

_MARKDOWN_TARGET_FORMATS: frozenset[str] = frozenset({"md", "markdown"})
_DOCUMENT_TEMPLATE_TARGETS: frozenset[str] = frozenset({"docx", "doc", "odt", "rtf", "wps", "pdf"})
_SPREADSHEET_TEMPLATE_TARGETS: frozenset[str] = frozenset({"xlsx", "xls", "ods", "csv"})
_AGGREGATE_ACTIONS: frozenset[str] = frozenset({"merge_pdfs", "merge_tables", "merge_images_to_tiff"})
_PROOFREAD_ACTIONS: frozenset[str] = frozenset({"validate"})
_ConversionRequestOrigin = Literal["action_area", "conversion_panel"]
_PROOFREAD_GUI_OPTION_ALIASES: dict[str, str] = {
    "symbol_pairing": "enable_symbol_pairing",
    "symbol_correction": "enable_symbol_correction",
    "typos_rule": "enable_typos_rule",
    "sensitive_word": "enable_sensitive_word",
}

logger = logging.getLogger(__name__)


class _OutputPolicyConfigError(RuntimeError):
    """Raised when persisted output settings cannot be read safely."""


def _result_warning_messages(result: ConversionResult) -> list[str]:
    """Return user-visible warning diagnostics from a successful result."""
    messages: list[str] = []
    for diagnostic in result.diagnostics:
        if diagnostic.level != "warning":
            continue
        message = diagnostic.message.strip() or diagnostic.code.strip()
        if not message:
            message = _t("main_window.conversion_warning", "Conversion completed with a warning")
        if diagnostic.location:
            message = f"{message} ({diagnostic.location})"
        messages.append(message)
    return messages


def _redacted_request_options(options: dict[str, Any]) -> dict[str, Any]:
    """Keep execution secrets out of GUI retry/history context."""

    redacted = dict(options)
    if "spreadsheet_password" in redacted:
        redacted["spreadsheet_password"] = "<redacted>"
    return redacted


def _to_markdown_locale_options(
    options: dict[str, Any],
    *,
    target_format: str,
    action_name: str = "",
    route_options: Sequence[str] | None = None,
) -> dict[str, Any]:
    from docwen_gui.i18n import get_locale

    enriched = dict(options)
    if route_options is not None:
        supported = frozenset(route_options)
        if "locale" in supported:
            enriched.setdefault("locale", get_locale())
        if target_format not in _MARKDOWN_TARGET_FORMATS or "yaml_key_labels" not in supported:
            return enriched
    elif target_format not in _MARKDOWN_TARGET_FORMATS:
        return enriched
    elif action_name:
        # Named routes must provide their canonical option surface. Do not
        # infer support from an action string.
        return enriched
    else:
        enriched.setdefault("locale", get_locale())
    enriched.setdefault(
        "yaml_key_labels",
        {
            "title": _t("yaml_keys.title", default="title"),
            "subtitle": _t("yaml_keys.subtitle", default="subtitle"),
        },
    )
    return enriched


def _normalize_proofread_action_options(options: dict[str, Any], *, action_name: str) -> dict[str, Any]:
    if action_name not in _PROOFREAD_ACTIONS:
        return options
    normalized = dict(options)
    for gui_key, plugin_key in _PROOFREAD_GUI_OPTION_ALIASES.items():
        if gui_key not in normalized:
            continue
        value = normalized.pop(gui_key)
        normalized.setdefault(plugin_key, bool(value))
    return normalized


def _route_scoped_options(
    options: dict[str, Any],
    *,
    route_options: Sequence[str] | None,
) -> dict[str, Any]:
    if route_options is None:
        return dict(options)
    supported = frozenset(route_options)
    return {key: value for key, value in options.items() if key in supported}


_OUTPUT_DATE_SUBFOLDER_TOKENS: dict[str, str] = {
    "%Y-%m-%d": "iso",
    "%Y%m%d": "compact",
    "%Y年%m月%d日": "chinese",
}


def _resolve_file_context(
    file_contexts: dict[str, tuple[str, str]],
    batch_list_vm: BatchListViewModel,
    file_path: str,
) -> tuple[str, str] | None:
    """Resolve file format/category from file contexts or batch list VM."""
    normalized = _normalize_path(file_path)
    context = file_contexts.get(normalized)
    if context is not None:
        return context
    entry = batch_list_vm.get_file_entry(file_path)
    if entry is not None:
        return entry.detected_format.lower(), entry.workflow_category.lower()
    return None


def _output_date_subfolder_token(date_folder_format: str) -> str:
    """Map GUI output date formats to runtime output-policy tokens."""
    return _OUTPUT_DATE_SUBFOLDER_TOKENS.get(date_folder_format, date_folder_format)


def _detected_dpi_scale() -> float:
    """Return the primary screen's bounded logical-DPI factor."""
    try:
        app = QApplication.instance()
        if isinstance(app, QApplication):
            screen = app.primaryScreen()
            if screen is not None:
                dpi = screen.logicalDotsPerInch()
                factor = float(dpi) / 96.0
                if 0.5 <= factor <= 4.0:
                    return factor
    except (AttributeError, RuntimeError):
        pass
    return 1.0


def _normalize_path(file_path: str) -> str:
    """Normalize a path for internal lookup across widgets/view-models."""
    return str(Path(file_path)).replace("\\", "/")


def _format_template_modified_label(modified_ns: object) -> str | None:
    """Format TemplateRegistry's nanosecond mtime for compact UI metadata."""
    if isinstance(modified_ns, bool) or not isinstance(modified_ns, int):
        return None
    try:
        timestamp = modified_ns / 1_000_000_000
        return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")
    except (OSError, OverflowError, ValueError):
        return None


def _move_path_to_front(file_paths: Sequence[str], preferred_path: str | None) -> list[str]:
    """Return file paths with the preferred path first when present."""
    ordered = list(file_paths)
    if not preferred_path:
        return ordered
    preferred_key = _normalize_path(preferred_path)
    for index, path in enumerate(ordered):
        if _normalize_path(path) != preferred_key:
            continue
        if index == 0:
            return ordered
        return [ordered[index], *ordered[:index], *ordered[index + 1 :]]
    return ordered


def _read_pdf_total_pages(file_path: str) -> int | None:
    """Return PDF page count when it can be read without disrupting the GUI."""
    try:
        import fitz

        with fitz.open(file_path, filetype="pdf") as doc:
            page_count = int(doc.page_count)
    except Exception:
        return None
    return page_count if page_count > 0 else None


class _ExecutionThread(QThread):
    """Run a single conversion request off the UI thread.

    Uses direct QThread.run() override instead of moveToThread +
    thread.started.connect to avoid signal-delivery deadlock in the
    MainWindow context.
    """

    result_signal = Signal(object, dict)
    error_signal = Signal(str, dict)

    def __init__(
        self,
        *,
        controller: ApplicationController,
        request: ConversionRequest,
        context: dict[str, Any],
        aggregate_action_name: str = "",
        batch_execution: bool = False,
        parent: QObject | None = None,
    ) -> None:
        super().__init__(parent)
        self._controller = controller
        self._request = request
        self._context = context
        self._aggregate_action_name = aggregate_action_name
        self._batch_execution = batch_execution

    def run(self) -> None:
        try:
            if self._aggregate_action_name:
                result = self._controller.execute_aggregate(self._request, self._aggregate_action_name)
            elif self._batch_execution:
                result = self._controller.execute_batch(self._request)
            else:
                result = self._controller.execute_single(self._request)
            self.result_signal.emit(result, self._context)
        except Exception as exc:
            self.error_signal.emit(str(exc), self._context)


class MainWindow(QWidget):
    """Top-level main window for the new GUI."""

    _EXECUTION_DRAIN_POLL_MS = 25
    _EXECUTION_DRAIN_TIMEOUT_SECONDS = 5.0

    def __init__(
        self,
        view_model: MainWindowViewModel,
        task_event_bridge: TaskEventBridge | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._view_model = view_model
        self._window_behavior = self._load_window_behavior()
        self._window_scale_factor = self._load_window_scale_factor()
        self._CENTER_PANEL_MIN_WIDTH: int = self._scale_window_value(self._window_behavior.center_panel_width)
        self._LEFT_PANEL_MIN_WIDTH: int = self._scale_window_value(self._window_behavior.left_panel_width)
        self._RIGHT_PANEL_MIN_WIDTH: int = self._scale_window_value(self._window_behavior.right_panel_width)
        persisted_window_geometry = self._load_window_geometry(center_offset=0)
        self._persisted_window_geometry_policy = persisted_window_geometry
        if not self._window_behavior.remember_gui_state:
            default_window_geometry = load_window_geometry_policy(
                None,
                center_offset=0,
                scale_value=self._scale_window_value,
            )
            self._window_geometry = WindowGeometryPolicy(
                rect=default_window_geometry.rect,
                min_width=persisted_window_geometry.min_width,
                min_height=persisted_window_geometry.min_height,
                source="startup-default",
                schema_version=persisted_window_geometry.schema_version,
                schema_supported=persisted_window_geometry.schema_supported,
            )
        else:
            self._window_geometry = persisted_window_geometry
        self._geometry_source_change_rect: WindowRect | None = None
        self._geometry_source_transition_rebase_pending = False
        self._geometry_source_transition_settling = False
        self._applying_internal_panel_geometry = False
        self._prepared_file_clear_normal_rect: WindowRect | None = None
        self._prepared_file_clear_center_x: int | None = None
        self._file_clear_updates_suspended = False
        self._last_normal_window_rect: WindowRect | None = None
        self._last_normal_center_offset = 0
        self._last_normal_left_visible = False
        self._last_normal_right_visible = False
        self._collapsed_normal_window_rect: WindowRect | None = None
        self._collapsed_normal_center_offset = 0
        self._window_state_needs_shown_restore = True
        self._shown_window_restore_finalized = False
        self._pending_normal_center_anchor: int | None = None
        self._pending_anchor_restore_scheduled = False
        self._pending_normal_frame_restored = False
        if task_event_bridge is None:
            from .qt_bridge.task_event_bridge import TaskEventBridge as _TaskEventBridge

            task_event_bridge = _TaskEventBridge(self)
        self._task_event_bridge = task_event_bridge
        self._task_event_bridge.setParent(self)

        self._input_area_vm: InputAreaViewModel = _cast("InputAreaViewModel", None)
        self._batch_list_vm: BatchListViewModel = _cast("BatchListViewModel", None)
        self._conversion_panel_vm: ConversionPanelViewModel = _cast("ConversionPanelViewModel", None)
        self._action_area_vm: ActionAreaViewModel = _cast("ActionAreaViewModel", None)
        self._info_area_vm: InfoAreaViewModel = _cast("InfoAreaViewModel", None)

        self._input_area: InputArea = _cast("InputArea", None)
        self._batch_list: BatchList = _cast("BatchList", None)
        self._conversion_panel: ConversionPanel = _cast("ConversionPanel", None)
        self._action_area: ActionArea = _cast("ActionArea", None)
        self._info_area: InfoArea = _cast("InfoArea", None)
        self._template_selector: TabbedTemplateSelector | None = None
        self._main_template_default_type: str | None = None
        self._left_panel_frame: QFrame = _cast("QFrame", None)
        self._right_panel_frame: QFrame = _cast("QFrame", None)
        self._right_stack: QStackedWidget = _cast("QStackedWidget", None)
        self._center_column: QWidget = _cast("QWidget", None)

        self._active_threads: dict[str, QThread] = {}
        self._execution_cleanup_by_thread: dict[QThread, tuple[str, ApplicationController, object]] = {}
        self._execution_close_pending = False
        self._execution_drain_timed_out = False
        self._execution_drain_deadline = 0.0
        self._execution_drain_timer: QTimer | None = None
        self._shutdown_finalized = False
        self._last_request_contexts: dict[str, dict[str, Any]] = {}
        self._current_mode = view_model.mode
        self._file_contexts: dict[str, tuple[str, str]] = {}
        self._always_on_top_enabled = False
        self._font_size_preset: str = "default"
        self._system_tray_icon: QSystemTrayIcon | None = None
        self._start_time: float | None = None
        self._settings_dialog: Any | None = None

        self._setup_window_properties()
        self.setup_ui()
        # Post-condition: all component and VM fields are non-None after setup_ui().
        assert self._input_area is not None
        assert self._batch_list is not None
        assert self._conversion_panel is not None
        assert self._action_area is not None
        assert self._info_area is not None
        assert self._input_area_vm is not None
        assert self._batch_list_vm is not None
        assert self._conversion_panel_vm is not None
        assert self._action_area_vm is not None
        assert self._info_area_vm is not None

    # ── Window properties ──────────────────────────────────────────

    def _load_window_behavior(self) -> WindowBehaviorPolicy:
        controller = self._view_model.controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        return load_window_behavior_policy(cfg_port)

    def _load_window_scale_factor(self) -> float:
        controller = self._view_model.controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        return load_window_scale_factor(
            cfg_port,
            detected_factor=_detected_dpi_scale(),
        )

    def _scale_window_value(self, value: int) -> int:
        return round(int(value) * self._window_scale_factor)

    def _unscale_window_value(self, value: int) -> int:
        return round(int(value) / self._window_scale_factor)

    def _load_window_geometry(self, *, center_offset: int) -> WindowGeometryPolicy:
        controller = self._view_model.controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        return load_window_geometry_policy(
            cfg_port,
            center_offset=center_offset,
            scale_value=self._scale_window_value,
        )

    def _setup_window_properties(self) -> None:
        self.setWindowTitle(_t("main_window.window_title", "DocWen Offline"))
        self.setObjectName("docwenMainWindow")

        geometry = self._window_geometry
        self.setMinimumSize(QSize(geometry.min_width, geometry.min_height))
        self.resize(
            max(geometry.rect.width, geometry.min_width),
            max(geometry.rect.height, geometry.min_height),
        )

    # ── UI construction ────────────────────────────────────────────

    def setup_ui(self) -> None:
        """Build the full widget tree.  Idempotent — safe to call more than once."""
        if getattr(self, "_ui_built", False):
            return
        self._ui_built = True

        self._create_actions()

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self._create_components()

        central_container = QWidget()
        central_container.setObjectName("centralContainer")
        grid = QGridLayout(central_container)
        grid.setContentsMargins(8, 8, 8, 8)
        grid.setSpacing(8)

        # ── Left panel: batch list (visible in batch mode) ──────────────
        self._left_panel_frame = QFrame()
        self._left_panel_frame.setObjectName("leftPanelFrame")
        self._left_panel_frame.setMinimumWidth(self._LEFT_PANEL_MIN_WIDTH)
        left_layout = QVBoxLayout(self._left_panel_frame)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)
        left_layout.addWidget(self._batch_list)
        self._left_panel_frame.setVisible(False)
        grid.addWidget(self._left_panel_frame, 0, 0)

        # ── Center column: input + action + info + bottom bar ───────────
        self._center_column = self._build_center_column()
        grid.addWidget(self._center_column, 0, 1)

        # ── Right panel: template selector | conversion panel ───────────
        self._right_panel_frame = QFrame()
        self._right_panel_frame.setObjectName("rightPanelFrame")
        self._right_panel_frame.setMinimumWidth(self._RIGHT_PANEL_MIN_WIDTH)
        self._right_panel_frame.setVisible(False)
        right_layout = QVBoxLayout(self._right_panel_frame)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)

        from .widgets.template_selector_tabbed import TabbedTemplateSelector

        self._template_selector = TabbedTemplateSelector(
            self._right_panel_frame,
            on_template_selected=self._on_main_window_template_selected,
            on_tab_changed=self._on_main_window_template_tab_changed,
            on_open_location=self._open_template_location,
            on_open_directory=self._open_template_directory,
        )
        self._template_selector.setObjectName("mainWindowTemplateSelector")

        self._right_stack = QStackedWidget(self._right_panel_frame)
        self._right_stack.setObjectName("rightPanelStack")
        self._right_stack.addWidget(self._template_selector)
        self._right_stack.addWidget(self._conversion_panel)
        right_layout.addWidget(self._right_stack, 1)
        self._load_templates_into_main_selector()

        grid.addWidget(self._right_panel_frame, 0, 2)

        grid.setColumnStretch(0, 0)  # hidden left panel should not reserve whitespace
        grid.setColumnStretch(1, 4)  # center column
        grid.setColumnStretch(2, 0)  # hidden right panel should not reserve whitespace

        root_layout.addWidget(central_container, 1)

        self._wire_child_components()
        self._wire_view_model()
        self._wire_projection()
        self._load_initial_preferences()
        self._install_shortcuts()
        self._sync_files_from_main_vm(self._view_model.files)
        self._setup_system_tray()
        self._restore_window_state()

    def _load_initial_preferences(self) -> None:
        controller = self._view_model.controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        if cfg_port is not None:
            preset = cfg_port.get("gui.font.size_preset", "default")
            self._apply_font_size_preset(str(preset or "default"), persist=False)
        self._apply_window_opacity()
        self._initialize_app_icon()

    def _create_actions(self) -> None:
        """Create centralized QAction objects shared between shortcuts and UI."""
        # ── About ─────────────────────────────────────────────────────
        self._action_about = QAction(self)
        self._action_about.setObjectName("actionAbout")
        self._action_about.setText(_t("main_window.about_tooltip", "About DocWen"))
        self._action_about.setShortcut(QKeySequence("Ctrl+?"))
        self._action_about.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        self._action_about.triggered.connect(self._show_about_dialog)
        self.addAction(self._action_about)

        # ── Settings ─────────────────────────────────────────────────
        self._action_settings = QAction(self)
        self._action_settings.setObjectName("actionSettings")
        self._action_settings.setText(_t("main_window.settings_tooltip", "Settings"))
        self._action_settings.setShortcut(QKeySequence("Ctrl+,"))
        self._action_settings.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        self._action_settings.triggered.connect(self._open_settings_dialog)
        self.addAction(self._action_settings)

        # ── Add File ─────────────────────────────────────────────────
        self._action_add_file = QAction(self)
        self._action_add_file.setObjectName("actionAddFile")
        self._action_add_file.setText(_t("main_window.add_file", "Add File"))
        self._action_add_file.setShortcut(QKeySequence("Ctrl+O"))
        self._action_add_file.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        self._action_add_file.triggered.connect(self._on_add_file_shortcut)
        self.addAction(self._action_add_file)

        # ── Start Conversion ─────────────────────────────────────────
        self._action_convert = QAction(self)
        self._action_convert.setObjectName("actionConvert")
        self._action_convert.setText(_t("main_window.start_conversion", "Start Conversion"))
        self._action_convert.setShortcut(QKeySequence("Ctrl+Return"))
        self._action_convert.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        self._action_convert.triggered.connect(self._on_trigger_primary_shortcut)
        self.addAction(self._action_convert)

        # ── Cancel ───────────────────────────────────────────────────
        self._action_cancel = QAction(self)
        self._action_cancel.setObjectName("actionCancel")
        self._action_cancel.setText(_t("main_window.cancel", "Cancel"))
        self._action_cancel.setShortcut(QKeySequence("Esc"))
        self._action_cancel.setShortcutContext(Qt.ShortcutContext.WindowShortcut)
        self._action_cancel.triggered.connect(self._on_esc_shortcut)
        self.addAction(self._action_cancel)

    def _create_components(self) -> None:
        from .view_models.action_area_vm import ActionAreaViewModel
        from .view_models.batch_list_vm import BatchListViewModel
        from .view_models.conversion_panel_vm import ConversionPanelViewModel
        from .view_models.info_area_vm import InfoAreaViewModel
        from .view_models.input_area_vm import InputAreaViewModel
        from .widgets.action_area import ActionArea
        from .widgets.batch_list import BatchList
        from .widgets.conversion_panel import ConversionPanel
        from .widgets.info_area import InfoArea
        from .widgets.input_area import InputArea

        self._input_area_vm = InputAreaViewModel(main_vm=self._view_model, parent=self)
        self._batch_list_vm = BatchListViewModel(main_vm=self._view_model, parent=self)
        self._conversion_panel_vm = ConversionPanelViewModel(main_vm=self._view_model, parent=self)
        self._action_area_vm = ActionAreaViewModel(main_vm=self._view_model, parent=self)
        self._info_area_vm = InfoAreaViewModel(parent=self)

        self._action_area_vm.set_mode(self._view_model.mode)

        self._input_area = InputArea(view_model=self._input_area_vm, parent=self)
        self._batch_list = BatchList(view_model=self._batch_list_vm, parent=self)
        self._conversion_panel = ConversionPanel(view_model=self._conversion_panel_vm, parent=self)
        self._action_area = ActionArea(view_model=self._action_area_vm, parent=self)
        self._info_area = InfoArea(view_model=self._info_area_vm, parent=self)

        # Root object names are set by each widget's __init__;
        # do NOT override them here so that CSS selectors remain effective.

    def _build_center_column(self) -> QWidget:
        column = QWidget(self)
        column.setObjectName("mainWindowCenterColumn")
        layout = QVBoxLayout(column)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        layout.addWidget(self._input_area)
        layout.addWidget(self._action_area)
        layout.addWidget(self._info_area, 1)
        layout.addWidget(self._build_bottom_bar())
        return column

    def _build_bottom_bar(self) -> QWidget:
        bar = QWidget()
        bar.setObjectName("bottomBar")
        bar.setFixedHeight(self._scale_window_value(48))

        layout = QGridLayout(bar)
        layout.setContentsMargins(12, 0, 12, 0)
        layout.setSpacing(4)

        left_actions = QWidget(bar)
        left_actions.setObjectName("bottomBarLeftActions")
        left_layout = QHBoxLayout(left_actions)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)

        self._font_size_btn = self._create_bottom_tool_button(
            object_name="fontSizeButton",
            icon_name="font_size.svg",
            fallback_text="A",
            tooltip=_t("main_window.font_size_tooltip", "Change font size"),
        )
        self._font_size_btn.clicked.connect(self._show_font_size_menu)
        left_layout.addWidget(self._font_size_btn)

        self._about_btn = self._create_bottom_tool_button(
            object_name="aboutButton",
            icon_name="about.svg",
            fallback_text="?",
            tooltip=_t("main_window.about_tooltip", "About DocWen"),
        )
        self._about_btn.clicked.connect(self._action_about.trigger)
        left_layout.addWidget(self._about_btn)

        right_actions = QWidget(bar)
        right_actions.setObjectName("bottomBarRightActions")
        right_layout = QHBoxLayout(right_actions)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(4)

        self._settings_btn = self._create_bottom_tool_button(
            object_name="settingsButton",
            icon_name="settings.svg",
            fallback_text="[=]",
            tooltip=_t("main_window.settings_tooltip", "Settings (Ctrl+,)"),
        )
        self._settings_btn.clicked.connect(self._action_settings.trigger)
        right_layout.addWidget(self._settings_btn)

        self._version_label = QLabel(_t("main_window.version_offline", "Offline"))
        self._version_label.setObjectName("versionLabel")
        self._version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(left_actions, 0, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(self._version_label, 0, 1, Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(right_actions, 0, 2, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        for column in range(3):
            layout.setColumnStretch(column, 1)
        return bar

    def _create_bottom_tool_button(
        self,
        *,
        object_name: str,
        icon_name: str,
        fallback_text: str,
        tooltip: str,
    ) -> QToolButton:
        button = QToolButton()
        button.setObjectName(object_name)
        button.setToolTip(tooltip)
        button.setAccessibleName(tooltip)
        button.setFixedSize(36, 36)

        icon = load_svg_icon(icon_name)
        if icon is not None and not icon.isNull():
            button.setIcon(icon)
            button.setIconSize(QSize(20, 20))
            button.setText("")
        else:
            button.setText(fallback_text)
        return button

    # ── Wiring ──────────────────────────────────────────────────────

    def _wire_child_components(self) -> None:
        self._input_area_vm.files_cleared.connect(self._view_model.clear_files)
        self._batch_list.selection_changed.connect(self._on_selected_file_changed)
        self._batch_list.entry_action_requested.connect(self._handle_batch_entry_action)
        self._conversion_panel_vm.conversion_requested.connect(self._handle_conversion_panel_conversion_requested)
        self._conversion_panel_vm.named_action_requested.connect(self._handle_named_action_requested)
        self._action_area_vm.conversion_requested.connect(self._handle_action_area_conversion_requested)
        self._action_area_vm.cancel_requested.connect(self._cancel_active_task)
        self._action_area_vm.state_changed.connect(self._sync_session_mutation_controls)
        self._info_area_vm.history_navigation_requested.connect(self._handle_navigation_request)
        self._info_area_vm.location_requested.connect(self._open_location)
        self._info_area_vm.task_guide_action_requested.connect(self._handle_task_guide_action)
        self._sync_session_mutation_controls()

    def _sync_session_mutation_controls(self) -> None:
        """Prevent clearing the working set while an execution can still emit results."""
        clear_button = getattr(self._input_area, "clear_button", None)
        if clear_button is not None:
            clear_button.setEnabled(not self._action_area_vm.cancel_visible)

    def _wire_view_model(self) -> None:
        vm = self._view_model
        vm.title_changed.connect(self.setWindowTitle)
        vm.status_message_changed.connect(self._on_status_message_changed)
        vm.window_activation_requested.connect(self.bring_to_front)
        vm.files_changed.connect(self._sync_files_from_main_vm)
        vm.files_cleared.connect(self._on_files_cleared)
        vm.mode_changed.connect(self._on_mode_changed)
        vm.shutdown_requested.connect(self.close)
        vm.ipc_file_received.connect(self._on_ipc_file_received)

        self._task_event_bridge.task_event.connect(vm.on_task_event)
        self._task_event_bridge.flush_error.connect(self._on_bridge_flush_error)
        self._task_event_bridge.start_auto_flush()

    def _wire_projection(self) -> None:
        """Connect VM projection changes to right-panel widget bindings."""
        self._view_model.ui_projection_changed.connect(self._on_projection_changed)
        # Apply initial projection so the window opens in the correct state.
        self._on_projection_changed(self._view_model.ui_projection)

    def _on_projection_changed(self, projection: MainWindowUiProjection) -> None:
        """Render the current UI projection into widget visibility and stack index."""
        from .view_models.interaction import RightPanelSlot

        if self._left_panel_frame is None or self._right_panel_frame is None or self._right_stack is None:
            return

        window_is_normal = not (self.isMaximized() or self.isFullScreen() or self.isMinimized())
        old_left_visible = self._left_panel_frame.isVisible()
        old_right_visible = self._right_panel_frame.isVisible()
        visibility_changed = (
            old_left_visible != projection.left_panel_visible or old_right_visible != projection.right_panel_visible
        )
        rebase_geometry_source_change = (
            visibility_changed
            and self._geometry_source_change_rect is not None
            and self._geometry_source_change_rect == self._normal_window_snapshot()[0]
        )
        self._activate_window_layout()
        preserved_normal_rect: WindowRect | None = None
        if not self.isVisible():
            preserved_center_x = None
        elif (
            window_is_normal
            and self._prepared_file_clear_normal_rect is not None
            and self._prepared_file_clear_center_x is not None
        ):
            preserved_normal_rect = self._prepared_file_clear_normal_rect
            preserved_center_x = self._prepared_file_clear_center_x
        elif window_is_normal:
            preserved_normal_rect = self._normal_window_snapshot()[0]
            preserved_center_x = self._center_column_screen_x()
        elif self._pending_normal_center_anchor is not None:
            preserved_center_x = self._pending_normal_center_anchor
        elif self._last_normal_window_rect is not None:
            preserved_center_x = self._last_normal_window_rect.x + self._last_normal_center_offset
        else:
            preserved_center_x = None
        atomic_visible_transition = (
            visibility_changed and window_is_normal and self.isVisible() and preserved_center_x is not None
        )
        resume_file_clear_updates = self._file_clear_updates_suspended
        if atomic_visible_transition and not resume_file_clear_updates:
            self.setUpdatesEnabled(False)
        if atomic_visible_transition:
            self._applying_internal_panel_geometry = True
        try:
            self._left_panel_frame.setVisible(projection.left_panel_visible)
            self._right_panel_frame.setVisible(projection.right_panel_visible)
            self._apply_projection_minimum_widths(
                left_visible=projection.left_panel_visible,
                right_visible=projection.right_panel_visible,
            )
            self._apply_side_panel_stretch(
                left_visible=projection.left_panel_visible,
                right_visible=projection.right_panel_visible,
            )
            if visibility_changed and window_is_normal:
                if atomic_visible_transition:
                    assert preserved_center_x is not None
                    assert preserved_normal_rect is not None
                    rect = self._normal_panel_transition_rect(
                        preserved_center_x=preserved_center_x,
                        preserved_normal_rect=preserved_normal_rect,
                        old_left_visible=old_left_visible,
                        old_right_visible=old_right_visible,
                        left_visible=projection.left_panel_visible,
                        right_visible=projection.right_panel_visible,
                    )
                    self._set_normal_frame_geometry(rect)
                else:
                    target_width = self._panel_transition_target_width(
                        old_left_visible=old_left_visible,
                        old_right_visible=old_right_visible,
                        left_visible=projection.left_panel_visible,
                        right_visible=projection.right_panel_visible,
                    )
                    if target_width != self.width():
                        self.resize(target_width, self.height())
            self._activate_window_layout()
        finally:
            self._applying_internal_panel_geometry = False
            if atomic_visible_transition or resume_file_clear_updates:
                self._prepared_file_clear_normal_rect = None
                self._prepared_file_clear_center_x = None
                self._file_clear_updates_suspended = False
                self.setUpdatesEnabled(True)
                self.update()
        if preserved_center_x is not None and not window_is_normal:
            self._pending_normal_center_anchor = preserved_center_x
            self._pending_normal_frame_restored = False
        self._capture_normal_window_geometry(rebase_collapsed=False)
        if rebase_geometry_source_change:
            self._geometry_source_change_rect = self._last_normal_window_rect
            self._geometry_source_transition_rebase_pending = True
            self._geometry_source_transition_settling = True
            QTimer.singleShot(16, self._finish_geometry_source_transition_settling)

        if projection.right_panel_slot == RightPanelSlot.TEMPLATE:
            template_selector = self._template_selector
            if template_selector is None:
                return
            self._right_stack.setCurrentWidget(template_selector)
            if projection.template_context is not None:
                file_path = projection.template_context.file_path
                self._show_template_target_mode(template_selector.current_tab, file_path)
        elif projection.right_panel_slot == RightPanelSlot.CONVERSION:
            self._right_stack.setCurrentWidget(self._conversion_panel)
            if projection.conversion_context is not None:
                ctx = projection.conversion_context
                if self._view_model.mode == "batch":
                    file_list = self._batch_list_vm.get_files_for_category(self._batch_list_vm.current_category)
                else:
                    file_list = self._batch_list_vm.get_files()
                self._conversion_panel_vm.set_file_info(
                    category=ctx.category,
                    current_format=ctx.current_format or "",
                    file_path=ctx.file_path,
                    file_list=file_list,
                    ui_mode=str(self._view_model.mode),
                )
                category = ctx.category
                file_path = ctx.file_path
                source_inputs = tuple(
                    OptimizationSource(
                        detected_format=entry.detected_format,
                        source_category=entry.workflow_category,
                    )
                    for path in file_list
                    if (entry := self._batch_list_vm.get_file_entry(path)) is not None
                )
                if not source_inputs:
                    source_inputs = (
                        OptimizationSource(
                            detected_format=ctx.current_format or category,
                            source_category=category,
                        ),
                    )
                if category == "document":
                    self._action_area_vm.setup_for_document_file(
                        file_path,
                        ctx.current_format or "docx",
                        source_inputs=source_inputs,
                    )
                elif category == "spreadsheet":
                    self._action_area_vm.setup_for_spreadsheet_file(
                        file_path,
                        ctx.current_format or "xlsx",
                        source_inputs=source_inputs,
                    )
                elif category == "image":
                    self._action_area_vm.setup_for_image_file(
                        file_path,
                        ctx.current_format or "image",
                        source_inputs=source_inputs,
                    )
                elif category == "layout":
                    self._action_area_vm.setup_for_layout_file(
                        file_path,
                        ctx.current_format or "pdf",
                        source_inputs=source_inputs,
                    )
                    if (ctx.current_format or "").strip().lower() == "pdf":
                        total_pages = _read_pdf_total_pages(file_path)
                        if total_pages is not None:
                            self._conversion_panel_vm.set_pdf_info(total_pages, Path(file_path).name)
                        else:
                            self._conversion_panel_vm.set_pdf_info(0, "")
        else:
            self._conversion_panel_vm.reset()
            selected_file = self._view_model.selected_file
            if selected_file is not None:
                file_path = getattr(selected_file, "path", "")
                fmt = getattr(selected_file, "format", "")
                category = getattr(selected_file, "category", "other") or "other"
                self._action_area_vm.setup_for_other_file(file_path, fmt, source_category=category)
            elif not resume_file_clear_updates:
                self._action_area_vm.reset()

    def _panel_transition_target_width(
        self,
        *,
        old_left_visible: bool,
        old_right_visible: bool,
        left_visible: bool,
        right_visible: bool,
    ) -> int:
        """Return the fitted width after adding/removing panel contributions."""
        central = self.findChild(QWidget, "centralContainer")
        layout = central.layout() if central is not None else None
        spacing = layout.horizontalSpacing() if isinstance(layout, QGridLayout) else 0
        old_contribution = (self._LEFT_PANEL_MIN_WIDTH + spacing if old_left_visible else 0) + (
            self._RIGHT_PANEL_MIN_WIDTH + spacing if old_right_visible else 0
        )
        new_contribution = (self._LEFT_PANEL_MIN_WIDTH + spacing if left_visible else 0) + (
            self._RIGHT_PANEL_MIN_WIDTH + spacing if right_visible else 0
        )
        collapsed = self._collapsed_normal_window_rect
        base_width = max(self.minimumWidth(), self.width() - old_contribution) if collapsed is None else collapsed.width
        target_width = base_width + new_contribution
        qt_screen = self.screen() or QApplication.primaryScreen()
        fit = self._screen_fit_rect(qt_screen) if qt_screen is not None else None
        if fit is not None:
            target_width = min(target_width, fit.width)
        target_width = max(self.minimumWidth(), target_width)
        return target_width

    def _normal_panel_transition_rect(
        self,
        *,
        preserved_center_x: int,
        preserved_normal_rect: WindowRect,
        old_left_visible: bool,
        old_right_visible: bool,
        left_visible: bool,
        right_visible: bool,
    ) -> WindowRect:
        """Resolve one final normal-frame rectangle before allowing a repaint."""
        target_width = self._panel_transition_target_width(
            old_left_visible=old_left_visible,
            old_right_visible=old_right_visible,
            left_visible=left_visible,
            right_visible=right_visible,
        )
        central = self.findChild(QWidget, "centralContainer")
        layout = central.layout() if central is not None else None
        if central is not None and layout is not None:
            self._apply_projection_minimum_widths(
                left_visible=left_visible,
                right_visible=right_visible,
                container_width=target_width,
            )
            # Ask Qt's real grid (including optional side-panel stretch) for the
            # final child positions without resizing the native top-level frame.
            layout.setGeometry(QRect(0, 0, target_width, max(1, central.height())))
        target_center_offset = self._center_column.mapTo(self, QPoint(0, 0)).x()
        collapsed = self._collapsed_normal_window_rect or preserved_normal_rect
        if self._collapsed_normal_window_rect is None:
            collapsed_center_x = int(preserved_center_x)
        else:
            collapsed_center_x = collapsed.x + self._collapsed_normal_center_offset
        if not left_visible and not right_visible:
            candidate = collapsed
        else:
            candidate = WindowRect(
                x=collapsed_center_x - target_center_offset,
                y=collapsed.y,
                width=target_width,
                height=collapsed.height,
            )
        recovered = recover_window_geometry(
            candidate,
            self._available_screen_rects(),
            min_width=self._window_geometry.min_width,
            min_height=self._window_geometry.min_height,
        )
        return recovered.rect

    def _apply_projection_minimum_widths(
        self,
        *,
        left_visible: bool,
        right_visible: bool,
        container_width: int | None = None,
    ) -> None:
        """Apply viewport-relative minima without retaining hidden-panel pixels."""
        plan = self._panel_width_plan(
            left_visible=left_visible,
            right_visible=right_visible,
            container_width=container_width,
        )
        self._left_panel_frame.setMinimumWidth(plan.left)
        self._center_column.setMinimumWidth(plan.center)
        self._right_panel_frame.setMinimumWidth(plan.right)

    def _panel_width_plan(
        self,
        *,
        left_visible: bool,
        right_visible: bool,
        container_width: int | None = None,
    ) -> PanelWidthPlan:
        """Resolve configured panel-width intent for one runtime viewport."""
        central = self.findChild(QWidget, "centralContainer")
        layout = central.layout() if central is not None else None
        if central is None or not isinstance(layout, QGridLayout):
            return PanelWidthPlan(left=0, center=0, right=0)
        spacing = layout.horizontalSpacing()
        margins = layout.contentsMargins()
        return plan_panel_widths(
            self._window_behavior,
            container_width=max(0, central.width() if container_width is None else int(container_width)),
            horizontal_margins=margins.left() + margins.right(),
            spacing=spacing,
            left_visible=left_visible,
            right_visible=right_visible,
            scale_factor=self._window_scale_factor,
        )

    def _apply_side_panel_stretch(self, *, left_visible: bool, right_visible: bool) -> None:
        """Apply expansion only to visible side panels; hidden columns stay inert."""
        left_parent = self._left_panel_frame.parentWidget()
        layout = left_parent.layout() if left_parent is not None else None
        if not isinstance(layout, QGridLayout):
            return
        expand = self._window_behavior.expand_side_panels
        layout.setColumnStretch(0, 2 if expand and left_visible else 0)
        layout.setColumnStretch(1, 4)
        layout.setColumnStretch(2, 3 if expand and right_visible else 0)

    def _apply_runtime_window_settings(self) -> None:
        """Refresh committed flags without applying startup-only placement."""
        persisted_geometry = self._load_window_geometry(center_offset=0)
        if persisted_geometry != self._persisted_window_geometry_policy:
            self._geometry_source_change_rect = self._normal_window_rect()
            self._geometry_source_transition_rebase_pending = False
            self._geometry_source_transition_settling = False
            self._persisted_window_geometry_policy = persisted_geometry
        self._window_behavior = self._load_window_behavior()
        self._CENTER_PANEL_MIN_WIDTH = self._scale_window_value(self._window_behavior.center_panel_width)
        self._LEFT_PANEL_MIN_WIDTH = self._scale_window_value(self._window_behavior.left_panel_width)
        self._RIGHT_PANEL_MIN_WIDTH = self._scale_window_value(self._window_behavior.right_panel_width)
        projection = self._view_model.ui_projection
        self._apply_projection_minimum_widths(
            left_visible=projection.left_panel_visible,
            right_visible=projection.right_panel_visible,
        )
        self._apply_side_panel_stretch(
            left_visible=projection.left_panel_visible,
            right_visible=projection.right_panel_visible,
        )
        self._restore_main_template_default()

    def _finish_geometry_source_transition_settling(self) -> None:
        """Finish the bounded internal-layout quiet period after a panel change."""
        if not self._geometry_source_transition_rebase_pending:
            return
        self._activate_window_layout()
        self._capture_normal_window_geometry()
        self._geometry_source_change_rect = self._last_normal_window_rect
        self._geometry_source_transition_settling = False

    def _finish_prepared_file_clear_panel_transition(self) -> None:
        """Release a prepared file-clear repaint if no projection consumed it."""
        if not self._file_clear_updates_suspended:
            return
        self._prepared_file_clear_normal_rect = None
        self._prepared_file_clear_center_x = None
        self._file_clear_updates_suspended = False
        self._applying_internal_panel_geometry = False
        self.setUpdatesEnabled(True)
        self.update()

    def _on_main_window_template_selected(self, template_type: str, template_name: str) -> None:
        """Handle template selection from the main-window template selector."""
        sel = self._view_model.selected_file
        selected_path = getattr(sel, "path", None) if sel is not None else None
        if not selected_path:
            return
        self._show_template_target_mode(template_type, selected_path)

    def _on_main_window_template_tab_changed(self, template_type: str, _previous_type: str) -> None:
        """Keep the centre action surface aligned with the active template tab."""
        sel = self._view_model.selected_file
        selected_path = getattr(sel, "path", None) if sel is not None else None
        if selected_path:
            self._show_template_target_mode(template_type, selected_path)

    def _show_template_target_mode(self, template_type: str, file_path: str) -> None:
        """Project a template target into the matching Markdown generation mode."""
        normalized_type = str(template_type or "").strip().lower()
        expected_mode = "md_to_spreadsheet" if normalized_type == "xlsx" else "docx"
        if (
            self._action_area_vm.mode == self._view_model.mode
            and self._action_area_vm.file_type == expected_mode
            and _normalize_path(self._action_area_vm.file_path or "") == _normalize_path(file_path)
        ):
            return
        self._action_area_vm.set_mode(self._view_model.mode)
        if normalized_type == "xlsx":
            self._action_area_vm.setup_for_md_to_spreadsheet(file_path)
        else:
            self._action_area_vm.setup_for_md_to_document(file_path)

    def _configured_main_template_type(self) -> str:
        """Read and normalize the persisted Markdown template target."""
        controller = getattr(self._view_model, "controller", None)
        config_port = getattr(controller, "config_port", None)
        if config_port is None:
            return "docx"
        try:
            value = str(config_port.get("gui.template.md_default_template", "docx") or "docx")
        except Exception:
            return "docx"
        normalized = value.strip().lower()
        return normalized if normalized in {"docx", "xlsx"} else "docx"

    def _restore_main_template_default(self, *, force: bool = False) -> str:
        """Apply a changed template default without disturbing unrelated saves."""
        preferred_type = self._configured_main_template_type()
        changed = preferred_type != self._main_template_default_type
        self._main_template_default_type = preferred_type
        selector = self._template_selector
        if selector is None or (not force and not changed):
            return preferred_type

        resolved_type = selector.restore_current_tab(preferred_type)
        projection = self._view_model.ui_projection
        from .view_models.interaction import RightPanelSlot

        if projection.right_panel_slot == RightPanelSlot.TEMPLATE and projection.template_context is not None:
            self._show_template_target_mode(resolved_type, projection.template_context.file_path)
        return resolved_type

    def _load_templates_into_main_selector(self) -> None:
        """Populate the main-window template selector from the runtime registry."""
        if self._template_selector is None:
            return
        try:
            from docwen_gui.widgets.template_selector import TemplateItemDetails
            from docwen_runtime.templates import TemplateRegistry

            registry = TemplateRegistry.default()
            templates: dict[str, list[str]] = {"docx": [], "xlsx": []}
            details: dict[str, dict[str, TemplateItemDetails]] = {"docx": {}, "xlsx": {}}
            for info in registry.list_templates():
                target = str(getattr(info, "target", "") or "").strip().lower()
                if target in templates:
                    name = str(getattr(info, "name", "") or "")
                    if not name:
                        continue
                    path_value = getattr(info, "path", "") or ""
                    path = Path(path_value) if path_value else None
                    templates[target].append(name)
                    details[target][name] = TemplateItemDetails(
                        resource_id=str(getattr(info, "id", "") or "") or None,
                        usage_hint=str(getattr(info, "description", "") or "") or None,
                        source_label=(path.parent.name or str(path.parent) or None) if path is not None else None,
                        source_path=str(path) if path is not None else None,
                        updated_label=_format_template_modified_label(getattr(info, "modified_ns", None)),
                    )
            self._template_selector.load_all_templates(templates, details=details)
            self._restore_main_template_default(force=True)
        except Exception:
            logger.exception("Unable to load the runtime template registry")
            message = _t(
                "main_window.template_catalog_failed",
                "Templates could not be loaded.",
            )
            self._template_selector.show_load_error(
                _t("components.template_selector.unavailable", "Unavailable"),
                message,
            )
            self._info_area_vm.add_message(message, "warning")

    def _merge_template_options_for_request(
        self,
        target_format: str,
        options: dict[str, Any],
        *,
        source_format: str,
        source_category: str,
        action_name: str,
    ) -> dict[str, Any]:
        """Add selected template metadata for Markdown document/spreadsheet targets."""
        merged = dict(options)
        target = str(target_format or "").lower()
        if action_name or "template_name" in merged:
            return merged
        from .view_models.interaction import FileCapability, resolve_capabilities

        if FileCapability.TEMPLATE_SELECTION not in resolve_capabilities(source_category):
            return merged
        selector = self._template_selector
        selected = selector.get_selected_template_resource() if selector is not None else None
        if selected is None:
            return merged
        template_type, template_id = selected
        if target in _DOCUMENT_TEMPLATE_TARGETS and template_type == "docx":
            merged["template_name"] = template_id
        if target in _SPREADSHEET_TEMPLATE_TARGETS and template_type == "xlsx":
            merged["template_name"] = template_id
        return merged

    def _open_template_location(self, template_type: str, template_name: str) -> None:
        """Open the selected template's containing folder."""
        template_path = self._resolve_template_path(template_type, template_name)
        if template_path is None:
            self._info_area_vm.set_transient_message(
                f"template:not-found:{template_type}:{template_name}",
                _t("main_window.template_not_found", "Template not found: {name}", name=template_name),
                "warning",
                ttl_ms=3000,
                source="template-selector",
            )
            return
        self._open_path(str(template_path), open_parent=True)

    def _open_template_directory(self, template_type: str) -> None:
        """Open the template directory for empty-state recovery."""
        try:
            from docwen_runtime.resources import ResourceRegistry

            templates_dir = ResourceRegistry.default().templates_dir()
        except Exception:
            self._info_area_vm.set_transient_message(
                f"template:dir-unavailable:{template_type}",
                _t("main_window.template_dir_unavailable", "Template directory is unavailable."),
                "warning",
                ttl_ms=3000,
                source="template-selector",
            )
            return
        self._open_path(str(templates_dir), open_parent=False)

    @staticmethod
    def _resolve_template_path(template_type: str, template_name: str) -> Path | None:
        """Resolve a template name through the runtime template registry."""
        try:
            from docwen_runtime.templates import TemplateRegistry

            requested_type = str(template_type or "").strip().lower()
            requested_name = str(template_name or "").strip()
            if not requested_name:
                return None
            registry = TemplateRegistry.default()
            for template in registry.list_templates(requested_type or None):
                if (
                    template.name.casefold() == requested_name.casefold()
                    or template.path.name.casefold() == requested_name.casefold()
                ):
                    return template.path
            return None
        except Exception:
            return None

    def _install_shortcuts(self) -> None:
        # Each shortcut uses WindowShortcut context so it fires regardless
        # of which child widget has focus.  Handlers guard against
        # triggering during text editing via _has_editable_text_focus().
        def _build(key: str, callback) -> QShortcut:
            sc = QShortcut(QKeySequence(key), self)
            sc.setContext(Qt.ShortcutContext.WindowShortcut)
            sc.setAutoRepeat(False)
            sc.activated.connect(callback)
            return sc

        # NOTE: Ctrl+O, Ctrl+,, Esc are now handled by centralized QActions
        # (see _create_actions) so that buttons and shortcuts share the same
        # action object.  Ctrl+Return (start conversion) is also a QAction.
        _build("Ctrl+Shift+T", self._on_toggle_always_on_top_shortcut)
        _build("Ctrl+Shift+O", self._on_add_folder_shortcut)
        _build("Ctrl+L", self._on_locate_output_shortcut)
        _build("Delete", self._on_remove_selected_shortcut)
        _build("Return", self._on_trigger_primary_shortcut)
        _build("Enter", self._on_trigger_primary_shortcut)

    # ── Focus protection ───────────────────────────────────────────

    @staticmethod
    def _has_editable_text_focus() -> bool:
        """Return True when the focus widget is a text-editing control.

        Covers QLineEdit, QTextEdit, QPlainTextEdit, QAbstractSpinBox,
        and editable QComboBox.  Shortcut handlers use this to avoid
        stealing keystrokes that should go to the text editor.
        """
        focus_widget = QApplication.focusWidget()
        if isinstance(focus_widget, (QLineEdit, QTextEdit, QPlainTextEdit, QAbstractSpinBox)):
            return True
        if isinstance(focus_widget, QComboBox):
            return focus_widget.isEditable()
        return False

    # ── Shortcut handlers ──────────────────────────────────────────

    def _on_add_file_shortcut(self) -> None:
        if self._has_editable_text_focus():
            return
        self._input_area.open_file_dialog()

    def _on_add_folder_shortcut(self) -> None:
        if self._has_editable_text_focus():
            return
        self._input_area.open_folder_dialog(force_batch_mode=True)

    def _on_locate_output_shortcut(self) -> None:
        if self._has_editable_text_focus():
            return
        current = self._batch_list.get_current_file()
        target_path: str | None = None
        if current:
            entry = self._batch_list_vm.get_file_entry(current)
            if entry is not None:
                target_path = entry.output_path or current
        if not target_path:
            sel = self._view_model.selected_file
            target_path = getattr(sel, "path", None) if sel is not None else None
        if not target_path:
            self._info_area_vm.set_transient_message(
                "locate:no-target",
                _t("main_window.no_file_for_locate", "No file selected to locate output."),
                "warning",
                ttl_ms=2500,
                source="shortcut",
            )
            return
        self._open_path(target_path, open_parent=True)

    def _on_remove_selected_shortcut(self) -> None:
        if self._has_editable_text_focus():
            return
        current = self._batch_list.get_current_file()
        if current:
            self._batch_list_vm.remove_file(current)

    def _on_esc_shortcut(self) -> None:
        if self._has_editable_text_focus():
            return
        self._cancel_active_task()

    def _on_toggle_always_on_top_shortcut(self) -> None:
        if self._has_editable_text_focus():
            return
        self.toggle_always_on_top()

    def _on_trigger_primary_shortcut(self) -> None:
        if self._has_editable_text_focus():
            return
        self._trigger_primary_action()

    # ── ViewModel sync ──────────────────────────────────────────────

    def _sync_files_from_main_vm(self, file_refs: Sequence[FileRef]) -> None:
        self._file_contexts = {
            _normalize_path(ref.path): (
                ref.format.lower(),
                ref.category.lower(),
            )
            for ref in file_refs
            if ref.path
        }
        refs_by_path = {_normalize_path(ref.path): ref for ref in file_refs if ref.path}
        desired_paths = set(refs_by_path)
        current_paths = {_normalize_path(path) for path in self._batch_list_vm.get_files()}

        missing = [ref.path for ref in file_refs if _normalize_path(ref.path) not in current_paths]
        if missing:

            def resolve_existing_ref(path: str) -> dict[str, Any] | None:
                ref = refs_by_path.get(_normalize_path(path))
                if ref is None:
                    return None
                return {
                    "detected_format": ref.format,
                    "workflow_category": ref.category,
                    "warning_message": ref.warning_message,
                    "metadata": dict(ref.metadata),
                }

            self._batch_list_vm.add_files(missing, file_resolver=resolve_existing_ref)

        for stale in current_paths - desired_paths:
            self._batch_list_vm.remove_file(stale)

        if file_refs:
            if not self._input_area_vm.selection_message:
                self._input_area_vm.sync_selection(file_refs)
            sel = self._view_model.selected_file
            preferred_key = _normalize_path(getattr(sel, "path", "")) if sel is not None else ""
            if preferred_key not in desired_paths:
                current_batch_file = self._batch_list_vm.get_current_file()
                preferred_key = _normalize_path(current_batch_file or "")
            preferred_ref = refs_by_path.get(preferred_key)
            if preferred_ref is not None:
                self._view_model.set_selected_file(preferred_ref)
                with contextlib.suppress(Exception):
                    self._batch_list.select_file(preferred_ref.path)

    def _on_files_cleared(self) -> None:
        self._prepare_file_clear_panel_transition()
        self._file_contexts.clear()
        self._batch_list_vm.clear_files()
        self._action_area_vm.reset()
        self._info_area_vm.reset_session()

    def _prepare_file_clear_panel_transition(self) -> None:
        """Freeze geometry before content reset and right-panel removal run synchronously."""
        if (
            not self.isVisible()
            or self.isMaximized()
            or self.isFullScreen()
            or self.isMinimized()
            or not self._right_panel_frame.isVisible()
        ):
            return
        self._activate_window_layout()
        self._prepared_file_clear_normal_rect = self._normal_window_snapshot()[0]
        self._prepared_file_clear_center_x = self._center_column_screen_x()
        self._file_clear_updates_suspended = True
        self._applying_internal_panel_geometry = True
        self.setUpdatesEnabled(False)
        QTimer.singleShot(0, self._finish_prepared_file_clear_panel_transition)

    def _on_mode_changed(self, mode: str) -> None:
        self._current_mode = mode
        self._action_area_vm.set_mode(mode)
        if mode == "single" and self._view_model.selected_file is None:
            self._select_current_batch_file()

    def _select_current_batch_file(self) -> bool:
        """Select the current batch-list file as the main selected file."""
        current_file = self._batch_list_vm.get_current_file()
        if not current_file:
            return False
        current_key = _normalize_path(current_file)
        for ref in self._view_model.files:
            if _normalize_path(getattr(ref, "path", "")) != current_key:
                continue
            self._view_model.set_selected_file(ref)
            with contextlib.suppress(Exception):
                self._batch_list.select_file(getattr(ref, "path", ""))
            return True
        return False

    def _on_ipc_file_received(self, file_path: str) -> None:
        self._input_area_vm.sync_selection(self._view_model.files)
        self._info_area_vm.add_message(
            _t(
                "main_window.ipc_file_received",
                "Received file from another instance: {filename}",
                filename=Path(file_path).name,
            ),
            "info",
            show_location=True,
            file_path=file_path,
            navigate_file_path=file_path,
        )

    def _on_status_message_changed(self, message: str) -> None:
        if not message:
            return
        key_base = "status"
        tone = "secondary"
        ttl_ms = 2500
        is_terminal = False
        if message.startswith(_t("main_window.task_processing_prefix")):
            key_base = "processing"
            tone = "info"
            ttl_ms = 0
        elif message.startswith(_t("main_window.task_progress_prefix")):
            key_base = "progress"
            tone = "info"
            ttl_ms = 0
        elif message.startswith(_t("main_window.task_failed_prefix")):
            key_base = "error"
            tone = "danger"
            ttl_ms = 4000
            is_terminal = True
        elif message == _t("main_window.task_completed_status"):
            key_base = "terminal"
            tone = "success"
            ttl_ms = 2500
            is_terminal = True
        elif message == _t("main_window.task_cancelled_status"):
            key_base = "terminal"
            tone = "warning"
            ttl_ms = 2500
            is_terminal = True
        if is_terminal:
            # Processing/progress status messages intentionally have no TTL
            # while a task is active. Remove both displayed and throttled
            # variants before publishing the terminal flash; otherwise they
            # reappear after the terminal TTL and permanently mask the task
            # summary rendered by InfoArea.
            self._info_area_vm.clear_transient_message("processing:main-window")
            self._info_area_vm.clear_transient_message("progress:main-window")
            # The authoritative execution result already published the
            # operation-specific terminal transient and detailed summary.
            # Keep MainWindowViewModel's terminal status observable without
            # creating a second generic terminal message.
            return
        self._info_area_vm.set_transient_message(
            f"{key_base}:main-window",
            message,
            tone,
            ttl_ms=ttl_ms,
            source="main-window",
        )

    def _on_bridge_flush_error(self, message: str) -> None:
        self._info_area_vm.add_message(message, "warning")

    # ── Selection and panel coordination ────────────────────────────

    def _on_selected_file_changed(self, file_path: str | None) -> None:
        if not file_path:
            self._view_model.clear_selected_file()
            return
        normalized = _normalize_path(file_path)
        for ref in self._view_model.files:
            if _normalize_path(getattr(ref, "path", "")) == normalized:
                self._view_model.set_selected_file(ref)
                return
        self._view_model.clear_selected_file()

    # ── Conversion execution ────────────────────────────────────────

    def _handle_conversion_panel_conversion_requested(
        self,
        target_format: str,
        file_path: str,
        options: dict[str, Any] | None = None,
    ) -> None:
        """Handle a standard conversion initiated by the right-hand panel."""
        self._handle_conversion_requested(
            target_format,
            file_path,
            options,
            origin="conversion_panel",
        )

    def _handle_action_area_conversion_requested(
        self,
        target_format: str,
        file_path: str,
        options: dict[str, Any] | None = None,
    ) -> None:
        """Handle a conversion or optimization initiated by the center action area."""
        self._handle_conversion_requested(
            target_format,
            file_path,
            options,
            origin="action_area",
        )

    def _handle_conversion_requested(
        self,
        target_format: str,
        file_path: str,
        options: dict[str, Any] | None = None,
        *,
        origin: _ConversionRequestOrigin,
    ) -> None:
        # Only the center ActionArea owns optimization action selection.  The
        # right ConversionPanel is an independent standard-conversion surface;
        # inheriting the ActionArea's stale selection routes a normal button to
        # a named runtime action that the user did not invoke.
        action_name = self._action_area_vm.action_name if origin == "action_area" else ""
        if self._view_model.mode == "batch":
            file_paths = self._batch_list_vm.get_files_for_category(self._batch_list_vm.current_category)
            if len(file_paths) > 1:
                self._start_batch_execution(
                    file_paths=file_paths,
                    target_format=str(target_format).lower(),
                    action_name=action_name,
                    options=dict(options or {}),
                )
                return
        self._start_execution(
            file_path=file_path,
            target_format=str(target_format).lower(),
            action_name=action_name,
            options=dict(options or {}),
        )

    def _handle_named_action_requested(
        self,
        action_name: str,
        file_path: str,
        options: dict[str, Any] | None = None,
    ) -> None:
        if action_name in _AGGREGATE_ACTIONS:
            file_paths = self._batch_list_vm.get_aggregate_file_list(action_name)
            if action_name == "merge_tables":
                file_paths = _move_path_to_front(file_paths, file_path)
            if len(file_paths) < 2:
                self._info_area_vm.add_message(
                    _t("main_window.aggregate_need_two", "At least two matching files are required."),
                    "warning",
                )
                return
            self._start_aggregate_execution(
                file_paths=file_paths,
                target_format="",
                action_name=action_name,
                options=dict(options or {}),
            )
            return

        self._start_execution(
            file_path=file_path,
            target_format="",
            action_name=action_name,
            options=dict(options or {}),
        )

    def _resolve_execution_route(
        self,
        *,
        file_paths: Sequence[str],
        target_format: str,
        action_name: str,
    ) -> tuple[str, RuntimeRouteChoice] | None:
        """Resolve one canonical route for every input before building a request."""

        controller = self._view_model.controller
        sources: list[RuntimeRouteSource] = []
        for file_path in file_paths:
            context = _resolve_file_context(self._file_contexts, self._batch_list_vm, file_path)
            if context is None:
                self._info_area_vm.add_message(
                    _t("main_window.route_unavailable", "No compatible operation is available for this file."),
                    "warning",
                )
                return None
            detected_format, source_category = context
            sources.append(RuntimeRouteSource(detected_format, source_category))
        result = discover_runtime_route_choices(
            controller,
            sources=tuple(sources),
            operation="action" if action_name else "conversion",
            action_name=action_name,
        )
        if result.status == "failed":
            self._info_area_vm.add_message(
                _t(
                    "main_window.route_catalog_failed",
                    "Available operations could not be loaded; the request was not started.",
                ),
                "warning",
            )
            return None
        normalized_target = str(target_format or "").strip().lower()
        choice = result.get(normalized_target) if normalized_target else None
        if choice is None and not normalized_target and len(result.choices) == 1:
            choice = result.choices[0]
        if choice is None:
            self._info_area_vm.add_message(
                _t("main_window.route_unavailable", "No compatible operation is available for this file."),
                "warning",
            )
            return None
        return choice.target, choice

    def _start_execution(
        self,
        *,
        file_path: str,
        target_format: str,
        action_name: str,
        options: dict[str, Any],
    ) -> None:
        if self._execution_close_pending or self._shutdown_finalized:
            return
        controller = self._view_model.controller
        if controller is None or not controller.has_runtime:
            self._info_area_vm.add_message(
                _t("main_window.runtime_unavailable", "Runtime is unavailable; conversion cannot start."),
                "warning",
            )
            return
        if self._active_threads:
            self._info_area_vm.add_message(
                _t(
                    "main_window.task_already_running",
                    "A task is already running. Cancel it before starting another one.",
                ),
                "warning",
            )
            return

        resolved = self._resolve_execution_route(
            file_paths=(file_path,),
            target_format=target_format,
            action_name=action_name,
        )
        if resolved is None:
            return
        target_format, route_choice = resolved

        try:
            request, context = self._build_request(
                file_path=file_path,
                target_format=target_format,
                action_name=action_name,
                options=options,
                route_options=route_choice.options,
            )
        except _OutputPolicyConfigError:
            self._report_output_policy_config_error()
            return
        if not self._confirm_request_admission(request):
            return

        def project_reserved_execution() -> None:
            self._start_time = time.monotonic()
            self._view_model.begin_execution_telemetry(request.request_id, (request.request_id,))
            self._last_request_contexts[_normalize_path(file_path)] = dict(context)
            self._batch_list_vm.set_file_status(
                file_path,
                "processing",
                operation_id=request.request_id,
            )
            self._action_area_vm.show_cancel()
            self._info_area_vm.add_message(
                f"Started: {Path(file_path).name}",
                "info",
                show_location=True,
                file_path=file_path,
                navigate_file_path=file_path,
            )

        self._launch_execution_thread(
            controller=controller,
            request=request,
            context=context,
            project_reserved_execution=project_reserved_execution,
        )

    def _start_batch_execution(
        self,
        *,
        file_paths: Sequence[str],
        target_format: str,
        action_name: str,
        options: dict[str, Any],
    ) -> None:
        if self._execution_close_pending or self._shutdown_finalized:
            return
        controller = self._view_model.controller
        if controller is None or not controller.has_runtime:
            self._info_area_vm.add_message(
                _t("main_window.runtime_unavailable", "Runtime is unavailable; conversion cannot start."),
                "warning",
            )
            return
        if self._active_threads:
            self._info_area_vm.add_message(
                _t(
                    "main_window.task_already_running",
                    "A task is already running. Cancel it before starting another one.",
                ),
                "warning",
            )
            return

        resolved = self._resolve_execution_route(
            file_paths=file_paths,
            target_format=target_format,
            action_name=action_name,
        )
        if resolved is None:
            return
        target_format, route_choice = resolved

        try:
            request, context = self._build_batch_request(
                file_paths=file_paths,
                target_format=target_format,
                action_name=action_name,
                options=options,
                route_options=route_choice.options,
            )
        except _OutputPolicyConfigError:
            self._report_output_policy_config_error()
            return
        if not self._confirm_request_admission(request):
            return
        task_id = request.request_id

        def project_reserved_execution() -> None:
            self._start_time = time.monotonic()
            self._view_model.begin_execution_telemetry(
                task_id,
                tuple(f"{task_id}-{index}" for index, _path in enumerate(context.get("file_paths", []))),
            )
            for path in context.get("file_paths", []):
                per_file_context = dict(context)
                per_file_context["file_path"] = path
                per_file_context["display_name"] = Path(path).name
                self._last_request_contexts[_normalize_path(path)] = per_file_context
                self._batch_list_vm.set_file_status(path, "processing", operation_id=task_id)
            self._action_area_vm.show_cancel()
            self._info_area_vm.add_message(
                f"Started: {context.get('display_name', 'Batch conversion')}",
                "info",
                show_location=False,
                operation_id=task_id,
            )

        self._launch_execution_thread(
            controller=controller,
            request=request,
            context=context,
            project_reserved_execution=project_reserved_execution,
            batch_execution=True,
        )

    def _start_aggregate_execution(
        self,
        *,
        file_paths: Sequence[str],
        target_format: str,
        action_name: str,
        options: dict[str, Any],
    ) -> None:
        if self._execution_close_pending or self._shutdown_finalized:
            return
        controller = self._view_model.controller
        if controller is None or not controller.has_runtime:
            self._info_area_vm.add_message(
                _t("main_window.runtime_unavailable", "Runtime is unavailable; conversion cannot start."),
                "warning",
            )
            return
        if self._active_threads:
            self._info_area_vm.add_message(
                _t(
                    "main_window.task_already_running",
                    "A task is already running. Cancel it before starting another one.",
                ),
                "warning",
            )
            return

        resolved = self._resolve_execution_route(
            file_paths=file_paths,
            target_format=target_format,
            action_name=action_name,
        )
        if resolved is None:
            return
        target_format, route_choice = resolved

        try:
            request, context = self._build_aggregate_request(
                file_paths=file_paths,
                target_format=target_format,
                action_name=action_name,
                options=options,
                route_options=route_choice.options,
            )
        except _OutputPolicyConfigError:
            self._report_output_policy_config_error()
            return
        if not self._confirm_request_admission(request):
            return
        task_id = request.request_id

        def project_reserved_execution() -> None:
            self._start_time = time.monotonic()
            self._view_model.begin_execution_telemetry(task_id, (task_id,))
            for path in context.get("file_paths", []):
                self._last_request_contexts[_normalize_path(path)] = dict(context)
                self._batch_list_vm.set_file_status(path, "processing", operation_id=task_id)
            self._action_area_vm.show_cancel()
            self._info_area_vm.add_message(
                f"Started: {context.get('display_name', action_name)}",
                "info",
                show_location=False,
                operation_id=task_id,
            )

        self._launch_execution_thread(
            controller=controller,
            request=request,
            context=context,
            project_reserved_execution=project_reserved_execution,
            aggregate_action_name=action_name,
        )

    def _launch_execution_thread(
        self,
        *,
        controller: ApplicationController,
        request: ConversionRequest,
        context: dict[str, Any],
        project_reserved_execution: Callable[[], None],
        aggregate_action_name: str = "",
        batch_execution: bool = False,
    ) -> bool:
        """Reserve, start and release one execution, or roll its projection back."""
        task_id = request.request_id
        reservation_missing = object()
        cancellation_reservation: object = reservation_missing
        thread: _ExecutionThread | None = None
        try:
            cancellation_reservation = controller.prepare_execution_cancellation(
                request,
                batch=batch_execution,
            )
            project_reserved_execution()
            thread = _ExecutionThread(
                controller=controller,
                request=request,
                context=context,
                aggregate_action_name=aggregate_action_name,
                batch_execution=batch_execution,
                parent=self,
            )
            thread.result_signal.connect(self._on_execution_finished)
            thread.error_signal.connect(self._on_execution_failed)
            thread.finished.connect(self._on_execution_thread_finished)
            self._active_threads[task_id] = thread
            self._execution_cleanup_by_thread[thread] = (
                task_id,
                controller,
                cancellation_reservation,
            )
            thread.start()
        except Exception as exc:
            if thread is not None and thread.isRunning():
                # A platform binding may report a start error after the native
                # worker has already entered ``run``.  Retain every owner and
                # reservation until ``finished`` rather than orphaning it.
                with contextlib.suppress(Exception):
                    controller.cancel(task_id)
                self._info_area_vm.add_message(
                    _t(
                        "main_window.thread_start_uncertain",
                        "The task worker started but startup reporting failed; cancellation was requested.",
                    ),
                    "warning",
                )
                return True
            self._active_threads.pop(task_id, None)
            if thread is not None:
                self._execution_cleanup_by_thread.pop(thread, None)
            if cancellation_reservation is not reservation_missing:
                with contextlib.suppress(Exception):
                    controller.release_execution_cancellation(task_id, cancellation_reservation)
            if thread is not None and not thread.isRunning():
                thread.deleteLater()
            self._on_execution_failed(str(exc), context)
            return False
        return True

    @Slot()
    def _on_execution_thread_finished(self) -> None:
        """Release one worker on the GUI thread after its result signal was queued."""
        thread = self.sender()
        if not isinstance(thread, QThread):
            return
        cleanup = self._execution_cleanup_by_thread.pop(thread, None)
        if cleanup is None:
            thread.deleteLater()
            return
        task_id, controller, cancellation_reservation = cleanup
        try:
            controller.release_execution_cancellation(task_id, cancellation_reservation)
        except Exception as exc:
            self._info_area_vm.add_message(str(exc), "warning")
        finally:
            self._active_threads.pop(task_id, None)
            thread.deleteLater()

    def _confirm_request_admission(self, request: ConversionRequest) -> bool:
        """Enforce core admission decisions before crossing into runtime."""
        from docwen_core.detection import inspect_file
        from docwen_core.models import (
            FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
            FILE_INSPECTION_METADATA_KEY,
            AdmissionDecision,
            admission_is_satisfied,
            make_admission_acceptance,
        )
        from docwen_gui.dialogs.feedback import confirm

        pending: list[tuple[FileRef, FileInspection]] = []
        for ref in request.input_refs:
            raw_inspection = ref.metadata.get(FILE_INSPECTION_METADATA_KEY)
            if raw_inspection is None:
                self._info_area_vm.add_message(
                    _t("main_window.file_admission_invalid", "File inspection data is invalid."),
                    "danger",
                )
                return False
            if not isinstance(raw_inspection, dict):
                self._info_area_vm.add_message(
                    _t("main_window.file_admission_invalid", "File inspection data is invalid."),
                    "danger",
                )
                return False
            try:
                inspection = inspect_file(ref.path)
            except (FileNotFoundError, OSError, TypeError, ValueError):
                self._info_area_vm.add_message(
                    _t("main_window.file_admission_invalid", "File inspection data is invalid."),
                    "danger",
                )
                return False
            if raw_inspection != inspection.to_dict():
                self._info_area_vm.add_message(
                    _t(
                        "main_window.file_admission_changed",
                        "The file changed after it was added. Remove it from the list and add it again to re-check the file, then retry.",
                    ),
                    "warning",
                    show_location=True,
                    file_path=ref.path,
                    navigate_file_path=ref.path,
                )
                return False
            if inspection.decision is AdmissionDecision.BLOCK:
                self._info_area_vm.add_message(
                    render_file_inspection_message(inspection, prefer_reason=True)
                    or _t("main_window.file_admission_blocked", "The selected file cannot be processed."),
                    "danger",
                    show_location=True,
                    file_path=ref.path,
                    navigate_file_path=ref.path,
                )
                return False
            if not admission_is_satisfied(inspection, ref.metadata):
                pending.append((ref, inspection))

        if not pending:
            return True

        details = "\n\n".join(
            f"{Path(ref.path).name}: {render_file_inspection_message(inspection)}" for ref, inspection in pending
        )
        confirmed = confirm(
            _t("main_window.file_admission_confirm_title", "Confirm detected file format"),
            _t(
                "main_window.file_admission_confirm_message",
                "The filename and detected content differ. Process using the detected format?",
            ),
            danger=True,
            parent=self,
            details=details,
            confirm_label=_t("main_window.file_admission_confirm_action", "Process as detected format"),
        )
        if not confirmed:
            self._info_area_vm.add_message(
                _t("main_window.file_admission_cancelled", "Processing was cancelled; the format was not accepted."),
                "warning",
            )
            return False

        for ref, inspection in pending:
            acceptance = make_admission_acceptance(inspection)
            ref.metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY] = acceptance
            normalized = _normalize_path(ref.path)
            for source_ref in self._view_model.files:
                if _normalize_path(getattr(source_ref, "path", "")) == normalized:
                    source_ref.metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY] = dict(acceptance)
            entry = self._batch_list_vm.get_file_entry(ref.path)
            if entry is not None:
                entry.metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY] = dict(acceptance)
        return True

    def _request_file_ref(
        self,
        source_path: str,
    ) -> FileRef:
        """Build a runtime ref without discarding the ingress inspection.

        Routing may normalize the runtime format/category (notably TXT to the
        Markdown workflow), but warning and inspection facts must remain
        attached so application/runtime admission can enforce the same
        decision without opening and guessing the file again.
        """
        from docwen_core.models.file_ref import FileRef

        normalized = _normalize_path(source_path)
        source_ref = next(
            (ref for ref in self._view_model.files if _normalize_path(getattr(ref, "path", "")) == normalized),
            None,
        )
        if source_ref is not None:
            return FileRef(
                path=source_path,
                format=source_ref.format,
                category=source_ref.category,
                encoding=source_ref.encoding,
                warning_message=source_ref.warning_message,
                size_bytes=source_ref.size_bytes,
                metadata=dict(source_ref.metadata),
            )

        entry = self._batch_list_vm.get_file_entry(source_path)
        if entry is not None:
            return FileRef(
                path=source_path,
                format=entry.detected_format,
                category=entry.workflow_category,
                warning_message=entry.warning_message or "",
                size_bytes=entry.size_bytes,
                metadata=dict(entry.metadata),
            )

        # Programmatic callers that bypass the visual list still cross the
        # same Core admission boundary here; no suffix-derived FileRef is ever
        # manufactured.
        from docwen_core.detection import FileAdmissionError, inspect_file
        from docwen_core.detection.ooxml_signature import OOXML_SIGNATURE_INFO_METADATA_KEY
        from docwen_core.models import FILE_INSPECTION_METADATA_KEY

        inspection = inspect_file(source_path)
        if not inspection.may_execute:
            raise FileAdmissionError(inspection)
        return FileRef(
            path=inspection.file_path,
            format=inspection.detected_format,
            category=inspection.workflow_category,
            warning_message=render_file_inspection_message(inspection),
            size_bytes=inspection.size_bytes,
            metadata={
                FILE_INSPECTION_METADATA_KEY: inspection.to_dict(),
                OOXML_SIGNATURE_INFO_METADATA_KEY: dict(inspection.ooxml_signature),
            },
        )

    def _build_request(
        self,
        *,
        file_path: str,
        target_format: str,
        action_name: str,
        options: dict[str, Any],
        route_options: Sequence[str] | None = None,
    ) -> tuple[ConversionRequest, dict[str, Any]]:
        from docwen_core.models.request import ConversionRequest

        fmt, category = _resolve_file_context(self._file_contexts, self._batch_list_vm, file_path) or (
            "unknown",
            "other",
        )
        request_options = self._merge_template_options_for_request(
            target_format,
            options,
            source_format=fmt,
            source_category=category,
            action_name=action_name,
        )
        request_options = _normalize_proofread_action_options(request_options, action_name=action_name)
        request_options = _to_markdown_locale_options(
            request_options,
            target_format=target_format,
            action_name=action_name,
            route_options=route_options,
        )
        request_options = _route_scoped_options(
            request_options,
            route_options=route_options,
        )
        source_path = str(Path(file_path))

        output_policy = self._build_output_policy()
        request_id = str(uuid.uuid4())
        request = ConversionRequest(
            request_id=request_id,
            input_refs=[self._request_file_ref(source_path)],
            target_format=target_format,
            action_name=action_name,
            options=request_options,
            output_policy=output_policy,
        )
        context = {
            "request_id": request_id,
            "file_path": _normalize_path(source_path),
            "display_name": Path(source_path).name,
            "target_format": target_format,
            "action_name": action_name,
            "options": _redacted_request_options(request_options),
            "open_after_done": output_policy.open_after_done,
        }
        return request, context

    def _build_batch_request(
        self,
        *,
        file_paths: Sequence[str],
        target_format: str,
        action_name: str,
        options: dict[str, Any],
        route_options: Sequence[str] | None = None,
    ) -> tuple[ConversionRequest, dict[str, Any]]:
        from docwen_core.models.request import ConversionRequest

        input_refs: list[FileRef] = []
        normalized_paths: list[str] = []
        first_source_format = "unknown"
        first_source_category = "other"
        for index, file_path in enumerate(file_paths):
            fmt, category = _resolve_file_context(self._file_contexts, self._batch_list_vm, file_path) or (
                "unknown",
                "other",
            )
            if index == 0:
                first_source_format = fmt
                first_source_category = category
            source_path = str(Path(file_path))
            normalized_paths.append(_normalize_path(source_path))
            input_refs.append(self._request_file_ref(source_path))

        request_options = self._merge_template_options_for_request(
            target_format,
            options,
            source_format=first_source_format,
            source_category=first_source_category,
            action_name=action_name,
        )
        request_options = _normalize_proofread_action_options(request_options, action_name=action_name)
        request_options = _to_markdown_locale_options(
            request_options,
            target_format=target_format,
            action_name=action_name,
            route_options=route_options,
        )
        request_options = _route_scoped_options(
            request_options,
            route_options=route_options,
        )
        output_policy = self._build_output_policy()
        request_id = str(uuid.uuid4())
        request = ConversionRequest(
            request_id=request_id,
            input_refs=input_refs,
            target_format=target_format,
            action_name=action_name,
            options=request_options,
            output_policy=output_policy,
        )
        display_name = f"Batch conversion ({len(input_refs)} files)"
        context = {
            "request_id": request_id,
            "file_path": normalized_paths[0] if normalized_paths else "",
            "file_paths": normalized_paths,
            "display_name": display_name,
            "target_format": target_format,
            "action_name": action_name,
            "options": _redacted_request_options(request_options),
            "total_count": len(input_refs),
            "batch": True,
            "open_after_done": output_policy.open_after_done,
        }
        return request, context

    def _build_output_policy(self) -> OutputPolicy:
        from docwen_core.models.request import OutputPolicy

        controller = self._view_model.controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        if cfg_port is None:
            return OutputPolicy()

        try:
            mode = str(cfg_port.get("output.directory.mode", "source") or "source")
            custom_path = str(cfg_port.get("output.directory.custom_path", "") or "").strip()
            create_date_subfolder = bool(cfg_port.get("output.directory.create_date_subfolder", False))
            date_folder_format = str(cfg_port.get("output.directory.date_folder_format", "%Y-%m-%d") or "%Y-%m-%d")
            auto_open_folder = bool(cfg_port.get("output.behavior.auto_open_folder", False))
        except Exception as exc:
            logger.exception("Unable to read persisted output settings")
            raise _OutputPolicyConfigError("Persisted output settings are unavailable") from exc

        try:
            output_dir = (
                str(Path(custom_path).expanduser().resolve(strict=False)) if mode == "custom" and custom_path else None
            )
        except (OSError, RuntimeError, ValueError) as exc:
            logger.exception("Unable to resolve persisted custom output path")
            raise _OutputPolicyConfigError("Persisted custom output path is invalid") from exc
        date_subfolder = _output_date_subfolder_token(date_folder_format) if create_date_subfolder else ""
        return OutputPolicy(
            output_dir=output_dir,
            date_subfolder=date_subfolder,
            overwrite_mode="rename",
            open_after_done=auto_open_folder,
        )

    def _report_output_policy_config_error(self) -> None:
        self._info_area_vm.add_message(
            _t(
                "main_window.output_settings_unavailable",
                "Output settings could not be loaded; the request was not started.",
            ),
            "error",
        )

    def _build_aggregate_request(
        self,
        *,
        file_paths: Sequence[str],
        target_format: str,
        action_name: str,
        options: dict[str, Any],
        route_options: Sequence[str] | None = None,
    ) -> tuple[ConversionRequest, dict[str, Any]]:
        from docwen_core.models.request import ConversionRequest

        input_refs: list[FileRef] = []
        normalized_paths: list[str] = []
        for file_path in file_paths:
            source_path = str(Path(file_path))
            normalized_paths.append(_normalize_path(source_path))
            input_refs.append(self._request_file_ref(source_path))

        output_policy = self._build_output_policy()
        request_id = str(uuid.uuid4())
        request = ConversionRequest(
            request_id=request_id,
            input_refs=input_refs,
            target_format=target_format,
            action_name=action_name,
            options=_route_scoped_options(options, route_options=route_options),
            output_policy=output_policy,
        )
        display_name = f"{action_name} ({len(input_refs)} files)"
        context = {
            "request_id": request_id,
            "file_path": normalized_paths[0] if normalized_paths else "",
            "file_paths": normalized_paths,
            "display_name": display_name,
            "target_format": target_format,
            "action_name": action_name,
            "options": _redacted_request_options(_route_scoped_options(options, route_options=route_options)),
            "total_count": len(input_refs),
            "aggregate": True,
            "open_after_done": output_policy.open_after_done,
        }
        return request, context

    @Slot(object, dict)
    def _on_execution_finished(self, result: object, context: dict[str, Any]) -> None:
        from docwen_core.models.result import ConversionResult

        task_id = context.get("request_id", "")
        file_path = context.get("file_path", "")
        file_paths = list(context.get("file_paths", []) or ([file_path] if file_path else []))
        total_count = int(context.get("total_count", len(file_paths) or 1))
        self._action_area_vm.hide_cancel()

        if context.get("batch"):
            if not isinstance(result, list):
                self._on_execution_failed("Invalid batch conversion result", context)
                return
            self._on_batch_execution_finished(result, context)
            return

        if not isinstance(result, ConversionResult):
            self._on_execution_failed("Invalid conversion result", context)
            return

        if result.success:
            output_path = self._pick_output_path(result)
            warning_messages = _result_warning_messages(result)
            completion_tone = "warning" if warning_messages else "success"
            for path in file_paths:
                self._batch_list_vm.set_file_status(
                    path,
                    "completed",
                    output_path=output_path,
                    operation_id=task_id,
                    error_message="",
                )
            self._info_area_vm.add_message(
                _t("info_area.history_completed", name=context.get("display_name", Path(file_path).name)),
                completion_tone,
                show_location=bool(output_path),
                file_path=output_path or file_path,
                navigate_file_path=output_path or "",
                operation_id=task_id,
            )
            for warning_message in warning_messages:
                self._info_area_vm.add_message(
                    warning_message,
                    "warning",
                    show_location=bool(output_path),
                    file_path=output_path or file_path,
                    navigate_file_path=output_path or "",
                    operation_id=task_id,
                )
            output_dir = str(Path(output_path).parent) if output_path else str(Path(file_path).parent)
            guide_actions = self._info_area_vm.compute_guide_actions(
                "success",
                output_dir=output_dir,
            )
            self._info_area_vm.set_task_summary(
                operation_id=task_id,
                current_file=context.get("display_name", Path(file_path).name),
                current_file_path=file_path,
                completed_count=total_count,
                total_count=total_count,
                failed_count=0,
                state="success",
                tone=completion_tone,
                navigate_file_path=output_path or "",
                navigation_kind="output",
                guide_actions=guide_actions,
            )
            self._publish_execution_summary("completed")
            if context.get("open_after_done") and (output_path or output_dir):
                self._open_path(output_path or output_dir, open_parent=bool(output_path))
            self._maybe_notify_task_completion(context)
        else:
            self._handle_unsuccessful_result(result, context)
            result_error = result.error
            cancelled = bool(result_error is not None and result_error.error_type == "cancelled")
            self._publish_execution_summary(
                "cancelled" if cancelled else "failed",
                message=result_error.message if result_error is not None else "Conversion failed",
            )
            self._maybe_notify_task_completion(context)

    @Slot(str, dict)
    def _on_execution_failed(self, error_message: str, context: dict[str, Any]) -> None:
        task_id = context.get("request_id", "")
        file_path = context.get("file_path", "")
        file_paths = list(context.get("file_paths", []) or ([file_path] if file_path else []))
        total_count = int(context.get("total_count", len(file_paths) or 1))
        self._action_area_vm.hide_cancel()
        message = error_message or "Conversion failed"
        for path in file_paths:
            self._batch_list_vm.set_file_status(
                path,
                "failed",
                error_message=message,
                operation_id=task_id,
            )
        self._info_area_vm.add_message(
            message,
            "danger",
            show_location=True,
            file_path=file_path,
            navigate_file_path=file_path,
            operation_id=task_id,
        )
        guide_actions = self._info_area_vm.compute_guide_actions(
            "failed",
            failed_details_path=file_path,
            retry_available=True,
        )
        self._info_area_vm.set_task_summary(
            operation_id=task_id,
            current_file=context.get("display_name", Path(file_path).name),
            current_file_path=file_path,
            completed_count=0,
            total_count=total_count,
            failed_count=total_count,
            state="failed",
            tone="danger",
            navigate_file_path=file_path,
            navigation_kind="failed",
            guide_actions=guide_actions,
        )
        self._publish_execution_summary("failed", message=message)
        self._maybe_notify_task_completion(context)

    def _handle_unsuccessful_result(self, result: ConversionResult, context: dict[str, Any]) -> None:
        error = result.error
        message = error.message if error is not None else "Conversion failed"
        cancelled = bool(error is not None and error.error_type == "cancelled")
        state = "cancelled" if cancelled else "failed"
        tone = "warning" if cancelled else "danger"
        task_id = context.get("request_id", "")
        file_path = context.get("file_path", "")
        file_paths = list(context.get("file_paths", []) or ([file_path] if file_path else []))
        total_count = int(context.get("total_count", len(file_paths) or 1))
        retained_output_path = "" if cancelled else self._pick_existing_output_path(result)

        entry_status = "cancelled" if cancelled else "failed"
        for path in file_paths:
            self._batch_list_vm.set_file_status(
                path,
                entry_status,
                output_path=retained_output_path,
                error_message=message,
                operation_id=task_id,
            )
        history_path = retained_output_path or file_path
        self._info_area_vm.add_message(
            message,
            tone,
            show_location=True,
            file_path=history_path,
            navigate_file_path=history_path,
            operation_id=task_id,
        )
        guide_actions = self._info_area_vm.compute_guide_actions(
            state,
            output_dir=str(Path(retained_output_path).parent) if retained_output_path else "",
            failed_details_path=file_path,
            retry_available=not cancelled,
        )
        self._info_area_vm.set_task_summary(
            operation_id=task_id,
            current_file=context.get("display_name", Path(file_path).name),
            current_file_path=file_path,
            completed_count=0,
            total_count=total_count,
            failed_count=0 if cancelled else total_count,
            cancelled_count=total_count if cancelled else 0,
            state=state,
            tone=tone,
            navigate_file_path=file_path,
            navigation_kind="failed",
            guide_actions=guide_actions,
        )

    def _on_batch_execution_finished(self, results: list[object], context: dict[str, Any]) -> None:
        from docwen_core.models.result import ConversionResult

        task_id = context.get("request_id", "")
        file_paths = list(context.get("file_paths", []))
        total_count = int(context.get("total_count", len(file_paths) or len(results)))
        success_count = 0
        failed_count = 0
        skipped_count = 0
        cancelled_count = 0
        output_paths: list[str] = []
        retained_failure_paths: list[str] = []
        warning_rows: list[tuple[str, str, str]] = []
        first_failed_path = ""
        first_error_message = ""
        first_error_output = ""
        first_retained_failure: tuple[str, str, str] | None = None

        for index, file_path in enumerate(file_paths):
            raw_result = results[index] if index < len(results) else None
            if not isinstance(raw_result, ConversionResult):
                failed_count += 1
                message = "Invalid batch conversion result"
                if not first_failed_path:
                    first_failed_path = file_path
                    first_error_message = message
                self._batch_list_vm.set_file_status(
                    file_path,
                    "failed",
                    output_path="",
                    error_message=message,
                    operation_id=task_id,
                )
                continue

            if raw_result.success:
                success_count += 1
                output_path = self._pick_output_path(raw_result)
                if output_path:
                    output_paths.append(output_path)
                for warning_message in _result_warning_messages(raw_result):
                    warning_rows.append((file_path, output_path, warning_message))
                self._batch_list_vm.set_file_status(
                    file_path,
                    "completed",
                    output_path=output_path,
                    error_message="",
                    operation_id=task_id,
                )
                continue

            error = raw_result.error
            error_type = getattr(error, "error_type", "") if error is not None else ""
            message = error.message if error is not None else "Conversion failed"
            if error_type == "cancelled":
                cancelled_count += 1
                self._batch_list_vm.set_file_status(
                    file_path,
                    "cancelled",
                    output_path="",
                    error_message=message,
                    operation_id=task_id,
                )
            elif error_type == "skipped":
                skipped_count += 1
                self._batch_list_vm.set_file_status(
                    file_path,
                    "skipped",
                    output_path="",
                    skip_reason=message,
                    error_message="",
                    operation_id=task_id,
                )
            else:
                failed_count += 1
                retained_output_path = self._pick_existing_output_path(raw_result)
                if retained_output_path:
                    retained_failure_paths.append(retained_output_path)
                self._batch_list_vm.set_file_status(
                    file_path,
                    "failed",
                    output_path=retained_output_path,
                    error_message=message,
                    operation_id=task_id,
                )
                if retained_output_path and first_retained_failure is None:
                    first_retained_failure = (file_path, retained_output_path, message)
                if not first_failed_path:
                    first_failed_path = file_path
                    first_error_message = message
                    first_error_output = retained_output_path

        completed_count = success_count + failed_count
        if cancelled_count and not failed_count:
            state = "cancelled"
            tone = "warning"
        elif failed_count or skipped_count:
            state = "partial" if success_count else "failed"
            tone = "warning" if success_count else "danger"
        else:
            state = "success"
            tone = "warning" if warning_rows else "success"

        successful_output_dir = str(Path(output_paths[0]).parent) if output_paths else ""
        retained_failure_output_dir = str(Path(retained_failure_paths[0]).parent) if retained_failure_paths else ""
        guide_output_dir = successful_output_dir or retained_failure_output_dir
        navigate_path = output_paths[0] if state == "success" and output_paths else first_failed_path
        navigation_kind = "output" if state == "success" else ("failed" if first_failed_path else "")
        guide_actions = self._info_area_vm.compute_guide_actions(
            state,
            output_dir=guide_output_dir,
            failed_details_path=first_failed_path,
            retry_available=bool(first_failed_path),
        )
        self._info_area_vm.add_message(
            _t(
                "components.info_area.batch_completed",
                "Batch finished: {success} succeeded, {failed} failed, {skipped} skipped, {cancelled} cancelled",
                success=success_count,
                failed=failed_count,
                skipped=skipped_count,
                cancelled=cancelled_count,
            ),
            tone,
            show_location=bool(guide_output_dir),
            file_path=guide_output_dir,
            navigate_file_path=guide_output_dir,
            operation_id=task_id,
        )
        for warning_file, warning_output, warning_message in warning_rows:
            self._info_area_vm.add_message(
                f"{Path(warning_file).name}: {warning_message}",
                "warning",
                show_location=bool(warning_output),
                file_path=warning_output or warning_file,
                navigate_file_path=warning_output or "",
                operation_id=task_id,
            )
        if first_error_message:
            history_path = first_error_output or first_failed_path
            self._info_area_vm.add_message(
                first_error_message,
                "danger",
                show_location=True,
                file_path=history_path,
                navigate_file_path=history_path,
                operation_id=task_id,
            )
        if first_retained_failure is not None and first_retained_failure[0] != first_failed_path:
            _, retained_output, retained_message = first_retained_failure
            self._info_area_vm.add_message(
                retained_message,
                "danger",
                show_location=True,
                file_path=retained_output,
                navigate_file_path=retained_output,
                operation_id=task_id,
            )
        self._info_area_vm.set_task_summary(
            operation_id=task_id,
            current_file=context.get("display_name", "Batch conversion"),
            current_file_path=context.get("file_path", ""),
            completed_count=completed_count,
            total_count=total_count,
            failed_count=failed_count,
            skipped_count=skipped_count,
            cancelled_count=cancelled_count,
            state=state,
            tone=tone,
            navigate_file_path=navigate_path,
            navigation_kind=navigation_kind,
            guide_actions=guide_actions,
        )
        self._publish_execution_summary("completed" if state == "success" else state, message=first_error_message)
        if context.get("open_after_done") and successful_output_dir:
            self._open_path(successful_output_dir, open_parent=False)
        self._maybe_notify_task_completion(context)

    @staticmethod
    def _pick_output_path(result: ConversionResult) -> str:
        primary = next((artifact.staging_path for artifact in result.artifacts if artifact.is_primary), "")
        if primary:
            return primary
        if result.artifacts:
            return result.artifacts[0].staging_path
        return ""

    @staticmethod
    def _pick_existing_output_path(result: ConversionResult) -> str:
        """Return a real retained artifact path suitable for failure navigation."""
        primary = [artifact for artifact in result.artifacts if artifact.is_primary]
        secondary = [artifact for artifact in result.artifacts if not artifact.is_primary]
        for artifact in (*primary, *secondary):
            if not artifact.staging_path:
                continue
            try:
                if filesystem_path(artifact.staging_path).is_file():
                    return artifact.staging_path
            except (OSError, ValueError):
                continue
        return ""

    def _cancel_active_task(self) -> None:
        controller = self._view_model.controller
        if controller is None or not controller.has_runtime:
            return
        active_parent_ids = list(self._active_threads)
        task_id = active_parent_ids[0] if active_parent_ids else self._view_model.current_task_id
        if not task_id:
            return
        try:
            controller.cancel(task_id)
            self._info_area_vm.set_transient_message(
                f"processing:{task_id}",
                _t("main_window.cancelling", "Cancelling..."),
                "warning",
                ttl_ms=3000,
                source=task_id,
            )
        except Exception as exc:
            self._info_area_vm.add_message(str(exc), "warning")

    def _publish_execution_summary(self, status: str, *, message: str = "") -> None:
        """Expose the already-committed InfoArea summary to GUI observers."""
        summary = self._info_area_vm.task_summary
        self._view_model.publish_execution_summary(
            status,
            {
                "task_id": summary.operation_id,
                "state": summary.state,
                "completed_count": summary.completed_count,
                "total_count": summary.total_count,
                "failed_count": summary.failed_count,
                "skipped_count": summary.skipped_count,
                "cancelled_count": summary.cancelled_count,
                "message": message,
            },
        )

    # ── Batch-list and info-area actions ────────────────────────────

    def _handle_batch_entry_action(self, action_key: str, file_path: str) -> None:
        if action_key == "open_output":
            entry = self._batch_list_vm.get_file_entry(file_path)
            if entry is not None and entry.output_path:
                self._open_path(entry.output_path)
            return
        if action_key == "open_source_location":
            self._open_path(file_path, open_parent=True)
            return
        if action_key == "show_error_details":
            entry = self._batch_list_vm.get_file_entry(file_path)
            message = entry.error_message if entry is not None else ""
            if message:
                self._info_area_vm.add_message(message, "danger")
            return
        if action_key == "copy_error_details":
            copied = self._batch_list.copy_error_details(file_path)
            self._info_area_vm.add_message(
                _t("main_window.error_copied", "Error details copied.")
                if copied
                else _t("main_window.no_error_available", "No error details available."),
                "success" if copied else "warning",
            )

    def _handle_navigation_request(self, target_path: str) -> None:
        self._open_path(target_path, open_parent=True)

    def _open_location(self, file_path: str) -> None:
        self._open_path(file_path, open_parent=True)

    def _handle_task_guide_action(self, action_key: str, target_path: str) -> None:
        if action_key == "open_output_dir":
            self._open_path(target_path, open_parent=False)
            return
        if action_key == "view_failed_details":
            if target_path:
                self._handle_batch_entry_action("show_error_details", target_path)
            return
        if action_key == "retry_failed":
            self._retry_failed_request()
            return
        if action_key == "add_more_files":
            self._input_area.add_button.setFocus(Qt.FocusReason.ShortcutFocusReason)
            self.bring_to_front()

    def _retry_failed_request(self) -> None:
        failed_files = self._batch_list_vm.get_failed_files()
        if not failed_files:
            self._info_area_vm.add_message(
                _t("main_window.no_failed_to_retry", "No failed task is available to retry."),
                "warning",
            )
            return
        file_path = failed_files[0]
        context = self._last_request_contexts.get(_normalize_path(file_path))
        if context is None:
            self._info_area_vm.add_message(
                _t("main_window.retry_context_unavailable", "Retry context is unavailable."),
                "warning",
            )
            return
        self._batch_list_vm.reset_failed_files([file_path])
        self._start_execution(
            file_path=context["file_path"],
            target_format=context["target_format"],
            action_name=context["action_name"],
            options=dict(context["options"]),
        )

    # ── Path helpers ────────────────────────────────────────────────

    def _open_path(self, target_path: str, *, open_parent: bool = False) -> bool:
        if not target_path:
            return False

        result = path_actions.reveal_path(target_path) if open_parent else path_actions.open_path(target_path)
        if result.success:
            return True
        if result.error_code == "missing_path":
            self._info_area_vm.add_message(
                _t("main_window.path_not_exist", "Path does not exist: {path}", path=target_path),
                "warning",
            )
            return False
        self._info_area_vm.add_message(
            _t("main_window.open_path_failed", "Failed to open path: {path}", path=target_path),
            "warning",
        )
        return False

    # ── About dialog ───────────────────────────────────────────────

    def _show_about_dialog(self) -> None:
        """Open the About dialog."""
        from .dialogs.about import AboutDialog

        dialog = AboutDialog(parent=self)
        dialog.show_dialog()

    # ── Window opacity ─────────────────────────────────────────────

    def _apply_window_opacity(self) -> None:
        """Apply window opacity from config (gui.transparency)."""
        controller = self._view_model.controller
        if controller is None:
            return
        cfg_port = getattr(controller, "config_port", None)
        if cfg_port is None:
            return
        try:
            enabled = cfg_port.get("gui.transparency.enabled", False)
            if enabled:
                default_value = cfg_port.get("gui.transparency.default_value", 1.0)
                opacity = max(0.1, min(1.0, float(default_value)))
                self.setWindowOpacity(opacity)
        except Exception:
            pass

    # ── App icon ────────────────────────────────────────────────────

    def _initialize_app_icon(self) -> None:
        """Set the application window icon from bundle assets."""
        try:
            from .resources import initialize_app_icon

            initialize_app_icon(self)
        except Exception:
            pass

    # ── Settings dialog ────────────────────────────────────────────

    _PUBLIC_SETTINGS_SECTIONS = ("proofread",)

    def supported_settings_sections(self) -> tuple[str, ...]:
        """Return semantic settings sections exposed through runtime/control."""

        return self._PUBLIC_SETTINGS_SECTIONS

    def _open_settings_dialog(self) -> None:
        self.open_settings(None)

    def open_settings(self, section: str | None, *, deadline: float | None = None) -> dict[str, Any]:
        """Open or focus the owned non-blocking modal settings dialog."""

        from .view_models.settings_vm import SettingsViewModel
        from .widgets.settings.dialog import SettingsDialog

        if section is not None and section not in self._PUBLIC_SETTINGS_SECTIONS:
            return {"accepted": False, "section": section, "reused": False}

        dialog = self._settings_dialog
        reused = dialog is not None
        if self._settings_request_expired(deadline):
            return self._settings_timeout_result(section, reused=reused)

        previous_section: str | None = None
        if dialog is not None:
            current_section = getattr(dialog, "current_section", None)
            if callable(current_section):
                raw_previous_section = current_section()
                if isinstance(raw_previous_section, str):
                    previous_section = raw_previous_section

        if dialog is None:
            vm = SettingsViewModel(controller=self._view_model.controller)

            # Explicit semantic control takes precedence over context-sensitive UI entry.
            initial_tab = section or self._determine_optimal_settings_tab()
            if initial_tab:
                vm.set_initial_tab(initial_tab)

            dialog = SettingsDialog(parent=self, view_model=vm)
            if self._settings_request_expired(deadline):
                dialog.deleteLater()
                return self._settings_timeout_result(section, reused=False)
            if section is not None and not dialog.activate_section(section):
                dialog.deleteLater()
                return {"accepted": False, "section": section, "reused": False}
            if self._settings_request_expired(deadline):
                dialog.deleteLater()
                return self._settings_timeout_result(section, reused=False)
            dialog.settings_source_changed.connect(self._apply_runtime_window_settings)
            dialog.destroyed.connect(lambda _object=None, owned=dialog: self._clear_settings_dialog(owned))
            dialog.finished.connect(lambda _result=0, owned=dialog: self._clear_settings_dialog(owned))
            self._settings_dialog = dialog
        elif section is not None and not dialog.activate_section(section):
            return {"accepted": False, "section": section, "reused": True}

        if self._settings_request_expired(deadline):
            if reused and previous_section is not None:
                dialog.activate_section(previous_section)
            elif not reused:
                self._clear_settings_dialog(dialog)
                dialog.deleteLater()
            return self._settings_timeout_result(section, reused=reused)

        self.bring_to_front()
        if self._settings_request_expired(deadline):
            if reused and previous_section is not None:
                dialog.activate_section(previous_section)
            elif not reused:
                self._clear_settings_dialog(dialog)
                dialog.deleteLater()
            return self._settings_timeout_result(section, reused=reused)
        dialog.open()
        dialog.raise_()
        dialog.activateWindow()
        return {"accepted": True, "section": section, "reused": reused}

    @staticmethod
    def _settings_request_expired(deadline: float | None) -> bool:
        return deadline is not None and time.monotonic() >= deadline

    @staticmethod
    def _settings_timeout_result(section: str | None, *, reused: bool) -> dict[str, Any]:
        return {
            "accepted": False,
            "section": section,
            "reused": reused,
            "error_code": "control_timeout",
        }

    def _clear_settings_dialog(self, dialog: Any) -> None:
        if self._settings_dialog is dialog:
            self._settings_dialog = None

    # ── Optimal settings tab (category → settings tab) ───────────────

    # Maps file categories to the most relevant Settings dialog tab key.
    _CATEGORY_TO_SETTINGS_TAB: dict[str, str] = {  # noqa: RUF012
        "text": "text",
        "document": "document",
        "spreadsheet": "spreadsheet",
        "image": "image",
        "layout": "layout",
    }

    def _determine_optimal_settings_tab(self) -> str | None:
        """Inspect the current batch file list and return the settings
        tab key that best matches the user's working context.

        Returns *None* when no files are present (the dialog falls back
        to the first tab — general).
        """
        file_paths = self._batch_list_vm.get_files()
        if not file_paths:
            return None

        # Count categories across all files currently in the batch list.
        from collections import Counter

        cat_counts: Counter[str] = Counter()
        for fp in file_paths:
            cat = self._batch_list_vm.get_file_display_category(fp) or "other"
            cat_counts[cat] += 1

        if not cat_counts:
            return None

        # Pick the most common category, breaking ties by priority order.
        max_count = max(cat_counts.values())
        priority = ["text", "spreadsheet", "document", "image", "layout", "other"]
        best_category = "other"
        for cat in priority:
            if cat_counts.get(cat, 0) == max_count:
                best_category = cat
                break

        return self._CATEGORY_TO_SETTINGS_TAB.get(best_category)

    # ── Misc helpers ────────────────────────────────────────────────

    def _trigger_primary_action(self) -> None:
        if self._action_area.isVisible():
            self._action_area.trigger_primary_action()

    # ── Font size presets (M-2) ─────────────────────────────────────

    _FONT_SIZE_PRESETS: ClassVar[dict[str, int]] = FONT_SIZE_PRESETS

    _FONT_PRESET_LABELS: ClassVar[dict[str, str]] = {
        "small": "Small",
        "default": "Default",
        "large": "Large",
        "xlarge": "XLarge",
    }

    def _show_font_size_menu(self) -> None:
        """Show a popup menu with font-size presets below the font button."""
        menu = QMenu(self)
        group = QActionGroup(menu)
        group.setExclusive(True)
        current = self._font_size_preset

        for preset in ("small", "default", "large", "xlarge"):
            action = QAction(
                _t(
                    f"components.font_size.{preset}",
                    self._FONT_PRESET_LABELS.get(preset, preset),
                ),
                menu,
            )
            action.setCheckable(True)
            action.setChecked(preset == current)
            action.setData(preset)
            action.triggered.connect(lambda checked=False, v=preset: self._apply_font_size_preset(v))
            group.addAction(action)
            menu.addAction(action)

        anchor = self._font_size_btn
        menu.exec(anchor.mapToGlobal(anchor.rect().bottomLeft()))

    def _apply_font_size_preset(self, preset: str, *, persist: bool = True) -> None:
        """Set the application-wide font size for the given preset."""
        normalized = normalize_font_size_preset(preset)

        self._font_size_preset = normalized

        ThemeManager.get_instance().apply_font_size_preset(normalized)

        if persist:
            controller = self._view_model.controller
            cfg_port = getattr(controller, "config_port", None) if controller is not None else None
            if cfg_port is not None:
                cfg_port.set("gui.font.size_preset", normalized)

        label = _t(
            f"components.font_size.{normalized}",
            self._FONT_PRESET_LABELS.get(normalized, normalized),
        )
        self._info_area_vm.set_transient_message(
            "font_size",
            _t("info_area.font_size_applied", "Font size: {label}", label=label),
            "info",
            ttl_ms=2500,
            source="font_size",
        )

    # ── Always-on-top toggle (M-3) ──────────────────────────────────

    def is_always_on_top_enabled(self) -> bool:
        return self._always_on_top_enabled

    def set_window_always_on_top(self, enabled: bool) -> None:
        desired = bool(enabled)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, desired)
        self._always_on_top_enabled = desired
        self.setProperty("windowAlwaysOnTop", desired)

        if self.isVisible():
            self.show()
            self.raise_()
            self.activateWindow()

        message = (
            _t("info_area.window_topmost_enabled", "Window always-on-top enabled")
            if desired
            else _t("info_area.window_topmost_disabled", "Window always-on-top disabled")
        )
        tone = "info" if desired else "secondary"
        self._info_area_vm.set_transient_message(
            "window_topmost",
            message,
            tone,
            ttl_ms=2500,
            source="window_topmost",
        )

    def toggle_always_on_top(self) -> None:
        self.set_window_always_on_top(not self._always_on_top_enabled)

    def _begin_execution_close(self) -> None:
        """Cancel all owned executions and keep the window alive until they finish."""
        if self._execution_close_pending:
            return

        self._execution_close_pending = True
        self._execution_drain_timed_out = False
        self._execution_drain_deadline = time.monotonic() + self._EXECUTION_DRAIN_TIMEOUT_SECONDS
        self.setEnabled(False)
        self._info_area_vm.add_message(
            _t(
                "main_window.close_waiting_for_tasks",
                "Cancelling active tasks before closing...",
            ),
            "warning",
        )

        controller = self._view_model.controller
        for task_id in tuple(self._active_threads):
            try:
                if controller is None:
                    raise RuntimeError(
                        _t(
                            "main_window.runtime_unavailable",
                            "Runtime is unavailable; conversion cannot start.",
                        )
                    )
                controller.cancel(task_id)
            except Exception as exc:
                self._info_area_vm.add_message(
                    _t(
                        "main_window.close_cancel_failed",
                        "Could not request cancellation for {task_id}: {message}",
                        task_id=task_id,
                        message=str(exc),
                    ),
                    "warning",
                )

        timer = self._execution_drain_timer
        if timer is None:
            timer = QTimer(self)
            timer.setInterval(self._EXECUTION_DRAIN_POLL_MS)
            timer.timeout.connect(self._poll_execution_close)
            self._execution_drain_timer = timer
        timer.start()
        QTimer.singleShot(0, self._poll_execution_close)

    def _poll_execution_close(self) -> None:
        """Re-enter close only after every parented worker emitted ``finished``."""
        timer = self._execution_drain_timer
        if self._shutdown_finalized or not self._execution_close_pending:
            if timer is not None:
                timer.stop()
            return
        if not self._active_threads:
            if timer is not None:
                timer.stop()
            QTimer.singleShot(0, self.close)
            return
        if self._execution_drain_timed_out or time.monotonic() < self._execution_drain_deadline:
            return

        self._execution_drain_timed_out = True
        self._info_area_vm.add_message(
            _t(
                "main_window.close_wait_timeout",
                "A task is still stopping. DocWen will remain open until cleanup finishes.",
            ),
            "warning",
        )

    def _finalize_shutdown(self) -> None:
        """Stop shared services exactly once, after execution ownership is drained."""
        if self._shutdown_finalized:
            return
        self._shutdown_finalized = True
        self._execution_close_pending = False
        if self._execution_drain_timer is not None:
            self._execution_drain_timer.stop()
        with contextlib.suppress(Exception):
            self._task_event_bridge.stop_auto_flush()
        controller = self._view_model.controller
        if controller is not None:
            with contextlib.suppress(Exception):
                controller.stop()

    # ── Public API ─────────────────────────────────────────────────

    @property
    def view_model(self) -> MainWindowViewModel:
        return self._view_model

    @property
    def task_event_bridge(self) -> TaskEventBridge | None:
        return self._task_event_bridge

    @property
    def input_area(self):
        return self._input_area

    @property
    def batch_list(self):
        return self._batch_list

    @property
    def conversion_panel(self):
        return self._conversion_panel

    @property
    def action_area(self):
        return self._action_area

    @property
    def info_area(self):
        return self._info_area

    @property
    def settings_button(self) -> QToolButton:
        return self._settings_btn

    def bring_to_front(self) -> None:
        previous_state = self.windowState()
        if not self.isVisible():
            self.show()
        if previous_state & Qt.WindowState.WindowMinimized:
            self.setWindowState(previous_state & ~Qt.WindowState.WindowMinimized)
            self.show()
        self.raise_()
        self.activateWindow()

    def handle_ipc_command(self, action: str, file_path: str | None = None) -> None:
        self._view_model.handle_ipc_command(action, file_path)

    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        if not self._window_state_needs_shown_restore:
            return
        self._window_state_needs_shown_restore = False
        QTimer.singleShot(0, self._finalize_shown_window_restore)

    def _finalize_shown_window_restore(self) -> None:
        """Restore after Qt has resolved native frame margins and child layout."""
        if not self.isVisible():
            return
        self._activate_window_layout()
        self._restore_window_state()
        self._capture_normal_window_geometry()
        self._shown_window_restore_finalized = True

    def moveEvent(self, event: QMoveEvent) -> None:
        super().moveEvent(event)
        self._schedule_pending_normal_anchor_restore_or_capture(spontaneous=event.spontaneous())
        if self._geometry_source_transition_rebase_pending:
            if self._applying_internal_panel_geometry or self._geometry_source_transition_settling:
                self._geometry_source_change_rect = self._last_normal_window_rect
            else:
                self._geometry_source_transition_rebase_pending = False

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        left_panel = getattr(self, "_left_panel_frame", None)
        right_panel = getattr(self, "_right_panel_frame", None)
        center_column = getattr(self, "_center_column", None)
        if left_panel is not None and right_panel is not None and center_column is not None:
            self._apply_projection_minimum_widths(
                # ``isVisible`` is false for explicitly enabled children until
                # their top-level parent is shown.  ``isHidden`` preserves the
                # projection intent during startup resizes as well as runtime.
                left_visible=not left_panel.isHidden(),
                right_visible=not right_panel.isHidden(),
            )
            self._activate_window_layout()
        self._schedule_pending_normal_anchor_restore_or_capture(spontaneous=event.spontaneous())
        if self._geometry_source_transition_rebase_pending:
            if event.spontaneous():
                self._geometry_source_transition_rebase_pending = False
            elif self._applying_internal_panel_geometry or self._geometry_source_transition_settling:
                self._geometry_source_change_rect = self._last_normal_window_rect
            else:
                self._geometry_source_transition_rebase_pending = False

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() != QEvent.Type.WindowStateChange:
            return
        if self._pending_normal_center_anchor is not None:
            if not self._pending_anchor_restore_scheduled:
                self._pending_anchor_restore_scheduled = True
                QTimer.singleShot(0, self._apply_pending_normal_center_anchor)
            return
        self._capture_normal_window_geometry()

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._active_threads:
            event.ignore()
            self._begin_execution_close()
            return

        self._save_gui_state()
        if self._system_tray_icon is not None:
            self._system_tray_icon.hide()
        self._finalize_shutdown()
        super().closeEvent(event)

    # ── Window geometry persistence ────────────────────────────────

    def _center_column_screen_x(self) -> int:
        center_column = self._center_column
        if center_column is None:
            return self.pos().x()
        try:
            return int(center_column.mapToGlobal(QPoint(0, 0)).x())
        except (AttributeError, RuntimeError):
            return self.pos().x()

    def _center_column_offset(self) -> int:
        """Return the center column's runtime X offset from normal window X."""
        return self._center_column_screen_x() - self.pos().x()

    def _restore_center_column_anchor(self, preserved_center_x: int) -> None:
        """Keep the center work area stable across visible side-panel changes."""
        delta_x = int(preserved_center_x) - self._center_column_screen_x()
        if delta_x == 0:
            return
        geometry = self._window_geometry
        candidate = WindowRect(
            x=self.pos().x() + delta_x,
            y=self.pos().y(),
            width=self.width(),
            height=self.height(),
        )
        recovered = recover_window_geometry(
            candidate,
            self._available_screen_rects(),
            min_width=geometry.min_width,
            min_height=geometry.min_height,
        )
        self._apply_recovered_geometry(recovered)

    def _schedule_pending_normal_anchor_restore_or_capture(
        self,
        *,
        spontaneous: bool,
    ) -> None:
        if self.isMaximized() or self.isFullScreen() or self.isMinimized():
            return
        if self._pending_normal_center_anchor is None:
            self._capture_normal_window_geometry(
                rebase_collapsed=not self._applying_internal_panel_geometry,
            )
            return
        if not spontaneous:
            # A direct move/resize issued after showNormal is an explicit newer
            # geometry decision. It wins over the queued panel-anchor restore.
            # Native normal-frame events are spontaneous on supported Qt/WM
            # paths; physical Windows variants remain an acceptance boundary.
            self._pending_normal_center_anchor = None
            self._pending_anchor_restore_scheduled = False
            self._capture_normal_window_geometry(rebase_collapsed=True)
            return
        # Any normal move/resize after WindowStateChange means Qt has delivered
        # a native normal-frame geometry. It may legitimately differ by a pixel
        # from the cached frame, so exact equality is not a readiness signal.
        self._pending_normal_frame_restored = True
        if self._pending_anchor_restore_scheduled:
            QTimer.singleShot(0, self._apply_pending_normal_center_anchor)
            return
        self._pending_anchor_restore_scheduled = True
        QTimer.singleShot(0, self._apply_pending_normal_center_anchor)

    def _apply_pending_normal_center_anchor(self) -> None:
        self._pending_anchor_restore_scheduled = False
        preserved_center_x = self._pending_normal_center_anchor
        if preserved_center_x is None:
            return
        if self.isMaximized() or self.isFullScreen() or self.isMinimized():
            return
        expected = self._last_normal_window_rect
        current = WindowRect(
            x=self.pos().x(),
            y=self.pos().y(),
            width=self.width(),
            height=self.height(),
        )
        if expected is not None and current != expected and not self._pending_normal_frame_restored:
            # WindowStateChange is delivered before Qt restores the native
            # normal frame on some platforms. Keep the target pending for a
            # short bounded quiet period. Native events normally re-arm the
            # zero-delay path; the fallback also covers the reverse ordering
            # where the final frame arrived just before WindowStateChange.
            self._pending_anchor_restore_scheduled = True
            QTimer.singleShot(16, self._force_pending_normal_center_anchor)
            return
        self._pending_normal_frame_restored = True
        self._pending_normal_center_anchor = None
        self._activate_window_layout()
        self._restore_center_column_anchor(preserved_center_x)
        self._capture_normal_window_geometry()

    def _force_pending_normal_center_anchor(self) -> None:
        """Finish one queued normal restore after a bounded WM quiet period."""
        self._pending_anchor_restore_scheduled = False
        if self._pending_normal_center_anchor is None:
            return
        if self.isMaximized() or self.isFullScreen() or self.isMinimized():
            return
        self._pending_normal_frame_restored = True
        self._apply_pending_normal_center_anchor()

    def _activate_window_layout(self) -> None:
        """Synchronize top-level and grid geometry before measuring the anchor."""
        root_layout = self.layout()
        if root_layout is not None:
            root_layout.activate()
        central = self.findChild(QWidget, "centralContainer")
        central_layout = central.layout() if central is not None else None
        if central_layout is not None:
            central_layout.activate()

    def _capture_normal_window_geometry(self, *, rebase_collapsed: bool = False) -> None:
        """Keep a paired frame rect/anchor offset for maximized close handling."""
        if self.isMaximized() or self.isFullScreen() or self.isMinimized():
            return
        center_column = getattr(self, "_center_column", None)
        if center_column is None:
            return
        position = self.pos()
        size = self.size()
        if size.width() <= 0 or size.height() <= 0:
            return
        self._last_normal_window_rect = WindowRect(
            x=position.x(),
            y=position.y(),
            width=size.width(),
            height=size.height(),
        )
        self._last_normal_center_offset = self._center_column_offset()
        self._last_normal_left_visible = self._left_panel_frame.isVisible()
        self._last_normal_right_visible = self._right_panel_frame.isVisible()
        if not self._last_normal_left_visible and not self._last_normal_right_visible:
            self._collapsed_normal_window_rect = self._last_normal_window_rect
            self._collapsed_normal_center_offset = self._last_normal_center_offset
        elif rebase_collapsed:
            self._rebase_collapsed_normal_geometry(
                rect=self._last_normal_window_rect,
                center_offset=self._last_normal_center_offset,
                left_visible=self._last_normal_left_visible,
                right_visible=self._last_normal_right_visible,
            )

    def _rebase_collapsed_normal_geometry(
        self,
        *,
        rect: WindowRect,
        center_offset: int,
        left_visible: bool,
        right_visible: bool,
    ) -> None:
        """Project an explicit normal-frame decision back to collapsed geometry."""
        central = self.findChild(QWidget, "centralContainer")
        layout = central.layout() if central is not None else None
        spacing = layout.horizontalSpacing() if isinstance(layout, QGridLayout) else 0
        width_plan = self._panel_width_plan(
            left_visible=left_visible,
            right_visible=right_visible,
            container_width=rect.width,
        )
        left_contribution = width_plan.left + spacing if left_visible else 0
        panel_contribution = self._context_panel_width_contribution(
            left_visible=left_visible,
            right_visible=right_visible,
            container_width=rect.width,
        )
        collapsed_center_offset = max(0, center_offset - left_contribution)
        center_x = rect.x + center_offset
        self._collapsed_normal_window_rect = WindowRect(
            x=center_x - collapsed_center_offset,
            y=rect.y,
            width=max(self.minimumWidth(), rect.width - panel_contribution),
            height=rect.height,
        )
        self._collapsed_normal_center_offset = collapsed_center_offset

    def _normal_window_snapshot(self) -> tuple[WindowRect, int]:
        """Return frame-consistent normal geometry and its matching center offset."""
        if self.isMaximized() or self.isFullScreen() or self.isMinimized():
            if self._last_normal_window_rect is not None:
                return self._last_normal_window_rect, self._last_normal_center_offset
        else:
            self._activate_window_layout()
            self._capture_normal_window_geometry()
            if self._last_normal_window_rect is not None:
                return self._last_normal_window_rect, self._last_normal_center_offset

        normal = self.normalGeometry()
        fallback = WindowRect(
            x=normal.x(),
            y=normal.y(),
            width=max(1, normal.width()),
            height=max(1, normal.height()),
        )
        return fallback, 0

    def _normal_window_rect(self) -> WindowRect:
        """Return only the normal frame rect for reset-source comparison."""
        return self._normal_window_snapshot()[0]

    @staticmethod
    def _screen_rect(screen: object) -> ScreenRect | None:
        try:
            available = screen.availableGeometry()  # type: ignore[attr-defined]
            width = int(available.width())
            height = int(available.height())
            if width <= 0 or height <= 0:
                return None
            return ScreenRect(
                x=int(available.x()),
                y=int(available.y()),
                width=width,
                height=height,
            )
        except (AttributeError, RuntimeError, TypeError, ValueError):
            return None

    def _screen_fit_rect(self, screen: object) -> ScreenRect | None:
        """Project a work area to frame-position/client-size capacity."""
        rect = self._screen_rect(screen)
        if rect is None:
            return None
        frame = self.frameGeometry()
        client = self.geometry()
        frame_width = max(0, frame.width() - client.width())
        frame_height = max(0, frame.height() - client.height())
        return ScreenRect(
            x=rect.x,
            y=rect.y,
            width=max(1, rect.width - frame_width),
            height=max(1, rect.height - frame_height),
        )

    def _set_normal_frame_geometry(self, rect: WindowRect) -> None:
        """Commit one client geometry whose native frame lands at ``rect.x/y``."""
        frame = self.frameGeometry()
        client = self.geometry()
        client_x = rect.x + max(0, client.x() - frame.x())
        client_y = rect.y + max(0, client.y() - frame.y())
        self.setGeometry(client_x, client_y, rect.width, rect.height)

    def _available_screen_rects(self) -> list[ScreenRect]:
        """Return current work areas with the primary screen first for stable ties."""
        primary = QApplication.primaryScreen()
        ordered = []
        if primary is not None:
            ordered.append(primary)
        for screen in QApplication.screens():
            if screen not in ordered:
                ordered.append(screen)
        return [rect for screen in ordered if (rect := self._screen_fit_rect(screen)) is not None]

    def _apply_recovered_geometry(self, recovered: RecoveredWindowGeometry) -> None:
        rebase_source_change = (
            self._geometry_source_change_rect is not None
            and self._geometry_source_change_rect == self._normal_window_snapshot()[0]
        )
        self.setMinimumSize(QSize(recovered.effective_min_width, recovered.effective_min_height))
        rect = recovered.rect
        self.resize(rect.width, rect.height)
        self.move(rect.x, rect.y)
        self._activate_window_layout()
        self._capture_normal_window_geometry()
        if rebase_source_change:
            self._geometry_source_change_rect = self._last_normal_window_rect

    def _save_gui_state(self) -> None:
        """Save canonical normal geometry to config, guarded by policy and env."""
        self._window_behavior = self._load_window_behavior()
        if not self._window_behavior.remember_gui_state:
            return
        if os.environ.get("DOCWEN_GUI_DISABLE_STATE_SAVE", "").strip().lower() in {
            "1",
            "true",
            "yes",
            "on",
        }:
            return

        controller = self._view_model.controller
        if controller is None:
            return
        cfg_port = getattr(controller, "config_port", None)
        if cfg_port is None:
            return
        if not self._persisted_window_geometry_policy.schema_supported:
            return
        if self.isVisible() and not self._shown_window_restore_finalized:
            return
        if not self.isVisible():
            self._capture_normal_window_geometry(rebase_collapsed=True)

        pending_anchor = self._pending_normal_center_anchor
        collapsed_rect = self._collapsed_normal_window_rect
        if collapsed_rect is not None:
            normal_rect = collapsed_rect
            center_offset = self._collapsed_normal_center_offset
        elif pending_anchor is not None and self._last_normal_window_rect is not None:
            normal_rect = self._last_normal_window_rect
            center_offset = pending_anchor - normal_rect.x
        else:
            normal_rect, center_offset = self._normal_window_snapshot()
        if self._geometry_source_transition_rebase_pending:
            self._geometry_source_transition_settling = False
            return
        if self._geometry_source_change_rect == normal_rect:
            return

        if collapsed_rect is None:
            if (
                self.isMaximized()
                or self.isFullScreen()
                or self.isMinimized()
                or self._pending_normal_center_anchor is not None
            ):
                left_visible = self._last_normal_left_visible
                right_visible = self._last_normal_right_visible
            else:
                left_visible = self._left_panel_frame.isVisible()
                right_visible = self._right_panel_frame.isVisible()
            panel_contribution = self._context_panel_width_contribution(
                left_visible=left_visible,
                right_visible=right_visible,
                container_width=normal_rect.width,
            )
        else:
            panel_contribution = 0
        if panel_contribution:
            normal_rect = WindowRect(
                x=normal_rect.x,
                y=normal_rect.y,
                width=max(self.minimumWidth(), normal_rect.width - panel_contribution),
                height=normal_rect.height,
            )

        with contextlib.suppress(Exception):
            values = build_canonical_geometry_values(
                normal_rect,
                center_offset=center_offset,
                unscale_value=self._unscale_window_value,
            )
            if cfg_port.set_many(values):
                self._persisted_window_geometry_policy = self._load_window_geometry(center_offset=0)
                self._geometry_source_change_rect = None
                self._geometry_source_transition_rebase_pending = False
                self._geometry_source_transition_settling = False

    def _context_panel_width_contribution(
        self,
        *,
        left_visible: bool,
        right_visible: bool,
        container_width: int | None = None,
    ) -> int:
        """Return side-panel space for semantic or viewport-constrained intent.

        Without ``container_width`` this is the configured contribution used
        to expand a collapsed normal window.  A concrete width requests the
        proportional runtime allocation used when rebasing or persisting an
        already constrained window.
        """
        central = self.findChild(QWidget, "centralContainer")
        layout = central.layout() if central is not None else None
        spacing = layout.horizontalSpacing() if isinstance(layout, QGridLayout) else 0
        if container_width is None:
            return (self._LEFT_PANEL_MIN_WIDTH + spacing if left_visible else 0) + (
                self._RIGHT_PANEL_MIN_WIDTH + spacing if right_visible else 0
            )
        width_plan = self._panel_width_plan(
            left_visible=left_visible,
            right_visible=right_visible,
            container_width=container_width,
        )
        return (width_plan.left + spacing if left_visible else 0) + (width_plan.right + spacing if right_visible else 0)

    def _restore_window_state(self) -> None:
        """Restore schema-v2 geometry and recover it to an available screen."""
        if self._window_behavior.auto_center or not self._window_behavior.remember_gui_state:
            self._center_on_screen()
            return

        current_center_offset = self._center_column_offset()
        self._window_geometry = self._load_window_geometry(center_offset=current_center_offset)
        policy = self._window_geometry
        left_visible = self._left_panel_frame.isVisible()
        right_visible = self._right_panel_frame.isVisible()
        contribution = self._context_panel_width_contribution(
            left_visible=left_visible,
            right_visible=right_visible,
        )
        central = self.findChild(QWidget, "centralContainer")
        layout = central.layout() if central is not None else None
        spacing = layout.horizontalSpacing() if isinstance(layout, QGridLayout) else 0
        left_contribution = self._LEFT_PANEL_MIN_WIDTH + spacing if left_visible else 0
        collapsed_center_offset = max(0, current_center_offset - left_contribution)
        configured_center_x = policy.rect.x + current_center_offset
        collapsed_candidate = WindowRect(
            x=configured_center_x - collapsed_center_offset,
            y=policy.rect.y,
            width=policy.rect.width,
            height=policy.rect.height,
        )
        recovered_collapsed = recover_window_geometry(
            collapsed_candidate,
            self._available_screen_rects(),
            min_width=policy.min_width,
            min_height=policy.min_height,
        )
        self._collapsed_normal_window_rect = recovered_collapsed.rect
        self._collapsed_normal_center_offset = collapsed_center_offset
        recovered_center_x = recovered_collapsed.rect.x + collapsed_center_offset
        candidate = WindowRect(
            x=recovered_center_x - current_center_offset,
            y=recovered_collapsed.rect.y,
            width=recovered_collapsed.rect.width + contribution,
            height=recovered_collapsed.rect.height,
        )
        recovered = recover_window_geometry(
            candidate,
            self._available_screen_rects(),
            min_width=policy.min_width,
            min_height=policy.min_height,
        )
        self._apply_recovered_geometry(recovered)

    def _center_on_screen(self) -> None:
        """Center the current size on its work area and cap oversized minima."""
        qt_screen = self.screen() or QApplication.primaryScreen()
        screen = self._screen_fit_rect(qt_screen) if qt_screen is not None else None
        if screen is None:
            return
        geometry = self._window_geometry
        recovered = center_window_geometry(
            width=self.width(),
            height=self.height(),
            screen=screen,
            min_width=geometry.min_width,
            min_height=geometry.min_height,
        )
        self._apply_recovered_geometry(recovered)

    # ── System tray ─────────────────────────────────────────────────

    def _setup_system_tray(self) -> None:
        """Create system tray icon with context menu."""
        controller = self._view_model.controller
        cfg_port = getattr(controller, "config_port", None) if controller is not None else None
        try:
            tray_enabled = bool(cfg_port.get("gui.notifications.system_tray", False)) if cfg_port else False
        except Exception:
            tray_enabled = False

        if not tray_enabled or not QSystemTrayIcon.isSystemTrayAvailable():
            return

        app = QApplication.instance()
        icon = app.windowIcon() if isinstance(app, QApplication) else None

        self._system_tray_icon = QSystemTrayIcon(icon, self)  # pyright: ignore[reportArgumentType]

        menu = QMenu(self)
        show_action = menu.addAction(_t("main_window.system_tray_show", "Show Window"))
        show_action.triggered.connect(self.bring_to_front)
        hide_action = menu.addAction(_t("main_window.system_tray_hide", "Hide Window"))
        hide_action.triggered.connect(self.hide)
        menu.addSeparator()
        quit_action = menu.addAction(_t("main_window.system_tray_quit", "Quit"))
        quit_action.triggered.connect(self.close)

        tray = self._system_tray_icon
        if tray is not None:
            tray.setContextMenu(menu)
            tray.activated.connect(self._on_tray_activated)
            tray.show()

    def _on_tray_activated(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.bring_to_front()

    # ── Task completion notifications ───────────────────────────────

    def _maybe_notify_task_completion(self, context: dict[str, Any]) -> None:
        """Show task completion notification if configured and warranted."""
        controller = self._view_model.controller
        if controller is None:
            return
        cfg_port = getattr(controller, "config_port", None)
        if cfg_port is None:
            return

        # Check if task completion notifications are enabled
        try:
            task_completion = cfg_port.get("gui.notifications.task_completion", True)
            if not task_completion:
                return
        except Exception:
            pass

        # Check minimum elapsed time
        if self._start_time is not None:
            elapsed = time.monotonic() - self._start_time
            try:
                min_elapsed_raw = cfg_port.get("gui.notifications.min_elapsed_seconds", 8.0)
                min_elapsed = float(min_elapsed_raw) if min_elapsed_raw is not None else 8.0
            except (TypeError, ValueError):
                min_elapsed = 8.0
            if elapsed < min_elapsed:
                return

        # Taskbar flash
        app = QApplication.instance()
        if app is not None:
            try:
                alert_fn = getattr(app, "alert", None)
                if callable(alert_fn):
                    alert_fn(self, 0)
            except Exception:
                pass

        # System tray notification
        try:
            tray_enabled = cfg_port.get("gui.notifications.system_tray", False)
        except Exception:
            tray_enabled = False

        if tray_enabled and QSystemTrayIcon.isSystemTrayAvailable():
            tray = self._system_tray_icon
            if tray is not None:
                try:
                    timeout_raw = cfg_port.get("gui.notifications.system_tray_timeout_ms", 6000)
                    timeout_ms = max(
                        1000,
                        min(30000, int(timeout_raw)),
                    )
                except (TypeError, ValueError):
                    timeout_ms = 6000

                display_name = context.get("display_name", "")
                summary = self._info_area_vm.task_summary
                state_key = {
                    "success": "components.info_area.task_completion_notification_success",
                    "partial": "components.info_area.task_completion_notification_partial",
                    "failed": "components.info_area.task_completion_notification_failed",
                    "cancelled": "components.info_area.task_completion_notification_cancelled",
                }.get(summary.state, "components.info_area.task_completion_notification_success")
                completed = summary.completed_count or (1 if summary.state == "" else 0)
                total = summary.total_count or (1 if summary.state == "" else 0)
                failed = summary.failed_count

                try:
                    from .i18n import t as i18n_t

                    title = i18n_t(
                        "components.info_area.task_completion_notification_title",
                        "DocWen",
                    )
                    body = i18n_t(
                        state_key,
                        f"Task completed: {display_name}",
                        completed=completed,
                        total=total,
                        failed=failed,
                    )
                except Exception:
                    title = "DocWen"
                    body = f"Task completed: {display_name}"

                tray.showMessage(
                    title,
                    body,
                    QSystemTrayIcon.MessageIcon.Information,
                    timeout_ms,
                )


__all__ = ["DEFAULT_CENTER_WIDTH", "DEFAULT_HEIGHT", "MIN_HEIGHT", "MIN_WIDTH", "MainWindow"]
