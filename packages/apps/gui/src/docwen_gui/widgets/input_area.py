"""InputArea widget — file drop area with mode switch, add/clear buttons.

Widgets never call runtime/plugins directly — they go through
``InputAreaViewModel`` which delegates to ``MainWindowViewModel``.

Key behaviors:
- Drag-and-drop from file manager (URL) or text editor (text/plain)
- Single/batch mode switch via ``FluentSegmentedWidget``
- Add button: file dialog (single) or popup menu (batch)
- Clear button: reset selection
- Compact layout at width <= 340px
- Default height 230px with a compact title icon and semantic prompt label
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import TYPE_CHECKING
from typing import cast as _cast

import shiboken6
from PySide6.QtCore import QEvent, QSize, Qt, QTimer, Signal
from PySide6.QtGui import (
    QAction,
    QCloseEvent,
    QDragEnterEvent,
    QDragLeaveEvent,
    QDragMoveEvent,
    QDropEvent,
    QIcon,
    QPalette,
    QPixmap,
    QResizeEvent,
    QShowEvent,
)
from PySide6.QtWidgets import (
    QApplication,
    QBoxLayout,
    QFileDialog,
    QFrame,
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QLabel,
    QMenu,
    QSizePolicy,
    QStyle,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    CaptionLabel,
    PrimaryPushButton,
    PushButton,
    SegmentedWidget,
    SimpleCardWidget,
    StrongBodyLabel,
)

from docwen_gui.i18n import t
from docwen_gui.styles.design_tokens import Sizing

from .batch_list import _MiddleElidedLabel

if TYPE_CHECKING:
    from ..view_models.input_area_vm import InputAreaViewModel

# ── Design constants ────────────────────────────────────────────────────
_DEFAULT_HEIGHT = 230
_COMPACT_WIDTH_THRESHOLD = 340
_ORNAMENT_SIZE = QSize(72, 72)
_SPACING_XS = 4
_SPACING_SM = 8
_SPACING_MD = 12
_PYRAMID_INDENTS = (72, 58, 44, 30, 18, 8)
_ACTION_BUTTON_MIN_WIDTH = 72

_SUPPORTED_TYPE_ROWS: tuple[tuple[str, str, str], ...] = (
    ("file_category.text_short", "Text", "MD, TXT"),
    ("file_category.layout_short", "Layout", "PDF, XPS, OFD"),
    ("file_category.spreadsheet_short", "Sheet", "XLSX, XLS, ET, CSV, TSV, ODS"),
    ("file_category.document_short", "Doc", "DOCX, DOC, WPS, RTF, ODT"),
    ("file_category.image_short", "Image", "JPG, PNG, BMP, GIF, HEIC, HEIF, WebP"),
    ("file_category.other_short", "Other", "HTML, MHTML, ENEX, PPTX, PPT, EPUB"),
)

# MIME types for drag-and-drop
_MIME_URL = "text/uri-list"
_MIME_TEXT = "text/plain"


def _load_hero_icon() -> QPixmap | None:
    """Load the composite drop artwork, falling back to a platform file icon."""
    from ..resources import load_svg_asset_icon

    artwork = load_svg_asset_icon("file_drop_empty_state.svg")
    if artwork is not None and not artwork.isNull():
        screen = QApplication.primaryScreen()
        dpr = screen.devicePixelRatio() if screen is not None else 1.0
        return artwork.pixmap(_ORNAMENT_SIZE, dpr)
    style = QApplication.style()
    if style is None:
        return None
    icon = style.standardIcon(QStyle.StandardPixmap.SP_FileIcon)
    if icon.isNull():
        return None
    return icon.pixmap(_ORNAMENT_SIZE)


def _i18n(key: str, default: str = "", **kwargs) -> str:
    """Translate one widget label through the GUI locale service."""
    return t(key, default=default, **kwargs)


# ── i18n key constants ───────────────────────────────────────────────────
_I_BATCH_MODE = "components.file_drop.batch_mode"
_I_SINGLE_MODE = "components.file_drop.single_mode"
_I_ADD_BUTTON = "components.file_drop.add_button"
_I_CLEAR_BUTTON = "components.file_drop.clear_button"
_I_ADD_FILE = "components.file_drop.add_file_action"
_I_ADD_FOLDER = "components.file_drop.add_folder_action"
_I_RECENT_FILES = "components.file_drop.recent_files_action"
_I_CLEAR_RECENT = "components.file_drop.clear_recent_files_action"
_I_SELECT_FILE = "components.file_drop.select_file_dialog"
_I_SELECT_FOLDER = "components.file_drop.select_folder_dialog"
_I_EMPTY_SINGLE = "components.file_drop.empty_hint_single"
_I_EMPTY_BATCH = "components.file_drop.empty_hint_batch"
_I_TRANSIENT_TITLE = "info_area.transient_title"


class InputArea(QFrame):
    """File drop area widget.

    Provides drag-and-drop file input, single/batch mode switching,
    and add/clear buttons. All user actions delegate to ``InputAreaViewModel``.

    Signals:
        height_changed: emitted when the widget's effective height changes.
    """

    height_changed = Signal(int)

    def __init__(
        self,
        view_model: InputAreaViewModel,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._vm = view_model
        self._drag_active = False
        self._selected_file: str = _cast(str, None)
        self._top_controls_compact = False
        self._recent_files: list[str] = []
        self._deferred_prompt_layout = False
        self._deferred_supported_type_layout = False
        self._deferred_layout_sync_enabled = True
        self._deferred_layout_sync_timer = QTimer(self)
        self._deferred_layout_sync_timer.setSingleShot(True)
        self._deferred_layout_sync_timer.setInterval(0)
        self._deferred_layout_sync_timer.timeout.connect(self._flush_deferred_layout_sync)

        self.setObjectName("inputArea")
        self.setAcceptDrops(True)
        self.setMinimumHeight(_DEFAULT_HEIGHT)

        self._build_ui()
        self._wire_view_model()
        self._sync_visual_state()

    # ── UI Construction ────────────────────────────────────────────

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, _SPACING_XS, 0, _SPACING_XS)
        layout.setSpacing(_SPACING_SM)

        # Drop group frame
        self._drop_group = QFrame(self)
        self._drop_group.setObjectName("fileDropGroup")
        drop_layout = QVBoxLayout(self._drop_group)
        drop_layout.setContentsMargins(_SPACING_MD, _SPACING_MD, _SPACING_MD, _SPACING_MD)
        drop_layout.setSpacing(_SPACING_SM)
        layout.addWidget(self._drop_group)

        # Top controls layout (mode switch + buttons)
        self._top_layout = QBoxLayout(QBoxLayout.Direction.LeftToRight)
        self._top_layout.setContentsMargins(0, 0, 0, 0)
        self._top_layout.setSpacing(_SPACING_XS)

        # Mode switch
        self._mode_switch = SegmentedWidget(self._drop_group)
        self._mode_switch.setObjectName("fileDropModeSwitch")
        self._mode_switch.setAccessibleName(_i18n(_I_BATCH_MODE))
        self._mode_switch.addItem(
            "batch",
            _i18n(_I_BATCH_MODE, "Batch"),
            onClick=lambda: self._vm.set_mode("batch"),
        )
        self._mode_switch.addItem(
            "single",
            _i18n(_I_SINGLE_MODE, "Single"),
            onClick=lambda: self._vm.set_mode("single"),
        )
        self._mode_switch.setCurrentItem(self._vm.mode)
        self._top_layout.addWidget(self._mode_switch)

        # Action buttons frame
        self._action_frame = QFrame(self._drop_group)
        self._action_frame.setObjectName("fileDropActionButtonsFrame")
        action_layout = QHBoxLayout(self._action_frame)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(_SPACING_XS)

        # Add button
        self._add_button = PrimaryPushButton(_i18n(_I_ADD_BUTTON, "Add"), self._drop_group)
        self._add_button.setObjectName("fileDropPrimaryButton")
        self._add_button.setMinimumSize(_ACTION_BUTTON_MIN_WIDTH, Sizing.CONTROL_HEIGHT)
        self._add_button.clicked.connect(self._on_add_clicked)
        action_layout.addWidget(self._add_button)

        # Clear button (danger theme for destructive action)
        self._clear_button = PushButton(_i18n(_I_CLEAR_BUTTON, "Clear"), self._drop_group)
        self._clear_button.setObjectName("fileDropClearButton")
        self._clear_button.setMinimumSize(_ACTION_BUTTON_MIN_WIDTH, Sizing.CONTROL_HEIGHT)
        self._clear_button.setProperty("danger", True)
        self._clear_button.setToolTip(_i18n(_I_CLEAR_BUTTON, "Clear"))
        self._clear_button.setAccessibleName(_i18n(_I_CLEAR_BUTTON, "Clear"))
        self._clear_button.setAccessibleDescription(_i18n(_I_CLEAR_BUTTON, "Clear"))
        self._clear_button.clicked.connect(self._on_clear_clicked)
        action_layout.addWidget(self._clear_button)

        self._top_layout.addWidget(self._action_frame)

        # Set tab order: mode switch items -> add -> clear
        self.setTabOrder(self._mode_switch, self._add_button)
        self.setTabOrder(self._add_button, self._clear_button)

        # Empty state card
        self._empty_state_frame = SimpleCardWidget(self._drop_group)
        self._empty_state_frame.setObjectName("fileDropEmptyStateFrame")
        empty_layout = QVBoxLayout(self._empty_state_frame)
        empty_layout.setContentsMargins(_SPACING_MD, _SPACING_SM, _SPACING_MD, _SPACING_MD)
        empty_layout.setSpacing(_SPACING_SM)
        empty_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        # Eyebrow label
        self._eyebrow_label = CaptionLabel(self._empty_state_frame)
        self._eyebrow_label.setObjectName("fileDropEyebrowLabel")
        self._eyebrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._eyebrow_label.setVisible(False)

        # Empty state content (hero icon + prompt + supported formats)
        self._empty_content = QWidget(self._empty_state_frame)
        self._empty_content.setObjectName("fileDropEmptyStateContent")
        self._empty_content.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        content_layout = QHBoxLayout(self._empty_content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(_SPACING_SM)
        content_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        self._empty_center_panel = QWidget(self._empty_content)
        self._empty_center_panel.setObjectName("fileDropEmptyStateCenterPanel")
        center_layout = QVBoxLayout(self._empty_center_panel)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(_SPACING_SM)
        center_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)

        self._empty_title_row = QWidget(self._empty_center_panel)
        self._empty_title_row.setObjectName("fileDropEmptyStateTitleRow")
        title_layout = QHBoxLayout(self._empty_title_row)
        title_layout.setContentsMargins(0, 0, 0, 0)
        title_layout.setSpacing(_SPACING_SM)
        title_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)

        # Hero icon
        self._hero_icon_label = QLabel(self._empty_content)
        self._hero_icon_label.setObjectName("fileDropHeroIconLabel")
        self._hero_icon_label.setProperty("heroVariant", "symbol")
        self._hero_icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._hero_icon_label.setFixedSize(_ORNAMENT_SIZE)
        hero_icon = _load_hero_icon()
        if hero_icon is not None and not hero_icon.isNull():
            self._hero_icon_label.setPixmap(hero_icon)
        self._hero_opacity = QGraphicsOpacityEffect(self._hero_icon_label)
        self._sync_hero_icon_opacity()
        self._hero_icon_label.setGraphicsEffect(self._hero_opacity)

        # Prompt label
        self._prompt_label = StrongBodyLabel(self._empty_title_row)
        self._prompt_label.setObjectName("fileDropPromptLabel")
        self._prompt_label.setWordWrap(True)
        self._prompt_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._prompt_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self._update_prompt_text()

        title_layout.addWidget(self._hero_icon_label, alignment=Qt.AlignmentFlag.AlignCenter)
        title_layout.addWidget(self._prompt_label, alignment=Qt.AlignmentFlag.AlignCenter)
        center_layout.addWidget(self._empty_title_row, alignment=Qt.AlignmentFlag.AlignCenter)

        self._types_container = QWidget(self._empty_center_panel)
        self._types_container.setObjectName("fileDropTypesContainer")
        self._types_layout = QVBoxLayout(self._types_container)
        self._types_layout.setContentsMargins(0, 0, 0, 0)
        self._types_layout.setSpacing(2)
        self._types_container.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self._type_prompt_rows: list[tuple[QWidget, QHBoxLayout, QLabel, QLabel]] = []
        for label_key, fallback_label, formats in _SUPPORTED_TYPE_ROWS:
            row_widget = QWidget(self._types_container)
            row_widget.setObjectName("fileDropTypesRow")
            row_widget.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(_SPACING_SM)

            type_label = QLabel(f"{_i18n(label_key, fallback_label)}:", row_widget)
            type_label.setObjectName("fileDropTypesTypeLabel")
            type_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
            type_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            value_label = QLabel(formats, row_widget)
            value_label.setObjectName("fileDropTypesValueLabel")
            value_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
            value_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

            row_layout.addWidget(type_label, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            row_layout.addStretch(1)
            row_layout.addWidget(value_label, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self._types_layout.addWidget(row_widget)
            self._type_prompt_rows.append((row_widget, row_layout, type_label, value_label))
        self._sync_supported_type_layout()
        center_layout.addWidget(self._types_container, alignment=Qt.AlignmentFlag.AlignCenter)
        content_layout.addWidget(
            self._empty_center_panel,
            stretch=1,
            alignment=Qt.AlignmentFlag.AlignVCenter,
        )

        # Feedback frame (selection message)
        self._feedback_frame = QFrame(self._empty_state_frame)
        self._feedback_frame.setObjectName("fileDropFeedbackFrame")
        feedback_layout = QVBoxLayout(self._feedback_frame)
        feedback_layout.setContentsMargins(_SPACING_MD, _SPACING_SM, _SPACING_MD, _SPACING_SM)
        feedback_layout.setSpacing(_SPACING_XS)

        self._feedback_title_label = CaptionLabel(self._feedback_frame)
        self._feedback_title_label.setObjectName("fileDropFeedbackTitleLabel")
        self._feedback_title_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self._feedback_title_label.setText(_i18n(_I_TRANSIENT_TITLE, "Status"))

        self._selection_label = QLabel(self._feedback_frame)
        self._selection_label.setObjectName("fileDropSelectionLabel")
        self._selection_label.setWordWrap(True)
        self._selection_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self._selection_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self._selection_label.setMinimumWidth(0)
        self._selection_label.setText("")

        self._selection_detail_label = _MiddleElidedLabel("", self._feedback_frame)
        self._selection_detail_label.setObjectName("fileDropSelectionDetailLabel")
        self._selection_detail_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self._selection_detail_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self._selection_detail_label.setVisible(False)

        feedback_layout.addWidget(self._feedback_title_label)
        feedback_layout.addWidget(self._selection_label)
        feedback_layout.addWidget(self._selection_detail_label)

        # Assemble drop layout
        drop_layout.addLayout(self._top_layout)
        drop_layout.addStretch(1)
        empty_layout.addWidget(self._eyebrow_label, alignment=Qt.AlignmentFlag.AlignCenter)
        empty_layout.addWidget(self._empty_content, alignment=Qt.AlignmentFlag.AlignVCenter)
        empty_layout.addWidget(self._feedback_frame)
        drop_layout.addWidget(self._empty_state_frame)
        drop_layout.addStretch(1)

        self._sync_top_control_layout()

    # ── ViewModel wiring ───────────────────────────────────────────

    def _wire_view_model(self) -> None:
        vm = self._vm
        vm.mode_changed.connect(self._on_mode_changed)
        vm.selection_message_changed.connect(self._on_selection_message_changed)

    def _on_mode_changed(self, mode: str) -> None:
        # Update switch to reflect current mode
        self._mode_switch.setCurrentItem(mode)
        self._update_prompt_text()
        self._sync_visual_state()

    def _on_selection_message_changed(self, message: str, tone: str) -> None:
        self._selection_label.setText(message)
        detail = self._vm.selection_detail
        self._selection_detail_label.set_full_text(detail)
        self._selection_detail_label.setVisible(bool(detail))
        self._feedback_frame.setProperty("feedbackTone", tone)
        self._selection_label.setProperty("feedbackTone", tone)
        self._sync_visual_state()

    # ── Drag-and-drop events ───────────────────────────────────────

    def _drag_paths_from_mime_data(self, mime_data) -> list[str]:
        """Extract local paths from drag MIME data for preview and drop."""
        paths: list[str] = []
        if mime_data.hasFormat(_MIME_URL):
            urls = mime_data.urls()
            paths = self._vm.extract_urls_from_mime_data(urls)
        if not paths and mime_data.hasFormat(_MIME_TEXT):
            try:
                paths = self._vm.extract_paths_from_text_payload(mime_data.text())
            except Exception:
                paths = []
        return paths

    def _restore_drag_preview_message(self) -> None:
        self._selection_label.setText(self._vm.selection_message or "")
        self._selection_label.setToolTip("")
        detail = self._vm.selection_detail
        self._selection_detail_label.set_full_text(detail)
        self._selection_detail_label.setVisible(bool(detail))
        self._feedback_frame.setProperty("feedbackTone", self._vm.selection_tone)
        self._selection_label.setProperty("feedbackTone", self._vm.selection_tone)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasFormat(_MIME_URL) or event.mimeData().hasFormat(_MIME_TEXT):
            event.acceptProposedAction()
            self._drag_active = True
            preview = self._vm.build_drag_preview(self._drag_paths_from_mime_data(event.mimeData()))
            self._selection_label.setText(preview.message)
            self._selection_label.setToolTip(preview.tooltip)
            self._selection_detail_label.set_full_text("")
            self._selection_detail_label.setVisible(False)
            self._feedback_frame.setProperty("feedbackTone", preview.tone)
            self._selection_label.setProperty("feedbackTone", preview.tone)
            self._sync_visual_state()
        else:
            event.ignore()

    def dragMoveEvent(self, event: QDragMoveEvent) -> None:
        """Reject drag if no supported content; accept otherwise."""
        mime_data = event.mimeData()
        has_urls = mime_data.hasFormat(_MIME_URL)
        has_text = mime_data.hasFormat(_MIME_TEXT)
        if has_urls or has_text:
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event: QDragLeaveEvent) -> None:
        self._drag_active = False
        self._restore_drag_preview_message()
        self._sync_visual_state()
        event.accept()

    def dropEvent(self, event: QDropEvent) -> None:
        self._drag_active = False
        paths = self._drag_paths_from_mime_data(event.mimeData())
        if paths:
            self._vm.add_files(paths)
        event.acceptProposedAction()
        self._sync_visual_state()

    # ── Button handlers ────────────────────────────────────────────

    def _on_add_clicked(self) -> None:
        if self._vm.mode == "batch":
            self._show_add_menu()
        else:
            self._open_file_dialog()

    def _open_file_dialog(self) -> None:
        initial_dir = self._resolve_initial_dir()
        filter_text = self._vm.build_file_dialog_filter()
        if self._vm.mode == "batch":
            paths, _ = QFileDialog.getOpenFileNames(
                self,
                _i18n(_I_SELECT_FILE, "Select Files"),
                initial_dir,
                filter_text,
            )
            if paths:
                self._vm.add_files(list(paths))
        else:
            path, _ = QFileDialog.getOpenFileName(
                self,
                _i18n(_I_SELECT_FILE, "Select File"),
                initial_dir,
                filter_text,
            )
            if path:
                self._vm.add_files([path])

    def _open_folder_dialog(self) -> None:
        initial_dir = self._resolve_initial_dir()
        path = QFileDialog.getExistingDirectory(
            self,
            _i18n(_I_SELECT_FOLDER, "Select Folder"),
            initial_dir,
        )
        if path:
            self._vm.add_files([path])

    def _show_add_menu(self) -> None:
        menu = QMenu(self)
        popup_pos = self._add_button.mapToGlobal(self._add_button.rect().bottomLeft())

        add_file_action = QAction(_i18n(_I_ADD_FILE, "Add File"), self)
        add_file_action.triggered.connect(self._open_file_dialog)
        menu.addAction(add_file_action)

        add_folder_action = QAction(_i18n(_I_ADD_FOLDER, "Add Folder"), self)
        add_folder_action.triggered.connect(self._open_folder_dialog)
        menu.addAction(add_folder_action)

        if self._recent_files:
            recent_menu = menu.addMenu(_i18n(_I_RECENT_FILES, "Recent Files"))
            for file_path in self._recent_files:
                action = QAction(Path(file_path).name or file_path, self)
                action.setToolTip(file_path)
                action.triggered.connect(lambda checked=False, p=file_path: self._vm.add_files([p]))
                recent_menu.addAction(action)
            recent_menu.addSeparator()
            clear_recent_action = QAction(_i18n(_I_CLEAR_RECENT, "Clear"), self)
            clear_recent_action.triggered.connect(self._clear_recent_files)
            recent_menu.addAction(clear_recent_action)

        menu.exec(popup_pos)

    def _on_clear_clicked(self) -> None:
        self._selected_file = None  # pyright: ignore[reportAttributeAccessIssue]
        self._selection_label.setText("")
        self._selection_label.setToolTip("")
        self._selection_detail_label.set_full_text("")
        self._selection_detail_label.setVisible(False)
        self._feedback_frame.setVisible(False)
        self._vm.clear_files()
        self._sync_visual_state()

    def _clear_recent_files(self) -> None:
        self._recent_files.clear()

    # ── Visual state management ────────────────────────────────────

    def _sync_visual_state(self) -> None:
        """Sync visual state. Selection state is derived from ViewModel message."""
        has_selection = bool(self._vm.selection_message.strip())

        self._empty_content.setVisible(not has_selection)
        self._feedback_frame.setVisible(has_selection)
        if self._eyebrow_label and self._eyebrow_label.isVisible():
            self._eyebrow_label.setVisible(has_selection)

        # 区域级拖拽反馈：drag-active 优先于已选中的常亮边框（样式见 styles/panel.py）
        if self._drag_active:
            drop_state: str | None = "drag-active"
        elif has_selection:
            drop_state = "selected"
        else:
            drop_state = None
        if self._drop_group.property("dropState") != drop_state:
            self._drop_group.setProperty("dropState", drop_state)
            style = self._drop_group.style()
            style.unpolish(self._drop_group)
            style.polish(self._drop_group)

        if has_selection:
            self.setMinimumHeight(self._compact_min_for_feedback())
        else:
            self.setMinimumHeight(_DEFAULT_HEIGHT)

    def _compact_min_for_feedback(self) -> int:
        return max(_SPACING_MD * 2, 80)

    def _update_prompt_text(self) -> None:
        if self._vm.mode == "batch":
            text = _i18n(_I_EMPTY_BATCH, "Drag files or folders here, or click Add")
        else:
            text = _i18n(_I_EMPTY_SINGLE, "Drag a single file here, or click Add")
        self._prompt_label.setText(text)
        self._schedule_deferred_layout_sync(prompt=True)

    def _resolve_initial_dir(self) -> str:
        if self._selected_file:
            candidate = Path(self._selected_file)
            if candidate.exists():
                if candidate.is_dir():
                    return str(candidate)
                if candidate.is_file():
                    return str(candidate.parent)
            parent = candidate.parent
            if parent.exists() and parent.is_dir():
                return str(parent)
        home = Path.home()
        if home.exists() and home.is_dir():
            return str(home)
        return os.getcwd()

    # ── Compact layout ─────────────────────────────────────────────

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        self._sync_top_control_layout()
        self._sync_prompt_layout()
        self._sync_supported_type_layout()
        self._schedule_deferred_layout_sync(prompt=True, supported_types=True)

    def showEvent(self, event: QShowEvent) -> None:
        self._deferred_layout_sync_enabled = True
        super().showEvent(event)
        self._schedule_deferred_layout_sync(prompt=True, supported_types=True)

    def closeEvent(self, event: QCloseEvent) -> None:
        self._cancel_deferred_layout_sync()
        super().closeEvent(event)

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() in {QEvent.Type.FontChange, QEvent.Type.StyleChange}:
            self._schedule_deferred_layout_sync(prompt=True, supported_types=True)
            self._sync_hero_icon_opacity()

    def _schedule_deferred_layout_sync(
        self,
        *,
        prompt: bool = False,
        supported_types: bool = False,
    ) -> None:
        """Coalesce layout work on a timer owned by this widget.

        Parenting the timer to ``InputArea`` makes queued layout work disappear
        with the widget instead of leaving a Python callback that can outlive
        its C++ children during application/test teardown.
        """
        if not self._deferred_layout_sync_enabled or not shiboken6.isValid(self._deferred_layout_sync_timer):
            return
        self._deferred_prompt_layout = self._deferred_prompt_layout or prompt
        self._deferred_supported_type_layout = self._deferred_supported_type_layout or supported_types
        if not self._deferred_layout_sync_timer.isActive():
            self._deferred_layout_sync_timer.start()

    def _cancel_deferred_layout_sync(self) -> None:
        self._deferred_layout_sync_enabled = False
        self._deferred_prompt_layout = False
        self._deferred_supported_type_layout = False
        if shiboken6.isValid(self._deferred_layout_sync_timer):
            self._deferred_layout_sync_timer.stop()

    def _flush_deferred_layout_sync(self) -> None:
        if not self._deferred_layout_sync_enabled or not shiboken6.isValid(self):
            self._deferred_prompt_layout = False
            self._deferred_supported_type_layout = False
            return
        sync_prompt = self._deferred_prompt_layout
        sync_supported_types = self._deferred_supported_type_layout
        self._deferred_prompt_layout = False
        self._deferred_supported_type_layout = False
        if sync_prompt:
            self._sync_prompt_layout()
        if sync_supported_types:
            self._sync_supported_type_layout()

    def _sync_hero_icon_opacity(self) -> None:
        effect = getattr(self, "_hero_opacity", None)
        if effect is None:
            return
        window_lightness = self.palette().color(QPalette.ColorRole.Window).lightness()
        effect.setOpacity(0.54 if window_lightness < 128 else 0.82)

    def _sync_prompt_layout(self) -> None:
        """Keep the hero prompt on one line whenever the visible card can hold it."""
        if not self._prompt_layout_objects_are_valid():
            return
        artwork_width = _ORNAMENT_SIZE.width() + _SPACING_SM
        fallback_width = self._drop_group.width() - (_SPACING_MD * 4) - artwork_width
        panel_width = self._empty_center_panel.width()
        visible_width = panel_width - artwork_width
        available_width = max(visible_width if panel_width > 0 else fallback_width, 0)
        if available_width <= 0:
            return
        text_width = self._prompt_label.fontMetrics().horizontalAdvance(self._prompt_label.text()) + 2
        should_wrap = text_width > available_width
        self._prompt_label.setWordWrap(should_wrap)
        self._prompt_label.setMinimumWidth(0 if should_wrap else text_width)
        self._prompt_label.setMaximumWidth(available_width if should_wrap else 16777215)
        self._prompt_label.updateGeometry()

    def _prompt_layout_objects_are_valid(self) -> bool:
        if not shiboken6.isValid(self):
            return False
        return all(
            shiboken6.isValid(widget)
            for widget in (
                getattr(self, "_drop_group", None),
                getattr(self, "_empty_center_panel", None),
                getattr(self, "_prompt_label", None),
            )
        )

    def _sync_top_control_layout(self) -> None:
        content_width = max(self._drop_group.width() - (_SPACING_MD * 2), 0)
        compact = 0 < content_width < _COMPACT_WIDTH_THRESHOLD

        if compact == self._top_controls_compact:
            return

        self._top_controls_compact = compact

        if compact:
            self._top_layout.setDirection(QBoxLayout.Direction.TopToBottom)
            self._top_layout.setSpacing(_SPACING_SM)
            self._clear_button.setText("")
            self._clear_button.setToolTip(_i18n(_I_CLEAR_BUTTON, "Clear"))
            self._clear_button.setIconSize(QSize(16, 16))
            style = self.style() or QApplication.style()
            if style is not None:
                icon = style.standardIcon(QStyle.StandardPixmap.SP_LineEditClearButton)
                if not icon.isNull():
                    self._clear_button.setIcon(icon)
        else:
            self._top_layout.setDirection(QBoxLayout.Direction.LeftToRight)
            self._top_layout.setSpacing(_SPACING_XS)
            self._clear_button.setText(_i18n(_I_CLEAR_BUTTON, "Clear"))
            self._clear_button.setIcon(QIcon())

        # Keep the two text actions visually balanced at normal widths.  The
        # icon-only clear action remains intentionally smaller in compact mode.
        self._add_button.setMinimumWidth(56 if compact else _ACTION_BUTTON_MIN_WIDTH)
        self._clear_button.setMinimumWidth(36 if compact else _ACTION_BUTTON_MIN_WIDTH)
        self.height_changed.emit(self.minimumHeight())

    def _sync_supported_type_layout(self) -> None:
        if not self._supported_type_layout_objects_are_valid():
            return
        fallback_width = self._drop_group.width() - (_SPACING_MD * 4)
        panel_width = self._empty_center_panel.width()
        content_width = max(panel_width if panel_width > 0 else fallback_width, 0)
        middle_gap = max(_SPACING_MD, 24)
        for index, (_, row_layout, type_label, value_label) in enumerate(self._type_prompt_rows):
            desired_indent = _PYRAMID_INDENTS[min(index, len(_PYRAMID_INDENTS) - 1)]
            required_width = (
                type_label.sizeHint().width() + value_label.sizeHint().width() + row_layout.spacing() + middle_gap
            )
            available_indent = max((content_width - required_width) // 2, 0)
            actual_indent = min(desired_indent, available_indent)
            row_layout.setContentsMargins(actual_indent, 0, actual_indent, 0)
        self._types_container.updateGeometry()

    def _supported_type_layout_objects_are_valid(self) -> bool:
        if not shiboken6.isValid(self):
            return False
        objects = [
            getattr(self, "_drop_group", None),
            getattr(self, "_empty_center_panel", None),
            getattr(self, "_types_container", None),
        ]
        for _row_widget, row_layout, type_label, value_label in getattr(self, "_type_prompt_rows", ()):
            objects.extend((row_layout, type_label, value_label))
        return all(shiboken6.isValid(obj) for obj in objects)

    # ── Public API ─────────────────────────────────────────────────

    @property
    def view_model(self) -> InputAreaViewModel:
        """The ViewModel driving this widget."""
        return self._vm

    @property
    def add_button(self) -> PrimaryPushButton:
        """Public access to the add button for external wiring."""
        return self._add_button

    @property
    def clear_button(self) -> PushButton:
        """Public access to the clear button for external wiring."""
        return self._clear_button

    @property
    def mode_switch(self) -> SegmentedWidget:
        """Public access to the mode switch for external wiring."""
        return self._mode_switch

    def open_file_dialog(self, *, force_batch_mode: bool = False) -> None:
        """Public entry point: open the file dialog (single or multi, per mode).

        Exposed for external callers such as the main-window shortcut
        handler (Ctrl+O).
        """
        self._vm.request_add_dialog(force_batch_mode=force_batch_mode)
        self._open_file_dialog()

    def open_folder_dialog(self, *, force_batch_mode: bool = False) -> None:
        """Public entry point: open a folder dialog.

        Exposed for external callers such as the main-window shortcut
        handler (Ctrl+Shift+O).
        """
        self._vm.request_add_folder_dialog(force_batch_mode=force_batch_mode)
        self._open_folder_dialog()

    def set_recent_files(self, paths: list[str]) -> None:
        """Pre-populate the recent files list (called by main window)."""
        self._recent_files = [p for p in paths if Path(p).exists()]

    def update_display(self, file_path: str) -> None:
        """Update the widget to reflect a single selected file.

        Called by external code (e.g. main window) when a file is
        selected through other means (IPC, etc.).
        Delegates to ViewModel so selection state stays consistent.

        Args:
            file_path: Absolute path of the selected file.
        """
        if not file_path:
            return
        self._selected_file = file_path
        self._vm.add_files([file_path])


__all__ = [
    "_COMPACT_WIDTH_THRESHOLD",
    "_DEFAULT_HEIGHT",
    "InputArea",
]
