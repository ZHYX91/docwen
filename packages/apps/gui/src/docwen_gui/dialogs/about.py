"""About dialog — version, copyright, acknowledgments, and tool list."""

from __future__ import annotations

import logging

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeyEvent
from PySide6.QtWidgets import (
    QDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QScrollArea,
    QToolButton,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

from ..i18n import t
from ..resources import load_svg_icon

logger = logging.getLogger(__name__)

# Tool list for acknowledgments
_TOOLS_LEFT: list[tuple[str, str]] = [
    ("Python-docx", "python_docx"),
    ("openpyxl", "openpyxl"),
    ("PyMuPDF (fitz)", "pymupdf"),
    ("pymupdf4llm", "pymupdf4llm"),
    ("pdf2docx", "pdf2docx"),
    ("easyofd", "easyofd"),
    ("RapidOCR", "rapidocr"),
    ("PaddleOCR", "paddleocr"),
    ("ONNX Runtime", "onnxruntime"),
    ("PySide6", "pyside6"),
    ("PySide6-Fluent-Widgets", "pyside6"),
    ("Pillow (PIL)", "pillow"),
]

_TOOLS_RIGHT: list[tuple[str, str]] = [
    ("pillow-heif", "pillow_heif"),
    ("img2pdf", "img2pdf"),
    ("pywin32", "pywin32"),
    ("lxml", "lxml"),
    ("latex2mathml", "latex2mathml"),
    ("PyYAML", "pyyaml"),
    ("tomlkit", "tomlkit"),
    ("pandas", "pandas"),
    ("numpy", "numpy"),
    ("olefile", "olefile"),
    ("emoji", "emoji"),
]


class _ToolInfoButton(QToolButton):
    """Keyboard-reachable disclosure button for one tool description."""

    description_requested = Signal(str)

    def __init__(self, description: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._description = description
        self.clicked.connect(self.show_description)

    def show_description(self, _checked: bool = False) -> None:
        if not self._description:
            return
        self.description_requested.emit(self._description)
        QToolTip.showText(
            self.mapToGlobal(self.rect().bottomLeft()),
            self._description,
            self,
        )

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if event.key() in {Qt.Key.Key_Enter, Qt.Key.Key_Return, Qt.Key.Key_Space}:
            self.show_description()
            event.accept()
            return
        super().keyPressEvent(event)


def _tool_tip_label(name: str, tooltip_key: str) -> QLabel:
    lbl = QLabel(name)
    tooltip = t(f"about.tools.{tooltip_key}", default="")
    if tooltip:
        lbl.setToolTip(tooltip)
    return lbl


def _tool_info_button(tooltip_key: str) -> QToolButton:
    tooltip = t(f"about.tools.{tooltip_key}", default="")
    btn = _ToolInfoButton(tooltip)
    btn.setObjectName("aboutToolInfoButton")
    btn.setAutoRaise(True)
    btn.setFixedSize(18, 18)
    btn.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
    btn.setToolTip(tooltip)
    btn.setAccessibleName(tooltip)
    btn.setAccessibleDescription(tooltip)

    icon = load_svg_icon("about.svg")
    if icon is not None and not icon.isNull():
        btn.setIcon(icon)
    else:
        btn.setText("i")
    return btn


def _tool_entry_widget(name: str, tooltip_key: str) -> QWidget:
    widget = QWidget()
    widget.setObjectName("aboutToolEntry")
    layout = QHBoxLayout(widget)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(4)

    label = _tool_tip_label(name, tooltip_key)
    layout.addWidget(label)
    layout.addWidget(_tool_info_button(tooltip_key))
    layout.addStretch(1)
    return widget


class AboutDialog(QDialog):
    """About dialog showing version, copyright, acknowledgments, and tool list."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("aboutDialog")
        self.setModal(True)
        self.setFixedSize(440, 680)
        self.setWindowTitle(t("about.title", default="About DocWen"))
        self._create_interface()
        self._center_on_parent()
        logger.info("About dialog initialized")

    def _center_on_parent(self) -> None:
        parent_widget = self.parentWidget()
        if parent_widget is not None:
            parent_geo = parent_widget.geometry()
            x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
            y = parent_geo.y() + (parent_geo.height() - self.height()) // 2
            self.move(x, y)

    def _create_interface(self) -> None:
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(8)

        # Scroll area for main content
        scroll = QScrollArea(self)
        scroll.setObjectName("aboutScrollArea")
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        root_layout.addWidget(scroll, 1)

        content = QWidget(scroll)
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(8)
        content_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        scroll.setWidget(content)

        # ── Hero card ──────────────────────────────────────────────
        hero_card = QWidget(content)
        hero_card.setObjectName("aboutHeroCard")
        hero_layout = QVBoxLayout(hero_card)
        hero_layout.setContentsMargins(16, 16, 16, 16)
        hero_layout.setSpacing(4)
        hero_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)

        title_label = QLabel(t("common.app_name", default="DocWen"))
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setObjectName("aboutTitle")
        hero_layout.addWidget(title_label)

        subtitle = t("about.subtitle", default="")
        if subtitle:
            sub_label = QLabel(subtitle)
            sub_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            sub_label.setObjectName("aboutSubtitle")
            hero_layout.addWidget(sub_label)

        from docwen_gui import __version__

        ver_label = QLabel(t("about.version_info", version=__version__))
        ver_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ver_label.setObjectName("aboutVersion")
        update_notice = t(
            "about.version_tooltip",
            default="This software is offline-only. Please check for updates manually.",
        )
        ver_label.setToolTip(update_notice)
        hero_layout.addWidget(ver_label)

        if update_notice:
            update_label = QLabel(update_notice)
            update_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            update_label.setWordWrap(True)
            update_label.setObjectName("aboutUpdateNotice")
            hero_layout.addWidget(update_label)

        contact_label = QLabel(t("about.contact", email="zhengyx91@hotmail.com"))
        contact_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        contact_label.setObjectName("aboutMeta")
        hero_layout.addWidget(contact_label)

        copyright_label = QLabel(t("about.copyright", year="2025-2026", author="ZhengYX"))
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        copyright_label.setObjectName("aboutMeta")
        hero_layout.addWidget(copyright_label)

        content_layout.addWidget(hero_card)

        # ── Disclaimer card ────────────────────────────────────────
        disclaimer_card = QWidget(content)
        disclaimer_card.setObjectName("aboutGroup")
        disclaimer_layout = QVBoxLayout(disclaimer_card)
        disclaimer_layout.setContentsMargins(16, 16, 16, 16)
        disclaimer_label = QLabel(t("common.disclaimer", default=""))
        disclaimer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        disclaimer_label.setWordWrap(True)
        disclaimer_label.setObjectName("warningLabel")
        disclaimer_layout.addWidget(disclaimer_label)
        content_layout.addWidget(disclaimer_card)

        # ── Acknowledgments card ───────────────────────────────────
        ack_card = QWidget(content)
        ack_card.setObjectName("aboutGroup")
        ack_layout = QVBoxLayout(ack_card)
        ack_layout.setContentsMargins(16, 16, 16, 16)

        ack_title = QLabel(t("about.acknowledgments", default="Acknowledgments"))
        ack_title.setObjectName("aboutGroupTitle")
        ack_layout.addWidget(ack_title)

        intro = QLabel(t("about.acknowledgments_intro", default=""))
        intro.setWordWrap(True)
        intro.setObjectName("aboutAcknowledgmentsIntro")
        ack_layout.addWidget(intro)

        tools_widget = QWidget()
        tools_widget.setObjectName("aboutToolsGrid")
        tools_grid = QGridLayout(tools_widget)
        tools_grid.setSpacing(4)
        tools_grid.setHorizontalSpacing(12)

        for row, (name, tooltip_key) in enumerate(_TOOLS_LEFT):
            tools_grid.addWidget(_tool_entry_widget(name, tooltip_key), row, 0, Qt.AlignmentFlag.AlignLeft)

        for row, (name, tooltip_key) in enumerate(_TOOLS_RIGHT):
            tools_grid.addWidget(_tool_entry_widget(name, tooltip_key), row, 1, Qt.AlignmentFlag.AlignLeft)

        ack_layout.addWidget(tools_widget)
        content_layout.addWidget(ack_card)

        # ── Close button ───────────────────────────────────────────
        try:
            from qfluentwidgets import PushButton as QFluentPushButton

            close_btn = QFluentPushButton(t("common.close", default="Close"))
        except ImportError:
            from PySide6.QtWidgets import QPushButton

            close_btn = QPushButton(t("common.close", default="Close"))
        close_btn.setObjectName("aboutCloseButton")
        close_btn.setFixedWidth(120)
        close_btn.clicked.connect(self.accept)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        btn_row.addStretch()
        root_layout.addLayout(btn_row)

    def show_dialog(self) -> None:
        """Show the dialog modally and wait for it to close."""
        logger.debug("Showing about dialog")
        self._center_on_parent()
        self.exec()
