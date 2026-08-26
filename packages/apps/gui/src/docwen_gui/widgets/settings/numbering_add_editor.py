"""Heading numbering addition scheme editor.

Port of the old ``NumberingAddDialog`` — scheme CRUD, per-level format
string editor with placeholder menu, real-time preview, and TOML
persistence.

Placed in ``widgets/settings/``; opened from ``TextTab`` editor buttons.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, ClassVar

from PySide6.QtCore import QEvent, Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFormLayout,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from ...i18n import t

# ── Theme helpers ────────────────────────────────────────────────────────


def _theme_class_color(theme_class: str) -> str:
    """Resolve a semantic class color for the current theme."""
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.styles.theme_semantics import get_theme_class_color

    theme_name = ThemeManager.get_instance().get_current_theme()
    return get_theme_class_color(theme_class, theme_name)


# ── Data ─────────────────────────────────────────────────────────────────


@dataclass
class NumberingScheme:
    """A single heading numbering scheme definition."""

    scheme_id: str
    name: str = ""
    description: str = ""
    enabled: bool = True
    is_system: bool = False
    locales: list[str] = field(default_factory=lambda: ["*"])
    name_key: str = ""
    description_key: str = ""
    levels: dict[int, str] = field(default_factory=lambda: dict.fromkeys(range(1, 10), ""))

    @classmethod
    def from_config(cls, scheme_id: str, config: dict[str, Any]) -> NumberingScheme:
        levels = dict.fromkeys(range(1, 10), "")
        for idx in range(1, 10):
            level_config = config.get(f"level_{idx}", {})
            if isinstance(level_config, dict):
                levels[idx] = str(level_config.get("format", ""))
        return cls(
            scheme_id=scheme_id,
            name=str(config.get("name", "")),
            description=str(config.get("description", "")),
            enabled=bool(config.get("enabled", True)),
            is_system=bool(config.get("is_system", False)),
            locales=[str(item) for item in config.get("locales", ["*"]) if isinstance(item, str)],
            name_key=str(config.get("name_key", "")),
            description_key=str(config.get("description_key", "")),
            levels=levels,
        )

    def display_name(self) -> str:
        if self.name_key:
            translated = t(f"editors.numbering_add.names.{self.name_key}")
            if translated != f"[editors.numbering_add.names.{self.name_key}]":
                return translated
        return self.name or self.scheme_id

    def display_description(self) -> str:
        if self.description_key:
            translated = t(f"editors.numbering_add.descriptions.{self.description_key}")
            if translated != f"[editors.numbering_add.descriptions.{self.description_key}]":
                return translated
        return self.description

    def copy(self, new_id: str, new_name: str) -> NumberingScheme:
        return NumberingScheme(
            scheme_id=new_id,
            name=new_name,
            description=self.display_description(),
            enabled=True,
            is_system=False,
            locales=["*"],
            levels=dict(self.levels),
        )


# ── Dialog ───────────────────────────────────────────────────────────────


class NumberingAddDialog(QDialog):
    """Dialog for managing heading numbering addition schemes.

    Features:
    - Scheme list with CRUD (create/copy/delete) and reorder (move up/down).
    - Default-scheme combo selector.
    - 9-level format string editor with placeholder insertion menu.
    - Real-time numbering preview.
    - Dirty-state tracking; save hands data to ``on_save`` callback (the
      Settings view model persists through its injected config port).
    """

    PLACEHOLDER_PATTERN: ClassVar[re.Pattern[str]] = re.compile(r"\{(\d+)\.(\w+)\}")

    DEFAULT_STYLES: ClassVar[frozenset[str]] = frozenset(
        {
            "chinese_lower",
            "chinese_upper",
            "arabic_half",
            "arabic_full",
            "arabic_circled",
            "letter_upper",
            "letter_lower",
            "roman_upper",
            "roman_lower",
        }
    )

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        config_data: dict[str, Any] | None = None,
        on_save: Any | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("numberingAddDialog")
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self._on_save = on_save

        self.number_styles: dict[str, dict[str, Any]] = {}
        self.schemes: dict[str, NumberingScheme] = {}
        self.order: list[str] = []
        self.default_scheme_id = ""
        self.current_scheme_id: str | None = None
        self.level_status_labels: dict[int, QLabel] = {}
        self._loading_form = False
        self._dirty = False
        self._saved_state: tuple[tuple[str, ...], str, tuple[tuple[str, tuple[Any, ...]], ...]] = ((), "", ())
        self._base_window_title = t("editors.numbering_add.window_title", "Numbering Scheme Editor")

        self.setWindowTitle(self._base_window_title)
        self.resize(1000, 720)

        self._build_ui()

        if config_data is not None:
            self._load_from_dict(config_data)

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() == QEvent.Type.PaletteChange and hasattr(self, "_compat_label"):
            self._update_word_native_compatibility()

    # ── UI construction ──────────────────────────────────────────────────

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)

        content = QHBoxLayout()
        root.addLayout(content, 1)

        # -- left panel: scheme list + buttons --
        left = QVBoxLayout()
        content.addLayout(left, 1)
        left.addWidget(QLabel(t("editors.numbering_add.scheme_list"), self))

        self.scheme_list = QListWidget(self)
        self.scheme_list.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self.scheme_list.currentRowChanged.connect(self._on_scheme_selected)
        left.addWidget(self.scheme_list, 1)

        left_btns = QHBoxLayout()
        self.new_btn = QPushButton(t("editors.numbering_add.new_scheme"), self)
        self.new_btn.clicked.connect(self._create_new_scheme)
        left_btns.addWidget(self.new_btn)
        self.copy_btn = QPushButton(t("editors.common.copy"), self)
        self.copy_btn.clicked.connect(self._copy_selected_scheme)
        left_btns.addWidget(self.copy_btn)
        self.delete_btn = QPushButton(t("editors.common.delete"), self)
        self.delete_btn.clicked.connect(self._delete_selected_scheme)
        left_btns.addWidget(self.delete_btn)
        left.addLayout(left_btns)

        order_btns = QHBoxLayout()
        self.up_btn = QPushButton(t("editors.common.move_up"), self)
        self.up_btn.clicked.connect(self._move_selected_scheme_up)
        order_btns.addWidget(self.up_btn)
        self.down_btn = QPushButton(t("editors.common.move_down"), self)
        self.down_btn.clicked.connect(self._move_selected_scheme_down)
        order_btns.addWidget(self.down_btn)
        left.addLayout(order_btns)

        # -- right panel: form + level grid + preview --
        right = QVBoxLayout()
        content.addLayout(right, 2)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel(t("editors.numbering_add.default_scheme"), self))
        self.default_combo = QComboBox(self)
        self._prepare_combo(self.default_combo)
        self.default_combo.currentIndexChanged.connect(self._on_default_scheme_changed)
        top_row.addWidget(self.default_combo, 1)
        help_btn = QPushButton(t("editors.numbering_add.format_help"), self)
        help_btn.clicked.connect(self._show_format_help)
        top_row.addWidget(help_btn)
        right.addLayout(top_row)

        form = QFormLayout()
        self.name_edit = QLineEdit(self)
        self.name_edit.textChanged.connect(self._on_name_changed)
        form.addRow(t("editors.common.name"), self.name_edit)

        self.desc_edit = QLineEdit(self)
        self.desc_edit.textChanged.connect(self._on_description_changed)
        form.addRow(t("editors.common.description"), self.desc_edit)
        right.addLayout(form)

        right.addWidget(QLabel(t("editors.numbering_add.level_format_config"), self))
        level_grid = QGridLayout()
        self.level_edits: dict[int, QLineEdit] = {}
        for idx in range(1, 10):
            label = QLabel(t(f"editors.numbering_add.level_{idx}"), self)
            edit = QLineEdit(self)
            edit.textChanged.connect(lambda text, level=idx: self._on_level_changed(level, text))
            placeholder_btn = QToolButton(self)
            placeholder_btn.setText("+")
            placeholder_btn.clicked.connect(lambda _checked=False, level=idx: self._show_placeholder_menu(level))
            status_label = QLabel("", self)
            status_label.setMinimumWidth(24)
            self.level_edits[idx] = edit
            self.level_status_labels[idx] = status_label
            row = idx - 1
            level_grid.addWidget(label, row, 0)
            level_grid.addWidget(edit, row, 1)
            level_grid.addWidget(placeholder_btn, row, 2)
            level_grid.addWidget(status_label, row, 3)
        right.addLayout(level_grid)

        # Word native compatibility display (read-only)
        self._compat_label = QLabel("", self)
        self._compat_label.setWordWrap(True)
        right.addWidget(self._compat_label)

        right.addWidget(QLabel(t("editors.numbering_add.realtime_preview"), self))
        self.preview_text = QTextEdit(self)
        self.preview_text.setReadOnly(True)
        right.addWidget(self.preview_text, 1)

        # -- bottom bar --
        bottom = QHBoxLayout()
        bottom.addStretch(1)
        self.save_btn = QPushButton(t("common.save"), self)
        self.save_btn.clicked.connect(self._save_to_disk)
        bottom.addWidget(self.save_btn)
        self.cancel_btn = QPushButton(t("common.cancel"), self)
        self.cancel_btn.clicked.connect(self.reject)
        bottom.addWidget(self.cancel_btn)
        root.addLayout(bottom)

    # ── Load ─────────────────────────────────────────────────────────────

    def _load_from_dict(self, data: dict[str, Any]) -> None:
        settings = data.get("settings", {})
        if not isinstance(settings, dict):
            settings = {}
        self.default_scheme_id = str(settings.get("default_scheme", ""))
        self.order = [str(item) for item in settings.get("order", []) if isinstance(item, str)]
        number_styles = data.get("number_styles", {})
        self.number_styles = number_styles if isinstance(number_styles, dict) else {}

        schemes_data = data.get("schemes", {})
        self.schemes = {}
        if isinstance(schemes_data, dict):
            for scheme_id, scheme_config in schemes_data.items():
                if isinstance(scheme_config, dict):
                    self.schemes[str(scheme_id)] = NumberingScheme.from_config(str(scheme_id), scheme_config)

        self.order = [sid for sid in self.order if sid in self.schemes]
        for sid in self.schemes:
            if sid not in self.order:
                self.order.append(sid)
        if not self.default_scheme_id or self.default_scheme_id not in self.schemes:
            self.default_scheme_id = self.order[0] if self.order else ""

        self._refresh_scheme_views()
        selected = self.default_scheme_id or (self.order[0] if self.order else None)
        if selected:
            self._select_scheme(selected)
        self._saved_state = self._capture_state()
        self._refresh_dirty_state()

    # ── Scheme list views ────────────────────────────────────────────────

    def _refresh_scheme_views(self) -> None:
        selected_id = self.current_scheme_id

        self.scheme_list.blockSignals(True)
        self.scheme_list.clear()
        for sid in self.order:
            display = self.schemes[sid].display_name()
            item = QListWidgetItem(display)
            item.setToolTip(display)
            self.scheme_list.addItem(item)
        self.scheme_list.blockSignals(False)

        self.default_combo.blockSignals(True)
        self.default_combo.clear()
        for sid in self.order:
            display = self.schemes[sid].display_name()
            self._combo_add_item(self.default_combo, display, sid)
        self._combo_set_data(self.default_combo, self.default_scheme_id)
        self.default_combo.blockSignals(False)

        if selected_id and selected_id in self.order:
            self._select_scheme(selected_id)
        else:
            self._refresh_action_states()

    def _select_scheme(self, scheme_id: str) -> None:
        if scheme_id not in self.schemes:
            return
        self.current_scheme_id = scheme_id
        row = self.order.index(scheme_id)
        self.scheme_list.blockSignals(True)
        self.scheme_list.setCurrentRow(row)
        self.scheme_list.blockSignals(False)
        self._populate_form(self.schemes[scheme_id])

    def _on_scheme_selected(self, row: int) -> None:
        if row < 0 or row >= len(self.order):
            return
        self.current_scheme_id = self.order[row]
        self._populate_form(self.schemes[self.current_scheme_id])

    def _populate_form(self, scheme: NumberingScheme) -> None:
        self._loading_form = True
        try:
            self.name_edit.setText(scheme.name or scheme.display_name())
            self.desc_edit.setText(scheme.description or scheme.display_description())
            is_custom = not scheme.is_system
            self.name_edit.setEnabled(is_custom)
            self.desc_edit.setEnabled(is_custom)
            for idx in range(1, 10):
                self.level_edits[idx].setText(scheme.levels.get(idx, ""))
            self._validate_all_levels()
            self._update_preview()
            self._update_word_native_compatibility(scheme)
            self._refresh_action_states()
        finally:
            self._loading_form = False

    # ── Dirty tracking ───────────────────────────────────────────────────

    def _mark_dirty(self) -> None:
        self._refresh_dirty_state()

    def _capture_state(self) -> tuple[tuple[str, ...], str, tuple[tuple[str, tuple[Any, ...]], ...]]:
        items: list[tuple[str, tuple[Any, ...]]] = []
        for sid in sorted(self.schemes):
            s = self.schemes[sid]
            items.append(
                (
                    sid,
                    (
                        s.name,
                        s.description,
                        bool(s.enabled),
                        bool(s.is_system),
                        tuple(s.locales),
                        s.name_key,
                        s.description_key,
                        tuple((idx, s.levels.get(idx, "")) for idx in range(1, 10)),
                    ),
                )
            )
        return (tuple(self.order), self.default_scheme_id, tuple(items))

    def _refresh_dirty_state(self) -> None:
        self._dirty = self._capture_state() != self._saved_state
        self.setWindowTitle(f"{self._base_window_title} *" if self._dirty else self._base_window_title)
        self.save_btn.setEnabled(self._dirty and self._can_save())

    def _can_save(self) -> bool:
        return not self._has_validation_errors()

    def _has_validation_errors(self) -> bool:
        return any(label.text() == "x" for label in self.level_status_labels.values())

    def _refresh_action_states(self) -> None:
        scheme = self._current_scheme()
        has = scheme is not None and scheme.scheme_id in self.order
        if not has or scheme is None:
            self.delete_btn.setEnabled(False)
            self.up_btn.setEnabled(False)
            self.down_btn.setEnabled(False)
            return
        idx = self.order.index(scheme.scheme_id)
        self.delete_btn.setEnabled((not scheme.is_system) and scheme.scheme_id != self.default_scheme_id)
        self.up_btn.setEnabled(idx > 0)
        self.down_btn.setEnabled(idx < len(self.order) - 1)

    def _current_scheme(self) -> NumberingScheme | None:
        if not self.current_scheme_id:
            return None
        return self.schemes.get(self.current_scheme_id)

    def _update_word_native_compatibility(self, scheme: NumberingScheme | None = None) -> None:
        """Update the read-only Word native compatibility display.

        Uses a lazy import from the markdown plugin to avoid hard
        dependency at module load time.
        """
        if scheme is None:
            scheme = self._current_scheme()
        if scheme is None:
            self._compat_label.setText("")
            return

        # Build the scheme_config dict from the scheme's levels
        scheme_config: dict[str, dict[str, str]] = {}
        for idx in range(1, 10):
            fmt = scheme.levels.get(idx, "")
            if fmt:
                scheme_config[f"level_{idx}"] = {"format": fmt}

        if not scheme_config:
            self._compat_label.setText("")
            return

        try:
            from docwen_core.text.numbering_word_adapter import translate_scheme
        except ImportError:
            self._compat_label.setText(t("editors.numbering_add.word_native_unavailable"))
            return

        result = translate_scheme(scheme_config)

        if result.verdict == "full":
            self._compat_label.setText(t("editors.numbering_add.word_native_full"))
            self._compat_label.setStyleSheet(f"color: {_theme_class_color('success')};")
        elif result.verdict == "approximate":
            self._compat_label.setText(t("editors.numbering_add.word_native_approximate", reason=result.reason))
            self._compat_label.setStyleSheet(f"color: {_theme_class_color('warning')};")
        else:
            self._compat_label.setText(t("editors.numbering_add.word_native_incompatible", reason=result.reason))
            self._compat_label.setStyleSheet(f"color: {_theme_class_color('danger')};")

    # ── Field change handlers ────────────────────────────────────────────

    def _on_name_changed(self, text: str) -> None:
        if self._loading_form:
            return
        scheme = self._current_scheme()
        if scheme is None or scheme.is_system:
            return
        scheme.name = text.strip()
        self._refresh_scheme_views()
        self._mark_dirty()

    def _on_description_changed(self, text: str) -> None:
        if self._loading_form:
            return
        scheme = self._current_scheme()
        if scheme is None or scheme.is_system:
            return
        scheme.description = text.strip()
        self._mark_dirty()

    def _on_level_changed(self, level: int, text: str) -> None:
        if self._loading_form:
            return
        scheme = self._current_scheme()
        if scheme is None:
            return
        scheme.levels[level] = text
        self._validate_level(level, text)
        self._update_preview()
        self._update_word_native_compatibility(scheme)
        self._mark_dirty()

    def _on_default_scheme_changed(self) -> None:
        data = self._combo_current_data(self.default_combo)
        if not isinstance(data, str) or data not in self.schemes:
            return
        self.default_scheme_id = data
        self._refresh_action_states()
        self._mark_dirty()

    # ── Scheme CRUD ──────────────────────────────────────────────────────

    def _generate_unique_id(self, base: str) -> str:
        candidate = re.sub(r"[^a-zA-Z0-9_]+", "_", base).strip("_").lower() or "custom_scheme"
        if candidate not in self.schemes:
            return candidate
        idx = 2
        while f"{candidate}_{idx}" in self.schemes:
            idx += 1
        return f"{candidate}_{idx}"

    def _create_new_scheme(self) -> None:
        base = t("editors.numbering_add.new_scheme")
        sid = self._generate_unique_id(base)
        scheme = NumberingScheme(scheme_id=sid, name=base)
        self.schemes[sid] = scheme
        self.order.append(sid)
        self._refresh_scheme_views()
        self._select_scheme(sid)
        self._mark_dirty()

    def _copy_selected_scheme(self) -> None:
        scheme = self._current_scheme()
        if scheme is None:
            self._notify_info(t("editors.common.select_item_first"))
            return
        new_name = f"{scheme.display_name()} {t('editors.numbering_add.copy_suffix')}"
        sid = self._generate_unique_id(f"{scheme.scheme_id}_copy")
        copied = scheme.copy(sid, new_name)
        self.schemes[sid] = copied
        insert_at = self.order.index(scheme.scheme_id) + 1
        self.order.insert(insert_at, sid)
        self._refresh_scheme_views()
        self._select_scheme(sid)
        self._mark_dirty()

    def _delete_selected_scheme(self) -> None:
        scheme = self._current_scheme()
        if scheme is None:
            self._notify_info(t("editors.common.select_item_first"))
            return
        if scheme.scheme_id == self.default_scheme_id:
            self._notify_warning(t("editors.numbering_add.cannot_delete_default_scheme"))
            return
        if scheme.is_system:
            self._notify_warning(t("editors.common.system_item_hint"))
            return
        if not self._confirm(
            t("editors.common.confirm_delete"),
            t("editors.common.confirm_delete_message", name=scheme.display_name()),
        ):
            return
        self.schemes.pop(scheme.scheme_id, None)
        self.order = [item for item in self.order if item != scheme.scheme_id]
        next_sid = self.default_scheme_id or (self.order[0] if self.order else None)
        self.current_scheme_id = None
        self._refresh_scheme_views()
        if next_sid:
            self._select_scheme(next_sid)
        self._mark_dirty()

    def _move_selected_scheme_up(self) -> None:
        scheme = self._current_scheme()
        if scheme is None:
            return
        idx = self.order.index(scheme.scheme_id)
        if idx <= 0:
            return
        self.order[idx - 1], self.order[idx] = self.order[idx], self.order[idx - 1]
        self._refresh_scheme_views()
        self._select_scheme(scheme.scheme_id)
        self._mark_dirty()

    def _move_selected_scheme_down(self) -> None:
        scheme = self._current_scheme()
        if scheme is None:
            return
        idx = self.order.index(scheme.scheme_id)
        if idx >= len(self.order) - 1:
            return
        self.order[idx + 1], self.order[idx] = self.order[idx], self.order[idx + 1]
        self._refresh_scheme_views()
        self._select_scheme(scheme.scheme_id)
        self._mark_dirty()

    # ── Format help & placeholder menu ───────────────────────────────────

    def _show_format_help(self) -> None:
        self._notify_info(
            t("editors.numbering_add.format_help_text"),
            title=t("editors.numbering_add.format_help_title"),
        )

    def _show_placeholder_menu(self, level: int) -> None:
        edit = self.level_edits.get(level)
        if edit is None:
            return

        from PySide6.QtWidgets import QMenu

        menu = QMenu(self)

        header = menu.addAction(t("editors.numbering_add.menu_current_level", level=level))
        header.setEnabled(False)

        styles = [
            ("chinese_lower", t("editors.numbering_add.style_chinese_lower")),
            ("chinese_upper", t("editors.numbering_add.style_chinese_upper")),
            ("arabic_half", t("editors.numbering_add.style_arabic_half")),
            ("arabic_circled", t("editors.numbering_add.style_arabic_circled")),
            ("letter_upper", t("editors.numbering_add.style_letter_upper")),
            ("roman_upper", t("editors.numbering_add.style_roman_upper")),
        ]
        for style_id, style_name in styles:
            placeholder = f"{{{level}.{style_id}}}"
            menu.addAction(
                f"{placeholder}  {style_name}",
                lambda _checked=False, value=placeholder: self._insert_placeholder(level, value),
            )

        menu.addSeparator()
        decorations_header = menu.addAction(t("editors.numbering_add.menu_decorations"))
        decorations_header.setEnabled(False)
        for decoration, desc in self._decoration_templates(level):
            menu.addAction(
                f"{decoration}  {desc}",
                lambda _checked=False, value=decoration: self._insert_placeholder(level, value),
            )

        if level > 1:
            menu.addSeparator()
            hier_header = menu.addAction(t("editors.numbering_add.menu_hierarchical"))
            hier_header.setEnabled(False)
            hierarchical = ".".join(f"{{{idx}.arabic_half}}" for idx in range(1, level + 1)) + " "
            menu.addAction(
                hierarchical,
                lambda _checked=False, value=hierarchical: self._insert_placeholder(level, value),
            )

        menu.exec(edit.mapToGlobal(edit.rect().bottomRight()))

    def _insert_placeholder(self, level: int, text: str) -> None:
        edit = self.level_edits.get(level)
        if edit is None:
            return
        edit.insert(text)
        edit.setFocus()

    # ── Validation ───────────────────────────────────────────────────────

    def _validate_level(self, level: int, format_str: str) -> None:
        label = self.level_status_labels.get(level)
        if label is None:
            return
        if not format_str.strip():
            label.setText("")
            label.setToolTip("")
            return

        placeholders = self.PLACEHOLDER_PATTERN.findall(format_str)
        if not placeholders:
            label.setText("!")
            label.setToolTip(t("editors.numbering_add.validation_fixed_text"))
            return

        valid_styles = set(self.number_styles) or self.DEFAULT_STYLES
        for ref_level_str, style in placeholders:
            try:
                ref_level = int(ref_level_str)
            except ValueError:
                label.setText("x")
                label.setToolTip(t("editors.numbering_add.validation_invalid_level", level=ref_level_str))
                return
            if not 1 <= ref_level <= 9:
                label.setText("x")
                label.setToolTip(t("editors.numbering_add.validation_level_out_of_range", level=ref_level))
                return
            if style not in valid_styles:
                label.setText("x")
                label.setToolTip(t("editors.numbering_add.validation_unknown_style", style=style))
                return

        label.setText("✓")
        label.setToolTip(t("editors.numbering_add.validation_ok"))

    def _validate_all_levels(self) -> None:
        scheme = self._current_scheme()
        if scheme is None:
            return
        for level in range(1, 10):
            self._validate_level(level, scheme.levels.get(level, ""))

    # ── Preview ──────────────────────────────────────────────────────────

    @staticmethod
    def _roman(value: int) -> str:
        pairs = [
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I"),
        ]
        result = []
        remaining = value
        for arabic, roman in pairs:
            while remaining >= arabic:
                result.append(roman)
                remaining -= arabic
        return "".join(result)

    @staticmethod
    def _style_sample(style: str, value: int) -> str:
        chinese_lower = ["一", "二", "三", "四", "五", "六", "七", "八", "九"]
        chinese_upper = ["壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
        circled = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨"]
        if style == "chinese_lower":
            return chinese_lower[value - 1] if value <= len(chinese_lower) else str(value)
        if style == "chinese_upper":
            return chinese_upper[value - 1] if value <= len(chinese_upper) else str(value)
        if style == "arabic_half":
            return str(value)
        if style == "arabic_full":
            return "".join(chr(ord("０") + int(ch)) for ch in str(value))
        if style == "arabic_circled":
            return circled[value - 1] if value <= len(circled) else str(value)
        if style == "letter_upper":
            return chr(ord("A") + value - 1) if 1 <= value <= 26 else str(value)
        if style == "letter_lower":
            return chr(ord("a") + value - 1) if 1 <= value <= 26 else str(value)
        if style == "roman_upper":
            return NumberingAddDialog._roman(value)
        if style == "roman_lower":
            return NumberingAddDialog._roman(value).lower()
        return str(value)

    def _render_format(self, format_str: str) -> str:
        def replacer(match: re.Match[str]) -> str:
            level = int(match.group(1))
            style = match.group(2)
            return self._style_sample(style, level)

        return self.PLACEHOLDER_PATTERN.sub(replacer, format_str)

    def _update_preview(self) -> None:
        scheme = self._current_scheme()
        if scheme is None:
            self.preview_text.clear()
            return
        lines: list[str] = []
        for idx in range(1, 10):
            level_name = t(f"editors.numbering_add.level_{idx}")
            prefix = scheme.levels.get(idx, "")
            title = t("editors.numbering_add.preview_sample", level_name=level_name)
            if prefix:
                rendered = self._render_format(prefix)
                lines.append(f"{rendered}{title}")
            else:
                lines.append(f"{title} ({t('editors.numbering_add.no_format')})")
        self.preview_text.setPlainText("\n".join(lines))

    @staticmethod
    def _decoration_templates(level: int) -> list[tuple[str, str]]:
        return [
            ("、", t("editors.numbering_add.decoration_dun")),
            (". ", t("editors.numbering_add.decoration_dot_space")),
            (f"（{{{level}.chinese_lower}}）", t("editors.numbering_add.decoration_bracket")),
            (f"第{{{level}.arabic_half}}章 ", t("editors.numbering_add.decoration_chapter")),
        ]

    # ── Save / Cancel ────────────────────────────────────────────────────

    def _save_to_disk(self) -> bool:
        if not self._can_save():
            return False

        doc_data = self._build_toml_dict()
        if callable(self._on_save) and self._on_save(doc_data) is False:
            return False

        self._saved_state = self._capture_state()
        self._refresh_dirty_state()
        self.accept()
        return True

    def _build_toml_dict(self) -> dict[str, Any]:
        return {
            "settings": {
                "default_scheme": self.default_scheme_id,
                "order": list(self.order),
            },
            "number_styles": dict(self.number_styles),
            "schemes": {sid: self._scheme_to_dict(self.schemes[sid]) for sid in self.order},
        }

    @staticmethod
    def _scheme_to_dict(scheme: NumberingScheme) -> dict[str, Any]:
        d: dict[str, Any] = {}
        if scheme.name_key:
            d["name_key"] = scheme.name_key
        else:
            d["name"] = scheme.name
        if scheme.description_key:
            d["description_key"] = scheme.description_key
        else:
            d["description"] = scheme.description
        d["enabled"] = bool(scheme.enabled)
        d["is_system"] = bool(scheme.is_system)
        d["locales"] = list(scheme.locales or ["*"])
        for idx in range(1, 10):
            d[f"level_{idx}"] = {"format": scheme.levels.get(idx, "")}
        return d

    # ── Cancel ─────────────────────────────────────────────────────────

    def reject(self) -> None:
        if self._dirty and not self._confirm(
            t("editors.common.confirm_close"),
            t("editors.common.unsaved_close_message"),
        ):
            return
        super().reject()

    # ── Helpers ──────────────────────────────────────────────────────────

    @staticmethod
    def _prepare_combo(combo: QComboBox) -> None:
        from PySide6.QtWidgets import QListView

        combo.setMaxVisibleItems(20)
        view = QListView()
        view.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        combo.setView(view)

    @staticmethod
    def _combo_add_item(combo: QComboBox, display: str, data: Any) -> None:
        combo.addItem(display, data)
        combo.setItemData(combo.count() - 1, display, Qt.ItemDataRole.ToolTipRole)

    @staticmethod
    def _combo_current_data(combo: QComboBox) -> Any:
        return combo.currentData()

    @staticmethod
    def _combo_set_data(combo: QComboBox, data: Any) -> bool:
        for i in range(combo.count()):
            if combo.itemData(i) == data:
                combo.setCurrentIndex(i)
                return True
        return False

    def _confirm(self, title: str, message: str) -> bool:
        from PySide6.QtWidgets import QMessageBox

        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(message)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        box.setDefaultButton(QMessageBox.StandardButton.No)
        return box.exec() == QMessageBox.StandardButton.Yes

    def _notify_info(self, message: str, *, title: str = "") -> None:
        from PySide6.QtWidgets import QMessageBox

        box = QMessageBox(self)
        box.setWindowTitle(title or t("editors.common.prompt"))
        box.setText(message)
        box.setIcon(QMessageBox.Icon.Information)
        box.setStandardButtons(QMessageBox.StandardButton.Ok)
        box.exec()

    def _notify_warning(self, message: str) -> None:
        from PySide6.QtWidgets import QMessageBox

        box = QMessageBox(self)
        box.setWindowTitle(t("editors.common.cannot_delete"))
        box.setText(message)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setStandardButtons(QMessageBox.StandardButton.Ok)
        box.exec()

    def _notify_error(self, title: str, message: str) -> None:
        from PySide6.QtWidgets import QMessageBox

        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(message)
        box.setIcon(QMessageBox.Icon.Critical)
        box.setStandardButtons(QMessageBox.StandardButton.Ok)
        box.exec()

    # ── Public API ───────────────────────────────────────────────────────

    def get_schemes_data(self) -> dict[str, Any]:
        """Return the current schemes as a dict ready for TOML serialization."""
        return self._build_toml_dict()
