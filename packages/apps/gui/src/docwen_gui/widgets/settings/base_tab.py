"""Base classes for Settings tab widgets.

``BaseSettingsTab`` provides common helpers (add section, add card, create
checkbox / combobox / spinbox) used by all 13 tabs.  All widgets bind to
``SettingsViewModel`` — they never touch config_manager or runtime directly.

``DynamicSettingsTab`` builds a form-driven tab from a schema dict.
"""

from __future__ import annotations

from typing import Any

from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QScrollArea,
    QSpinBox,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from ...resources import load_svg_icon

# Design token — matches GUI行为与交互规范.md §3.9
SPACING_XS = 4
SPACING_SM = 8
SPACING_MD = 12
SPACING_LG = 16
SPACING_XL = 24

CONTROL_HEIGHT = 32

SETTINGS_TOGGLE_OBJECT_NAME = "settingsToggle"
SETTINGS_INFO_BUTTON_OBJECT_NAME = "settingsInfoButton"


def _apply_control_height(widget: QWidget) -> None:
    """Set a standard control height on a widget."""
    widget.setMinimumHeight(CONTROL_HEIGHT)


def _prepare_combo(widget: QComboBox) -> None:
    """Configure a QComboBox for settings use."""
    from PySide6.QtWidgets import QListView

    widget.setMaxVisibleItems(20)
    # Use a list view for the popup to get middle-elide on Windows
    view = QListView()
    view.setTextElideMode(Qt.TextElideMode.ElideMiddle)
    widget.setView(view)


def _create_info_button(tooltip: str, parent: QWidget | None = None) -> QToolButton:
    """Create a small visible info affordance for settings help text."""
    btn = QToolButton(parent)
    btn.setObjectName(SETTINGS_INFO_BUTTON_OBJECT_NAME)
    btn.setAutoRaise(True)
    btn.setFixedSize(18, 18)
    btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    btn.setToolTip(tooltip)
    btn.setAccessibleName(tooltip)
    icon = load_svg_icon("info.svg")
    btn.setIconSize(QSize(14, 14))
    if icon is not None and not icon.isNull():
        btn.setIcon(icon)
    else:
        btn.setText("i")
    return btn


class BaseSettingsTab(QWidget):
    """Base class for settings tab pages.

    Provides helpers for building form-based UI and binding to the
    SettingsViewModel.  Subclasses override ``_create_interface()``.
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("settingsTabRoot")

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(SPACING_SM, 0, 0, 0)
        root_layout.setSpacing(SPACING_SM)

        self._tab_title = QLabel("", self)
        self._tab_title.setObjectName("settingsTabTitle")
        self._tab_title.setVisible(False)
        root_layout.addWidget(self._tab_title)

        # Tab-level description (hidden by default; subclasses call set_tab_description)
        self._tab_desc = QLabel("", self)
        self._tab_desc.setObjectName("settingsTabDescription")
        self._tab_desc.setProperty("settingsRole", "tabDescription")
        self._tab_desc.setWordWrap(True)
        self._tab_desc.setVisible(False)
        root_layout.addWidget(self._tab_desc)

        self._scroll_area = QScrollArea(self)
        self._scroll_area.setObjectName("settingsTabScrollArea")
        self._scroll_area.setWidgetResizable(True)
        self._scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        root_layout.addWidget(self._scroll_area, 1)

        self._scroll_container = QWidget(self._scroll_area)
        self._scroll_container.setObjectName("settingsTabScrollContainer")
        self._scroll_layout = QVBoxLayout(self._scroll_container)
        self._scroll_layout.setContentsMargins(SPACING_MD, SPACING_MD, SPACING_MD, SPACING_MD)
        self._scroll_layout.setSpacing(SPACING_MD)
        self._scroll_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self._scroll_area.setWidget(self._scroll_container)

        self._create_interface()

    def set_tab_description(self, text: str) -> None:
        """Set the tab-level description text shown above the scroll area."""
        self._tab_desc.setText(text)
        self._tab_desc.setVisible(bool(text.strip()))

    def set_tab_title(self, text: str) -> None:
        """Set the persistent page heading paired with sidebar navigation."""
        self._tab_title.setText(text)
        self._tab_title.setVisible(bool(text.strip()))

    # ── UI-building helpers ─────────────────────────────────────────────────

    def add_section(self, title: str) -> tuple[QGroupBox, QFormLayout]:
        """Add a titled section and return (group, form)."""
        group = QGroupBox(title, self._scroll_container)
        group.setObjectName("settingsSectionGroup")
        section_form = self._make_form(group)
        group.setLayout(section_form)
        self._scroll_layout.addWidget(group)
        return group, section_form

    def add_settings_card(
        self,
        title: str,
        description: str = "",
        *,
        object_name: str | None = None,
    ) -> tuple[QWidget, QFormLayout]:
        """Add a card-style settings block."""
        card = QWidget(self._scroll_container)
        card.setObjectName(object_name or "settingsSectionCard")
        card.setProperty("settingsRole", "settingsCard")

        layout = QVBoxLayout(card)
        layout.setContentsMargins(SPACING_LG, SPACING_LG, SPACING_LG, SPACING_LG)
        layout.setSpacing(10)

        title_label = QLabel(title, card)
        title_label.setObjectName("settingsCardTitle")
        layout.addWidget(title_label)

        desc = description.strip()
        if desc:
            desc_label = QLabel(desc, card)
            desc_label.setObjectName("settingsCardDescription")
            desc_label.setWordWrap(True)
            layout.addWidget(desc_label)

        form = self._make_form()
        layout.addLayout(form)

        self._scroll_layout.addWidget(card)
        return card, form

    @staticmethod
    def add_form_row(form: QFormLayout, label_text: str, widget: QWidget, tooltip: str | None = None) -> QWidget | None:
        """Add a label+widget row to a form layout."""
        effective_tooltip = tooltip or widget.toolTip()
        if effective_tooltip:
            widget.setToolTip(effective_tooltip)
        if not label_text.strip():
            form.addRow(widget)
            return None
        label_widget = BaseSettingsTab.create_label_with_info(label_text, effective_tooltip)
        if effective_tooltip:
            label_widget.setToolTip(effective_tooltip)
        form.addRow(label_widget, widget)
        return label_widget

    @staticmethod
    def add_form_description(form: QFormLayout, text: str) -> QLabel:
        """Add a full-width description row."""
        label = QLabel(text)
        label.setWordWrap(True)
        label.setProperty("class", "secondary")
        label.setProperty("settingsRole", "sectionDescription")
        form.addRow(label)
        return label

    @staticmethod
    def create_label_with_info(text: str, tooltip: str | None = None, parent: QWidget | None = None) -> QWidget:
        """Create a label row with optional info icon."""
        container = QWidget(parent)
        container.setObjectName("settingsLabelWithInfo")
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        label = QLabel(text, container)
        label.setWordWrap(True)
        label.setProperty("settingsRole", "fieldLabel")
        layout.addWidget(label)
        if tooltip:
            label.setToolTip(tooltip)
            layout.addWidget(_create_info_button(tooltip, container))
        layout.addStretch(1)
        return container

    # ── Widget factory helpers ──────────────────────────────────────────────

    def create_checkbox(
        self,
        text: str,
        tooltip: str | None = None,
        default: bool = False,
        *,
        object_name: str | None = None,
    ) -> QCheckBox:
        """Create a checkbox with standard styling hook."""
        try:
            from qfluentwidgets import CheckBox as FluentCheckBox

            cb = FluentCheckBox(text, self._scroll_container)
        except ImportError:
            cb = QCheckBox(text, self._scroll_container)
        if object_name:
            cb.setObjectName(object_name)
        if tooltip:
            cb.setToolTip(tooltip)
        cb.setChecked(default)
        return cb

    def create_settings_toggle(self, text: str, tooltip: str | None = None, default: bool = False) -> QCheckBox:
        """Create a toggle checkbox with the unified settings object name."""
        return self.create_checkbox(text, tooltip, default=default, object_name=SETTINGS_TOGGLE_OBJECT_NAME)

    def create_combobox(
        self, items: list[tuple[str, Any]], tooltip: str | None = None, default_data: Any = None
    ) -> QComboBox:
        """Create a combobox with items (label, data) tuples."""
        cb = QComboBox(self._scroll_container)
        _prepare_combo(cb)
        if tooltip:
            cb.setToolTip(tooltip)
        for label, data in items:
            cb.addItem(label, data)
            if tooltip:
                cb.setItemData(cb.count() - 1, tooltip, Qt.ItemDataRole.ToolTipRole)
        if default_data is not None:
            self.set_combo_data(cb, default_data)
        return cb

    def create_spinbox(self, minimum: int, maximum: int, tooltip: str | None = None, default: int = 0) -> QSpinBox:
        """Create a spinbox."""
        sb = QSpinBox(self._scroll_container)
        sb.setRange(minimum, maximum)
        if tooltip:
            sb.setToolTip(tooltip)
        sb.setValue(default)
        _apply_control_height(sb)
        return sb

    def create_double_spinbox(
        self,
        minimum: float,
        maximum: float,
        tooltip: str | None = None,
        default: float = 1.0,
        decimals: int = 2,
        single_step: float = 0.05,
    ) -> QDoubleSpinBox:
        """Create a double spinbox."""
        sb = QDoubleSpinBox(self._scroll_container)
        sb.setRange(minimum, maximum)
        sb.setDecimals(decimals)
        sb.setSingleStep(single_step)
        if tooltip:
            sb.setToolTip(tooltip)
        sb.setValue(default)
        _apply_control_height(sb)
        return sb

    # ── ComboBox helper methods ─────────────────────────────────────────────

    @staticmethod
    def set_combo_data(combo: QComboBox, data: Any) -> bool:
        """Select a combo item by its data value."""
        for i in range(combo.count()):
            if combo.itemData(i) == data:
                combo.setCurrentIndex(i)
                return True
        return False

    @staticmethod
    def get_combo_data(combo: QComboBox) -> Any:
        """Get the data value of the currently selected combo item."""
        return combo.currentData()

    # ── Internal helpers ────────────────────────────────────────────────────

    @staticmethod
    def _make_form(parent: QWidget | None = None) -> QFormLayout:
        form = QFormLayout(parent) if parent is not None else QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setVerticalSpacing(10)
        form.setHorizontalSpacing(SPACING_MD)
        form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapLongRows)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        return form

    def _create_interface(self) -> None:
        """Override in subclasses to build the tab UI."""
        raise NotImplementedError


# ── DynamicSettingsTab ──────────────────────────────────────────────────────


class DynamicSettingsTab(BaseSettingsTab):
    """Form-driven settings tab built from a schema dict.

    Usage::

        class MyTab(DynamicSettingsTab):
            def __init__(self, parent, ...):
                schema = [
                    {"title": "Section", "presentation": "card", "fields": [
                        {"key": "my_key", "type": "checkbox", "text": "Label"},
                    ]},
                ]
                super().__init__(parent, config_name="...", section_name="...",
                                 form_schema=schema)
    """

    def __init__(
        self,
        parent: QWidget | None,
        config_name: str,
        section_name: str,
        form_schema: list[dict[str, Any]],
    ) -> None:
        self._config_name = config_name
        self._section_name = section_name
        self._form_schema = form_schema
        self._widgets: dict[str, QWidget] = {}
        self._initial_values: dict[str, Any] = {}
        super().__init__(parent)
        self._load_initial_values()
        # Connect widget signals to ViewModel if a ViewModel is available
        # (set by subclass as self._vm before calling super().__init__)
        self._connect_widgets_to_vm()

    @property
    def config_name(self) -> str:
        return self._config_name

    @property
    def section_name(self) -> str:
        return self._section_name

    @property
    def widgets(self) -> dict[str, QWidget]:
        return dict(self._widgets)

    def _create_interface(self) -> None:
        for section in self._form_schema:
            title = str(section.get("title", "") or "")
            section_desc = str(section.get("description", "") or "").strip()
            presentation = str(section.get("presentation", "section") or "section").strip().lower()

            if presentation == "card":
                _card, form = self.add_settings_card(title, section_desc)
            else:
                _group, form = self.add_section(title)
                if section_desc:
                    self.add_form_description(form, section_desc)

            for field in section.get("fields", []):
                key = field["key"]
                widget_type = field["type"]
                label_text = field.get("label", "")
                tooltip = field.get("tooltip", "")

                if widget_type == "description":
                    self.add_form_description(form, str(field.get("text", "") or ""))
                    continue

                widget: QWidget | None = None
                if widget_type == "checkbox":
                    widget = self.create_checkbox(field.get("text", ""), tooltip)
                elif widget_type == "combobox":
                    widget = self.create_combobox(field.get("items", []), tooltip)
                elif widget_type == "spinbox":
                    widget = self.create_spinbox(field.get("min", 0), field.get("max", 100), tooltip)
                elif widget_type == "text":
                    text_widget = QLineEdit(self._scroll_container)
                    if tooltip:
                        text_widget.setToolTip(tooltip)
                    widget = text_widget

                if widget is not None:
                    self._widgets[key] = widget
                    self.add_form_row(form, label_text, widget, tooltip)

    def _connect_widgets_to_vm(self) -> None:
        """Connect widget signals to the ViewModel for dirty tracking.

        If the subclass set ``self._vm`` before calling super().__init__(),
        widget changes will be forwarded to the ViewModel.
        """
        vm = getattr(self, "_vm", None)
        if vm is None or not hasattr(vm, "set_field"):
            return
        # Map section_name to ViewModel section constant
        from ...view_models.settings_vm import SECTION_CONVERSION_DEFAULTS

        section = SECTION_CONVERSION_DEFAULTS
        category = self._section_name
        set_conversion_default = getattr(vm, "set_conversion_default", None)

        for key, widget in self._widgets.items():
            if callable(set_conversion_default) and section == SECTION_CONVERSION_DEFAULTS:
                if isinstance(widget, QCheckBox):
                    widget.toggled.connect(
                        lambda checked, k=key, cat=category: vm.set_conversion_default(cat, k, checked)
                    )
                elif isinstance(widget, QComboBox):
                    widget.currentIndexChanged.connect(
                        lambda _idx, k=key, w=widget, cat=category: vm.set_conversion_default(cat, k, w.currentData())
                    )
                elif isinstance(widget, QSpinBox):
                    widget.valueChanged.connect(lambda val, k=key, cat=category: vm.set_conversion_default(cat, k, val))
                elif isinstance(widget, QLineEdit):
                    widget.textChanged.connect(lambda txt, k=key, cat=category: vm.set_conversion_default(cat, k, txt))
            elif isinstance(widget, QCheckBox):
                widget.toggled.connect(lambda checked, k=key: vm.set_field(section, k, checked))
            elif isinstance(widget, QComboBox):
                widget.currentIndexChanged.connect(
                    lambda _idx, k=key, w=widget: vm.set_field(section, k, w.currentData())
                )
            elif isinstance(widget, QSpinBox):
                widget.valueChanged.connect(lambda val, k=key: vm.set_field(section, k, val))
            elif isinstance(widget, QLineEdit):
                widget.textChanged.connect(lambda txt, k=key: vm.set_field(section, k, txt))

    def load_values_from_dict(self, data: dict[str, Any]) -> None:
        """Load values from a data dict into widgets (signals blocked)."""
        # Block all widget signals during bulk load
        for widget in self._widgets.values():
            widget.blockSignals(True)
        try:
            for key, widget in self._widgets.items():
                if key not in data:
                    continue
                val = data[key]
                if isinstance(widget, QCheckBox):
                    widget.setChecked(bool(val))
                elif isinstance(widget, QComboBox):
                    self.set_combo_data(widget, val)
                elif isinstance(widget, QSpinBox):
                    widget.setValue(int(val))
                elif isinstance(widget, QLineEdit):
                    widget.setText(str(val))
        finally:
            for widget in self._widgets.values():
                widget.blockSignals(False)

    def collect_values(self) -> dict[str, Any]:
        """Collect current widget values into a dict."""
        result: dict[str, Any] = {}
        for key, widget in self._widgets.items():
            if isinstance(widget, QCheckBox):
                result[key] = widget.isChecked()
            elif isinstance(widget, QComboBox):
                result[key] = self.get_combo_data(widget)
            elif isinstance(widget, QSpinBox):
                result[key] = widget.value()
            elif isinstance(widget, QLineEdit):
                result[key] = widget.text()
        return result

    def _load_initial_values(self) -> None:
        self._initial_values = {key: self._extract_value(widget) for key, widget in self._widgets.items()}

    def get_changed_settings(self) -> dict[str, dict[str, Any]]:
        """Return only the changed values (old/new pairs)."""
        changes: dict[str, dict[str, Any]] = {}
        for key, widget in self._widgets.items():
            current = self._extract_value(widget)
            initial = self._initial_values.get(key)
            if current != initial:
                changes[key] = {"old": initial, "new": current}
        return changes

    @staticmethod
    def _extract_value(widget: QWidget) -> Any:
        if isinstance(widget, QCheckBox):
            return widget.isChecked()
        if isinstance(widget, QComboBox):
            return widget.currentData()
        if isinstance(widget, QSpinBox):
            return widget.value()
        if isinstance(widget, QLineEdit):
            return widget.text()
        return None
