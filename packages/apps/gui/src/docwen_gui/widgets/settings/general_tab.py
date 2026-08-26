"""General settings tab — language, theme, transparency, window behavior.

Matches old GeneralTab behavior:
- Language (QComboBox with available locales)
- Theme (light/dark/system with live preview)
- Transparency toggle + value (0.20-1.00), linked state
- Remember window state, auto-center, expand side panels
- Default mode (single/batch)
"""

from __future__ import annotations

from typing import cast as _cast

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QHBoxLayout,
    QLabel,
    QVBoxLayout,
    QWidget,
)

from ...i18n import t
from ...view_models.settings_vm import SECTION_GUI, SettingsViewModel
from .base_tab import BaseSettingsTab, _apply_control_height, _prepare_combo

GENERAL_THEME_VALUES = ["light", "dark", "system"]
GENERAL_MODE_VALUES = ["single", "batch"]


class GeneralTab(BaseSettingsTab):
    """General/appearance settings tab."""

    def __init__(self, view_model: SettingsViewModel) -> None:
        self._vm = view_model
        self._language_combo: QComboBox = _cast(QComboBox, None)
        self._theme_combo: QComboBox = _cast(QComboBox, None)
        self._transparency_enabled: QCheckBox = _cast(QCheckBox, None)
        self._transparency_value: QDoubleSpinBox = _cast(QDoubleSpinBox, None)
        self._transparency_value_label: QLabel = _cast(QLabel, None)
        self._remember_state: QCheckBox = _cast(QCheckBox, None)
        self._auto_center: QCheckBox = _cast(QCheckBox, None)
        self._expand_side_panels: QCheckBox = _cast(QCheckBox, None)
        self._default_mode: QComboBox = _cast(QComboBox, None)
        super().__init__()
        self.set_tab_description(t("settings.descriptions.general", "Language, theme, and window behavior"))
        self._load_values()

    def _create_interface(self) -> None:
        # ── Language card ───────────────────────────────────────────────
        lang_card, lang_form = self.add_settings_card(
            t("settings.general.language_section", "Language"),
            t(
                "settings.general.language_tooltip",
                "Select the display language. Requires restart to take full effect.",
            ),
            object_name="generalLanguageCard",
        )
        lang_container = QWidget(lang_card)
        lang_container_layout = QVBoxLayout(lang_container)
        lang_container_layout.setContentsMargins(0, 0, 0, 0)
        lang_container_layout.setSpacing(8)

        lang_combo = QComboBox(lang_container)
        self._language_combo = lang_combo
        _prepare_combo(lang_combo)
        lang_combo.addItem(t("settings.general.languages.zh_CN", "Chinese (Simplified)"), "zh_CN")
        lang_combo.addItem(t("settings.general.languages.en_US", "English"), "en_US")
        lang_combo.addItem(t("settings.general.languages.de_DE", "German"), "de_DE")
        lang_combo.addItem(t("settings.general.languages.es_ES", "Spanish"), "es_ES")
        lang_combo.addItem(t("settings.general.languages.fr_FR", "French"), "fr_FR")
        lang_combo.addItem(t("settings.general.languages.ja_JP", "Japanese"), "ja_JP")
        lang_combo.addItem(t("settings.general.languages.ko_KR", "Korean"), "ko_KR")
        lang_combo.addItem(t("settings.general.languages.pt_BR", "Portuguese (Brazil)"), "pt_BR")
        lang_combo.addItem(t("settings.general.languages.ru_RU", "Russian"), "ru_RU")
        lang_combo.addItem(t("settings.general.languages.vi_VN", "Vietnamese"), "vi_VN")
        lang_combo.addItem(t("settings.general.languages.zh_TW", "Chinese (Traditional)"), "zh_TW")
        lang_container_layout.addWidget(lang_combo)

        hint = QLabel(
            t("settings.general.language_restart_hint", "Requires application restart to take full effect."),
            lang_container,
        )
        hint.setObjectName("generalLanguageHint")
        hint.setWordWrap(True)
        hint.setAlignment(Qt.AlignmentFlag.AlignLeft)
        lang_container_layout.addWidget(hint)

        self.add_form_row(
            lang_form,
            t("settings.general.language_label", "Display Language:"),
            lang_container,
            t("settings.general.language_tooltip", "Select the display language"),
        )
        lang_combo.currentIndexChanged.connect(self._on_language_changed)

        # ── Theme card ──────────────────────────────────────────────────
        theme_card, theme_form = self.add_settings_card(
            t("settings.general.theme_section", "Theme"),
            t("settings.general.theme_description", "Choose between light, dark, or follow system theme."),
            object_name="generalThemeCard",
        )
        theme_combo = QComboBox(theme_card)
        self._theme_combo = theme_combo
        _prepare_combo(theme_combo)
        theme_combo.addItem(t("settings.general.themes.light", "Light"), "light")
        theme_combo.addItem(t("settings.general.themes.dark", "Dark"), "dark")
        theme_combo.addItem(t("settings.general.themes.system", "Follow System"), "system")
        self.add_form_row(
            theme_form,
            t("settings.general.theme_label", "App Theme:"),
            theme_combo,
            t("settings.general.theme_tooltip", "Choose theme mode"),
        )
        theme_combo.currentIndexChanged.connect(self._on_theme_changed)

        # Theme preview
        preview_container = QWidget(theme_card)
        preview_layout = QVBoxLayout(preview_container)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(8)
        preview_title = QLabel(t("settings.general.theme_preview", "Preview:"), preview_container)
        preview_title.setObjectName("generalThemePreviewTitle")
        preview_title.setAlignment(Qt.AlignmentFlag.AlignLeft)
        preview_layout.addWidget(preview_title)

        preview_frame = QWidget(preview_container)
        preview_frame.setObjectName("generalThemePreviewFrame")
        preview_frame_layout = QHBoxLayout(preview_frame)
        preview_frame_layout.setContentsMargins(12, 12, 12, 12)
        preview_frame_layout.setSpacing(12)
        sample_button = QLabel(f"[ {t('settings.general.sample_button', 'Sample Button')} ]", preview_frame)
        sample_button.setObjectName("generalThemePreviewButton")
        sample_text = QLabel(
            t("settings.general.sample_text", "This is a sample text for theme preview."), preview_frame
        )
        sample_text.setObjectName("generalThemePreviewText")
        preview_frame_layout.addWidget(sample_button)
        preview_frame_layout.addWidget(sample_text, 1)
        preview_layout.addWidget(preview_frame)
        theme_form.addWidget(preview_container)

        # ── Transparency card ───────────────────────────────────────────
        transp_card, transp_form = self.add_settings_card(
            t("settings.general.transparency.section", "Window Transparency"),
            t("settings.general.transparency.tooltip", "Enable window transparency and adjust the opacity level."),
            object_name="generalTransparencyCard",
        )
        transp_container = QWidget(transp_card)
        transp_container_layout = QVBoxLayout(transp_container)
        transp_container_layout.setContentsMargins(0, 0, 0, 0)
        transp_container_layout.setSpacing(8)

        enabled = self.create_settings_toggle(
            t("settings.general.transparency.enable", "Enable Window Transparency"),
            t("settings.general.transparency.tooltip", "Enable window transparency effect"),
        )
        enabled.setParent(transp_container)
        self._transparency_enabled = enabled
        transp_container_layout.addWidget(enabled)

        value_row = QWidget(transp_container)
        value_row.setObjectName("generalTransparencyValueRow")
        value_row_layout = QHBoxLayout(value_row)
        value_row_layout.setContentsMargins(0, 0, 0, 0)
        value_row_layout.setSpacing(8)

        value = QDoubleSpinBox(value_row)
        value.setObjectName("generalTransparencySpinBox")
        value.setRange(0.20, 1.00)
        value.setSingleStep(0.05)
        value.setDecimals(2)
        value.setMaximumWidth(92)
        value.setToolTip(
            t(
                "settings.general.transparency.opacity_value_tooltip",
                "Adjust window opacity (0.20 = most transparent, 1.00 = fully opaque)",
            )
        )
        _apply_control_height(value)
        self._transparency_value = value
        value_row_layout.addWidget(value)

        value_label = QLabel(value_row)
        value_label.setObjectName("generalTransparencyPercentLabel")
        value_label.setMinimumWidth(40)
        self._transparency_value_label = value_label
        value_row_layout.addWidget(value_label)
        value_row_layout.addStretch(1)
        transp_container_layout.addWidget(value_row)
        transp_form.addWidget(transp_container)

        enabled.toggled.connect(self._on_transparency_toggled)
        value.valueChanged.connect(self._on_transparency_value_changed)

        # ── Window card ─────────────────────────────────────────────────
        window_card, window_form = self.add_settings_card(
            t("settings.general.window.section", "Window Behavior"),
            t("settings.general.window.description", "Configure window state and default mode."),
            object_name="generalWindowCard",
        )

        remember = self.create_settings_toggle(
            t("settings.general.window.remember_state", "Remember window size and position"),
            t("settings.general.window.remember_state_tooltip", "Save and restore window geometry between sessions"),
        )
        remember.setParent(window_card)
        self._remember_state = remember
        window_form.addWidget(remember)
        remember.toggled.connect(self._on_remember_state_toggled)

        auto_center = self.create_settings_toggle(
            t("settings.general.window.auto_center", "Auto-center window on startup"),
            t("settings.general.window.auto_center_tooltip", "Center the window on screen when launching"),
        )
        auto_center.setParent(window_card)
        self._auto_center = auto_center
        window_form.addWidget(auto_center)
        auto_center.toggled.connect(self._on_auto_center_toggled)

        expand = self.create_settings_toggle(
            t("settings.general.window.expand_side_panels", "Expand side panels with window"),
            t(
                "settings.general.window.expand_side_panels_tooltip",
                "Allow visible side panels to expand when the window grows",
            ),
        )
        expand.setParent(window_card)
        self._expand_side_panels = expand
        window_form.addWidget(expand)
        expand.toggled.connect(self._on_expand_side_panels_toggled)

        default_mode = QComboBox(window_card)
        self._default_mode = default_mode
        _prepare_combo(default_mode)
        default_mode.addItem(t("settings.general.window.single_mode", "Single File Mode"), "single")
        default_mode.addItem(t("settings.general.window.batch_mode", "Batch Mode"), "batch")
        self.add_form_row(
            window_form,
            t("settings.general.window.default_mode_label", "Default Mode:"),
            default_mode,
            t(
                "settings.general.window.default_mode_tooltip",
                "Sets the default input mode when the application starts",
            ),
        )
        default_mode.currentIndexChanged.connect(self._on_default_mode_changed)

    # ── Value loading ───────────────────────────────────────────────────────

    def _load_values(self) -> None:
        """Load values from ViewModel into widgets."""
        config = self._vm.config
        gui = config.gui

        # Block signals to prevent spurious dirty-state emissions during load
        signal_widgets = [
            self._language_combo,
            self._theme_combo,
            self._transparency_enabled,
            self._transparency_value,
            self._remember_state,
            self._auto_center,
            self._expand_side_panels,
            self._default_mode,
        ]
        for w in signal_widgets:
            if w is not None:
                w.blockSignals(True)

        try:
            if self._language_combo is not None:
                self.set_combo_data(self._language_combo, gui.language)
            if self._theme_combo is not None:
                self.set_combo_data(self._theme_combo, gui.theme)
            if self._transparency_enabled is not None:
                self._transparency_enabled.setChecked(gui.transparency_enabled)
            if self._transparency_value is not None:
                val = max(0.20, min(1.00, gui.transparency_value))
                self._transparency_value.setValue(val)
                self._transparency_value.setEnabled(gui.transparency_enabled)
                self._update_transparency_label()
            if self._remember_state is not None:
                self._remember_state.setChecked(gui.remember_gui_state)
            if self._auto_center is not None:
                self._auto_center.setChecked(gui.auto_center)
            if self._expand_side_panels is not None:
                self._expand_side_panels.setChecked(gui.expand_side_panels)
            if self._default_mode is not None:
                self.set_combo_data(self._default_mode, gui.default_mode)
        finally:
            for w in signal_widgets:
                if w is not None:
                    w.blockSignals(False)

    def reload_from_config(self) -> None:
        """Reload UI from ViewModel (called after reset)."""
        self._load_values()

    # ── Signal handlers ─────────────────────────────────────────────────────

    def _on_language_changed(self, _index: int) -> None:
        locale = self.get_combo_data(self._language_combo)
        if locale:
            self._vm.set_field(SECTION_GUI, "language", locale)

    def _on_theme_changed(self, _index: int) -> None:
        theme = self.get_combo_data(self._theme_combo)
        if theme:
            self._vm.set_field(SECTION_GUI, "theme", theme)
            # Apply theme immediately for live preview
            from docwen_gui.styles.theme_manager import ThemeManager

            ThemeManager.get_instance().apply_theme(theme)

    def _on_transparency_toggled(self, checked: bool) -> None:
        if self._transparency_value is not None:
            self._transparency_value.setEnabled(bool(checked))
        self._update_transparency_label()
        self._vm.set_field(SECTION_GUI, "transparency_enabled", bool(checked))
        self._apply_opacity_preview()

    def _on_transparency_value_changed(self, _value: float) -> None:
        self._update_transparency_label()
        if self._transparency_enabled is not None and not self._transparency_enabled.isChecked():
            return
        val = float(self._transparency_value.value()) if self._transparency_value else 1.0
        self._vm.set_field(SECTION_GUI, "transparency_value", val)
        self._apply_opacity_preview()

    def _on_remember_state_toggled(self, checked: bool) -> None:
        self._vm.set_field(SECTION_GUI, "remember_gui_state", bool(checked))

    def _on_auto_center_toggled(self, checked: bool) -> None:
        self._vm.set_field(SECTION_GUI, "auto_center", bool(checked))

    def _on_expand_side_panels_toggled(self, checked: bool) -> None:
        self._vm.set_field(SECTION_GUI, "expand_side_panels", bool(checked))

    def _on_default_mode_changed(self, _index: int) -> None:
        mode = self.get_combo_data(self._default_mode)
        if mode:
            self._vm.set_field(SECTION_GUI, "default_mode", mode)

    def _update_transparency_label(self) -> None:
        if self._transparency_value_label is None or self._transparency_value is None:
            return
        pct = round(float(self._transparency_value.value()) * 100)
        self._transparency_value_label.setText(f"{pct}%")
        self._transparency_value_label.setEnabled(self._transparency_value.isEnabled())

    def _apply_opacity_preview(self) -> None:
        """Preview opacity on the dialog's parent window without persisting."""
        enabled = bool(self._transparency_enabled.isChecked()) if self._transparency_enabled is not None else False
        value = float(self._transparency_value.value()) if self._transparency_value is not None else 1.0
        target_opacity = value if enabled else 1.0
        self._vm.preview_opacity(target_opacity)

        dialog = self.window()
        main_window = dialog.parentWidget() if dialog is not None else None
        if main_window is None:
            return
        setter = getattr(main_window, "setWindowOpacity", None)
        if callable(setter):
            setter(target_opacity)
