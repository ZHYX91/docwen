"""Focused tests split from test_settings_dialog_shell.py."""

from __future__ import annotations

from ._settings_dialog_shell_support import (
    pytest,
)

pytestmark = pytest.mark.gui


def test_settings_dialog_button_semantic_tiers(qapp) -> None:
    from PySide6.QtWidgets import QPushButton

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    dialog = SettingsDialog(view_model=SettingsViewModel())

    buttons = {
        "settingsResetTabButton": "secondary",
        "settingsResetAllButton": "secondary",
        "settingsOkButton": "primary",
        "settingsCancelButton": "secondary",
        "settingsApplyButton": "secondary",
    }
    for object_name, expected_class in buttons.items():
        button = dialog.findChild(QPushButton, object_name)
        assert button is not None
        assert button.property("class") == expected_class

    dialog.close()


def test_settings_dialog_title_uses_i18n(qapp) -> None:
    from docwen_gui.i18n import t
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    dialog = SettingsDialog(view_model=SettingsViewModel())

    assert dialog.windowTitle() == t("settings.title")
    assert dialog.windowTitle() != "Settings - DocWen Offline"

    dialog.close()


def test_settings_dialog_sidebar_icons_use_current_resource_helper(qapp) -> None:
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    for tab_key in TAB_KEYS:
        icon = SettingsDialog._load_tab_icon(tab_key)  # pyright: ignore[reportPrivateUsage]
        assert icon is not None, f"{tab_key} sidebar icon should load from assets/icons"
        assert not icon.isNull()

    assert SettingsDialog._load_tab_icon("__missing__") is None  # pyright: ignore[reportPrivateUsage]


def test_settings_dialog_sidebar_icon_missing_file_gracefully_degrades(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    monkeypatch.setattr("docwen_gui.resources.load_svg_icon", lambda _name: None)

    assert SettingsDialog._load_tab_icon("general") is None  # pyright: ignore[reportPrivateUsage]


def test_settings_dialog_focus_target_skips_containers_and_uses_first_control(qapp) -> None:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QComboBox, QScrollArea, QWidget

    from docwen_gui.widgets.settings.dialog import SettingsDialog

    tab = QWidget()
    tab.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
    scroll = QScrollArea(tab)
    scroll.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
    combo = QComboBox(tab)
    combo.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

    assert SettingsDialog._find_focus_target(tab) is combo  # pyright: ignore[reportPrivateUsage]

    combo.setEnabled(False)
    assert SettingsDialog._find_focus_target(tab) is tab  # pyright: ignore[reportPrivateUsage]
    tab.close()


def test_settings_dialog_action_buttons_keep_readable_geometry(qapp) -> None:
    from PySide6.QtWidgets import QPushButton

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import ACTION_BUTTON_MIN_HEIGHT, SettingsDialog

    dialog = SettingsDialog(view_model=SettingsViewModel())
    try:
        dialog.show()
        qapp.processEvents()

        bottom_margin = 8
        for object_name in (
            "settingsResetTabButton",
            "settingsResetAllButton",
            "settingsOkButton",
            "settingsCancelButton",
            "settingsApplyButton",
        ):
            button = dialog.findChild(QPushButton, object_name)
            assert button is not None
            assert button.height() >= ACTION_BUTTON_MIN_HEIGHT

            top_left = button.mapTo(dialog, button.rect().topLeft())
            bottom_right = button.mapTo(dialog, button.rect().bottomRight())
            assert dialog.rect().contains(top_left)
            assert dialog.rect().contains(bottom_right)
            assert dialog.height() - bottom_right.y() >= bottom_margin
    finally:
        dialog.close()


def test_settings_dialog_tab_order_uses_pyside6_order_and_tk_union(qapp) -> None:
    from PySide6.QtWidgets import QTabWidget

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    dialog = SettingsDialog(view_model=SettingsViewModel())
    tab_widget = dialog.findChild(QTabWidget)

    assert tab_widget is not None
    assert tab_widget.count() == len(TAB_KEYS)
    tk_settings_tabs = {
        "general",
        "text",
        "export",
        "document",
        "spreadsheet",
        "image",
        "layout",
        "other",
        "link",
        "formatting",
        "output",
        "logging",
    }
    pyside6_settings_order = [
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

    assert set(TAB_KEYS).issuperset(tk_settings_tabs)
    assert "proofread" in TAB_KEYS
    assert pyside6_settings_order == TAB_KEYS
    assert tab_widget.tabBar().isHidden() is True
    assert tab_widget.usesScrollButtons() is False

    dialog.close()


def test_settings_dialog_sidebar_items_are_reachable_and_click_through(qapp) -> None:
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog, _try_fluent_panel

    dialog = SettingsDialog(view_model=SettingsViewModel())
    try:
        assert dialog._navigation is not None  # pyright: ignore[reportPrivateUsage]
        items = {
            key: _try_fluent_panel(dialog._navigation, key)  # pyright: ignore[reportPrivateUsage]
            for key in TAB_KEYS
        }
        assert all(item is not None for item in items.values())

        proofread_item = items["proofread"]
        assert proofread_item is not None
        proofread_item.click()
        qapp.processEvents()

        assert dialog._tab_widget.currentIndex() == TAB_KEYS.index("proofread")  # pyright: ignore[reportPrivateUsage]
        assert dialog.current_section() == "proofread"
    finally:
        dialog.close()


def test_settings_dialog_isolates_one_page_construction_failure(
    qapp, monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QLabel, QTabWidget

    from docwen_gui.i18n import get_locale, set_locale
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module

    original_build_tab = dialog_module.SettingsDialog._build_tab  # pyright: ignore[reportPrivateUsage]

    def build_tab_or_fail(self, key: str):
        if key == "document":
            raise RuntimeError("document page exploded")
        return original_build_tab(self, key)

    monkeypatch.setattr(dialog_module.SettingsDialog, "_build_tab", build_tab_or_fail)

    previous_locale = get_locale()
    set_locale("zh_CN")
    with caplog.at_level("ERROR"):
        dialog = dialog_module.SettingsDialog(view_model=SettingsViewModel())
    try:
        tab_widget = dialog.findChild(QTabWidget)
        assert tab_widget is not None
        assert tab_widget.count() == len(dialog_module.TAB_KEYS)
        assert set(dialog._tabs) == set(dialog_module.TAB_KEYS)  # pyright: ignore[reportPrivateUsage]

        failed_page = dialog._tabs["document"]  # pyright: ignore[reportPrivateUsage]
        assert failed_page.objectName() == "settingsTabLoadErrorPage"
        assert failed_page.property("failedTabKey") == "document"
        title = failed_page.findChild(QLabel, "settingsTabTitle")
        message = failed_page.findChild(QLabel, "settingsTabLoadError")
        assert title is not None and title.text() == dialog_module.TAB_NAMES["document"]
        assert message is not None
        assert message.text() == (
            f"{dialog_module.TAB_NAMES['document']}加载失败\n\n错误：RuntimeError: document page exploded"
        )
        assert message.textFormat() == Qt.TextFormat.PlainText

        assert dialog._tabs["general"].objectName() != "settingsTabLoadErrorPage"  # pyright: ignore[reportPrivateUsage]
        assert dialog.activate_section("document") is False
        assert "Failed to build Settings tab: document" in caplog.text
    finally:
        dialog.close()
        set_locale(previous_locale)


def test_settings_dialog_activates_semantic_proofread_section(qapp) -> None:
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    dialog = SettingsDialog(view_model=SettingsViewModel())
    try:
        assert dialog.activate_section("proofread") is True
        assert TAB_KEYS[dialog._tab_widget.currentIndex()] == "proofread"  # pyright: ignore[reportPrivateUsage]
        assert dialog.activate_section("not-public") is False
    finally:
        dialog.close()


def test_settings_dialog_isolates_one_page_import_failure(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    from PySide6.QtWidgets import QTabWidget

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module

    original = dialog_module._TAB_SPECS["document"]  # pyright: ignore[reportPrivateUsage]

    def import_or_fail(_view_model: SettingsViewModel):
        raise ImportError("document dependency unavailable")

    monkeypatch.setitem(
        dialog_module._TAB_SPECS,  # pyright: ignore[reportPrivateUsage]
        "document",
        original._replace(factory=import_or_fail),
    )

    dialog = dialog_module.SettingsDialog(view_model=SettingsViewModel())
    try:
        tab_widget = dialog.findChild(QTabWidget)
        assert tab_widget is not None and tab_widget.count() == len(dialog_module.TAB_KEYS)
        assert dialog._tabs["document"].objectName() == "settingsTabLoadErrorPage"  # pyright: ignore[reportPrivateUsage]
        assert dialog._tabs["general"].objectName() != "settingsTabLoadErrorPage"  # pyright: ignore[reportPrivateUsage]
    finally:
        dialog.close()


def test_settings_dialog_nav_item_failure_exposes_scrollable_tab_fallback(
    qapp,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from PySide6.QtWidgets import QTabWidget

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    original_load_icon = SettingsDialog._load_tab_icon  # pyright: ignore[reportPrivateUsage]

    def load_icon_or_fail(key: str):
        if key == "document":
            raise RuntimeError("document nav item exploded")
        return original_load_icon(key)

    monkeypatch.setattr(SettingsDialog, "_load_tab_icon", staticmethod(load_icon_or_fail))

    dialog = SettingsDialog(view_model=SettingsViewModel())
    try:
        tab_widget = dialog.findChild(QTabWidget)
        assert tab_widget is not None
        assert tab_widget.tabBar().isHidden() is False
        assert tab_widget.usesScrollButtons() is True
        assert dialog._navigation is None  # pyright: ignore[reportPrivateUsage]

        document_index = TAB_KEYS.index("document")
        tab_widget.setCurrentIndex(document_index)
        assert tab_widget.currentWidget() is dialog._tabs["document"]  # pyright: ignore[reportPrivateUsage]
    finally:
        dialog.close()


def test_every_settings_page_has_a_visible_title_matching_navigation(qapp) -> None:
    from PySide6.QtWidgets import QLabel, QTabWidget

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, TAB_NAMES, SettingsDialog

    dialog = SettingsDialog(view_model=SettingsViewModel())
    try:
        tab_widget = dialog.findChild(QTabWidget)
        assert tab_widget is not None
        for index, key in enumerate(TAB_KEYS):
            page = dialog._tabs[key]  # pyright: ignore[reportPrivateUsage]
            title = page.findChild(QLabel, "settingsTabTitle")
            assert title is not None
            assert title.text() == TAB_NAMES[key]
            assert title.isHidden() is False
            assert tab_widget.tabText(index) == TAB_NAMES[key]
            assert tab_widget.tabToolTip(index) == TAB_NAMES[key]
    finally:
        dialog.close()


def test_settings_dialog_restores_visible_info_affordances(qapp) -> None:
    from PySide6.QtWidgets import QToolButton

    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    dialog = SettingsDialog(view_model=SettingsViewModel())
    info_buttons = dialog.findChildren(QToolButton, "settingsInfoButton")

    assert len(info_buttons) >= 12
    assert all(btn.toolTip() for btn in info_buttons)
    assert all((not btn.icon().isNull()) or btn.text() == "i" for btn in info_buttons)

    dialog.close()


def test_settings_dialog_changes_summary_uses_locale(qapp) -> None:
    from docwen_gui.i18n import t
    from docwen_gui.view_models.settings_vm import SECTION_GUI, SECTION_OUTPUT, SettingsViewModel
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    vm = SettingsViewModel()
    dialog = SettingsDialog(view_model=vm)

    vm.set_field_batch(
        SECTION_GUI,
        {
            "language": "en_US",
            "theme": "dark",
            "transparency_enabled": True,
            "transparency_value": 0.7,
            "remember_gui_state": False,
            "auto_center": True,
            "expand_side_panels": True,
            "default_mode": "batch",
        },
    )
    vm.set_field_batch(
        SECTION_OUTPUT,
        {
            "output_mode": "custom",
            "custom_path": "C:/out",
            "create_date_subfolder": True,
            "date_folder_format": "%Y%m%d",
            "auto_open_folder": True,
            "save_intermediate_files": True,
        },
    )

    changes = vm.get_change_summary()
    dialog._refresh_changes_summary()  # pyright: ignore[reportPrivateUsage]

    assert dialog._changes_label.text() == t("settings.changes.summary", count=len(changes))  # pyright: ignore[reportPrivateUsage]
    assert t("settings.changes.more", count=len(changes) - 10) in dialog._changes_label.toolTip()  # pyright: ignore[reportPrivateUsage]

    dialog.close()


def test_settings_dialog_current_tab_reset_uses_precise_registry_group(qapp, monkeypatch) -> None:
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import RESET_GROUPS, TAB_KEYS, SettingsDialog

    vm = SettingsViewModel()
    dialog = SettingsDialog(view_model=vm)
    reset_groups: list[str] = []

    def _reset_group(group: str) -> bool:
        reset_groups.append(group)
        return True

    def _reset_section(_section: str) -> bool:
        raise AssertionError("SettingsDialog should reset by registry group, not broad section")

    assert RESET_GROUPS["document"] == "document"
    assert RESET_GROUPS["spreadsheet"] == "spreadsheet"
    assert RESET_GROUPS["image"] == "image"
    assert RESET_GROUPS["layout"] == "layout"
    assert RESET_GROUPS["other"] == "other"
    assert "conversion_defaults" not in {
        RESET_GROUPS["document"],
        RESET_GROUPS["spreadsheet"],
        RESET_GROUPS["image"],
        RESET_GROUPS["layout"],
        RESET_GROUPS["other"],
    }

    monkeypatch.setattr(vm, "reset_group", _reset_group)
    monkeypatch.setattr(vm, "reset_section", _reset_section)
    dialog._tab_widget.setCurrentIndex(TAB_KEYS.index("document"))  # pyright: ignore[reportPrivateUsage]
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)

    dialog._on_reset_tab()  # pyright: ignore[reportPrivateUsage]

    assert reset_groups == ["document"]

    dialog.close()


def test_settings_dialog_reset_general_updates_visual_cancel_baseline(qapp, monkeypatch: pytest.MonkeyPatch) -> None:
    from PySide6.QtWidgets import QWidget

    from docwen_gui.models.settings_config import GUIConfig, SettingsConfig
    from docwen_gui.styles.theme_manager import ThemeManager
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import TAB_KEYS, SettingsDialog

    ThemeManager.reset_instance()
    manager = ThemeManager.get_instance()
    manager.initialize(qapp, "dark")

    parent = QWidget()
    parent.setWindowOpacity(0.55)
    initial = SettingsConfig(gui=GUIConfig(theme="dark", transparency_enabled=True, transparency_value=0.55))
    vm = SettingsViewModel(config=initial)
    dialog = SettingsDialog(parent=parent, view_model=vm)
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)

    def _reset_group_to_defaults(group: str) -> bool:
        assert group == "general"
        vm.load_full_config(SettingsConfig())
        vm.begin_session()
        return True

    monkeypatch.setattr(vm, "reset_group", _reset_group_to_defaults)
    rejected = False
    try:
        dialog._tab_widget.setCurrentIndex(TAB_KEYS.index("general"))  # pyright: ignore[reportPrivateUsage]

        dialog._on_reset_tab()  # pyright: ignore[reportPrivateUsage]

        assert manager.get_current_theme() == "light"
        assert abs(parent.windowOpacity() - 1.0) < 0.01

        manager.apply_theme("dark")
        parent.setWindowOpacity(0.55)
        dialog._on_cancel()  # pyright: ignore[reportPrivateUsage]
        rejected = True

        assert manager.get_current_theme() == "light"
        assert abs(parent.windowOpacity() - 1.0) < 0.01
    finally:
        if not rejected:
            dialog.close()
        parent.close()
        manager.apply_theme("light")


def test_settings_dialog_reset_all_requires_confirmation(qapp, monkeypatch) -> None:
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    dialog = SettingsDialog(view_model=SettingsViewModel())
    vm = dialog.view_model
    reset_calls = 0

    def _reset_all_spy() -> bool:
        nonlocal reset_calls
        reset_calls += 1
        return True

    monkeypatch.setattr(vm, "reset_all", _reset_all_spy)
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: False)

    dialog._on_reset_all()  # pyright: ignore[reportPrivateUsage]

    assert reset_calls == 0

    dialog.close()


def test_settings_dialog_reset_all_calls_view_model_after_confirmation(qapp, monkeypatch) -> None:
    from docwen_gui.view_models.settings_vm import SettingsViewModel
    from docwen_gui.widgets.settings import dialog as dialog_module
    from docwen_gui.widgets.settings.dialog import SettingsDialog

    dialog = SettingsDialog(view_model=SettingsViewModel())
    vm = dialog.view_model
    reset_calls = 0

    def _reset_all_spy() -> bool:
        nonlocal reset_calls
        reset_calls += 1
        return True

    monkeypatch.setattr(vm, "reset_all", _reset_all_spy)
    monkeypatch.setattr(dialog_module, "_show_confirm", lambda *_args, **_kwargs: True)

    dialog._on_reset_all()  # pyright: ignore[reportPrivateUsage]

    assert reset_calls == 1

    dialog.close()
