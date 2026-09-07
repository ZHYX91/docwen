"""User-visible execution scope stays aligned with actual dispatch and settings."""

from pathlib import Path

import pytest

from docwen_gui.i18n import t

from ._main_window_projection_binding_support import _load_request_templates
from ._main_window_projection_binding_support import window as window

pytestmark = pytest.mark.gui


def test_batch_label_counts_current_category_and_matches_dispatch(window, tmp_path, monkeypatch):
    first, second = tmp_path / "first.md", tmp_path / "second.md"
    table = tmp_path / "table.csv"
    for path in (first, second):
        path.write_text("# Example\n", encoding="utf-8")
    table.write_text("Name,Value\nExample,1\n", encoding="utf-8")
    _load_request_templates(window)
    window._input_area_vm.set_mode("batch")
    window._input_area_vm.add_files([str(first), str(second), str(table)])
    window._batch_list.select_file(str(first))
    calls = []
    monkeypatch.setattr(window, "_start_batch_execution", lambda **kwargs: calls.append(kwargs))
    button = window._action_area.convert_docx_button
    assert button.text() == t("common.action_file_count", action=f"{t('action_area.generate')} DOCX", count=2)
    button.click()
    assert len(calls) == 1
    assert {Path(path) for path in calls[0]["file_paths"]} == {first, second}
    window._batch_list.select_file(str(table))
    conversion = window._conversion_panel._conversion_button
    assert conversion.text() == t("common.action_file_count", action=conversion.property("baseActionLabel"), count=1)


def test_output_preview_survives_resize_and_reflects_committed_policy(window, tmp_path, qapp):
    source = tmp_path / "source.md"
    source.write_text("# Example\n", encoding="utf-8")
    window._input_area_vm.add_files([str(source)])
    label = window._output_location_label
    assert str(tmp_path) in label.full_text
    destination = tmp_path / "A long output folder name" / "Another long folder name"
    config = window._view_model.controller.config_port
    config._values.update({"output.directory.mode": "custom", "output.directory.custom_path": str(destination)})
    window._apply_runtime_window_settings()
    assert str(destination) in label.full_text
    assert window._build_output_policy().output_dir in label.full_text
    window.show()
    label.resize(120, label.sizeHint().height())
    qapp.processEvents()
    assert label.text()
    assert str(destination) in label.toolTip()
    window._input_area.clear_button.click()
    assert label.isHidden()


@pytest.mark.parametrize("theme", ["light", "dark"])
def test_settings_actions_keep_dimensions_after_theme_change(qtbot, qapp, theme):
    from PySide6.QtWidgets import QDialog, QPushButton, QVBoxLayout

    from docwen_gui.styles.design_tokens import Sizing
    from docwen_gui.styles.settings import build_settings_stylesheet

    dialog = QDialog()
    dialog.setObjectName("settingsDialog")
    qtbot.addWidget(dialog)
    layout = QVBoxLayout(dialog)
    buttons = []
    for name in ("settingsOkButton", "settingsCancelButton", "settingsApplyButton"):
        button = QPushButton("OK", dialog)
        button.setObjectName(name)
        layout.addWidget(button)
        buttons.append(button)
    dialog.setStyleSheet(build_settings_stylesheet("light"))
    dialog.show()
    dialog.setStyleSheet(build_settings_stylesheet(theme))
    qapp.processEvents()
    assert len({button.height() for button in buttons}) == 1
    for button in buttons:
        assert button.width() >= Sizing.BUTTON_MIN_WIDTH
        assert button.height() >= Sizing.CONTROL_HEIGHT


def test_execution_caption_wraps_without_losing_keyboard_activation(qtbot):
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont

    from docwen_gui.widgets.action_button import ActionButton

    button = ActionButton("Export Markdown (12 files)")
    qtbot.addWidget(button)
    button.setFont(QFont("Segoe UI", 18))
    button.resize(160, 40)
    button.show()
    qtbot.waitUntil(lambda: button.height() >= button.heightForWidth(button.width()))
    assert button.height() > 40
    assert button.text() == "Export Markdown (12 files)"
    button.setFocus()
    with qtbot.waitSignal(button.clicked):
        qtbot.keyClick(button, Qt.Key.Key_Space)
    button.resize(500, button.height())
    qtbot.waitUntil(lambda: button.minimumHeight() == 40)
