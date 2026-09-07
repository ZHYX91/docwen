"""Actual painted headers match their borders in both themes."""

import pytest
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QLabel

from docwen_gui.styles.global_aggregate import build_global_stylesheet
from docwen_gui.styles.theme_semantics import get_card_colors
from docwen_gui.widgets.panel_card import PanelCard

pytestmark = pytest.mark.gui


@pytest.mark.parametrize("theme", ["light", "dark"])
@pytest.mark.parametrize("tone", ["primary", "success", "warning", "danger"])
def test_card_header_and_border_paint_the_same_colour(qtbot, qapp, theme, tone):
    card = PanelCard("Result")
    qtbot.addWidget(card)
    card.content_layout.addWidget(QLabel("Document content", card))
    card.setTone(tone)
    card.setStyleSheet(build_global_stylesheet(theme))
    card.resize(360, 160)
    card.show()
    qapp.processEvents()
    surface, text = get_card_colors(tone, theme)
    pixels = card.grab().toImage()
    header_point = card.header.mapTo(card, card.header.rect().topLeft())
    assert pixels.pixelColor(0, card.height() // 2) == QColor(surface)
    assert pixels.pixelColor(header_point.x() + 8, header_point.y() + card.header.height() // 2) == QColor(surface)

    def luminance(colour):
        color = QColor(colour)
        rgb = (color.redF(), color.greenF(), color.blueF())
        linear = [value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4 for value in rgb]
        return sum(value * weight for value, weight in zip(linear, (0.2126, 0.7152, 0.0722), strict=True))

    light, dark = sorted((luminance(surface), luminance(text)), reverse=True)
    assert (light + 0.05) / (dark + 0.05) >= 4.5


@pytest.mark.parametrize("theme", ["light", "dark"])
def test_dialog_footer_buttons_receive_geometry_when_created_after_theme(qtbot, qapp, theme):
    from PySide6.QtWidgets import QDialogButtonBox, QWidget

    from docwen_gui.dialogs.activity_records import ActivityRecordsDialog
    from docwen_gui.styles.design_tokens import Sizing
    from docwen_gui.view_models.activity_records import ActivityRecordsModel
    from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
    from docwen_gui.view_models.task_history import TaskHistory
    from docwen_gui.widgets.settings.proofread_transfer import RuleImportDialog
    from docwen_runtime.config.proofread_transfer import plan_rule_import

    previous = qapp.styleSheet()
    feedback = InfoAreaViewModel()
    try:
        qapp.setStyleSheet(build_global_stylesheet(theme))
        parent = QWidget()
        qtbot.addWidget(parent)
        model = ActivityRecordsModel(TaskHistory(), feedback)
        plan = plan_rule_import("proofread/typos.toml", '[entries]\na = ["b"]\n', '[entries]\nc = ["d"]\n')
        dialogs = [ActivityRecordsDialog(model, parent), RuleImportDialog(plan, parent)]
        for dialog in dialogs:
            qtbot.addWidget(dialog)
            dialog.show()
            qapp.processEvents()
            box = dialog.findChild(QDialogButtonBox)
            assert box is not None
            for button in box.buttons():
                assert button.height() >= Sizing.CONTROL_HEIGHT
                assert button.width() >= Sizing.BUTTON_MIN_WIDTH
            dialog.close()
    finally:
        feedback.stop_all_timers()
        qapp.setStyleSheet(previous)
