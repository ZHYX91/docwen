"""全局样式表聚合入口。"""

from __future__ import annotations

from .about_dialog import build_about_dialog_stylesheet
from .action_area import build_action_area_stylesheet
from .batch_list import build_batch_list_stylesheet
from .conversion_panel import build_conversion_panel_stylesheet
from .design_tokens import Typography
from .disabled_button import build_disabled_button_stylesheet
from .info_area import build_info_area_stylesheet
from .main_window import build_main_window_stylesheet
from .panel import build_panel_stylesheet
from .settings import build_settings_stylesheet
from .template_selector import build_template_selector_stylesheet


def build_global_stylesheet(theme_name: str, font_size_preset: str | None = None) -> str:
    """组装应用级全局样式表。

    Args:
        theme_name: 主题名（如 ``"light"`` / ``"dark_blue"``）。
    """
    return "\n\n".join(
        [
            "\n".join(
                [
                    "/* docwen-application-typography */",
                    "QWidget {",
                    f"    font-size: {Typography.qss(Typography.BODY_SIZE, font_size_preset)};",
                    "}",
                ]
            ),
            build_disabled_button_stylesheet(theme_name),
            build_panel_stylesheet(theme_name, font_size_preset),
            build_settings_stylesheet(theme_name, font_size_preset),
            build_action_area_stylesheet(font_size_preset),
            build_info_area_stylesheet(theme_name, font_size_preset),
            build_batch_list_stylesheet(theme_name, font_size_preset),
            build_main_window_stylesheet(),
            build_conversion_panel_stylesheet(theme_name, font_size_preset),
            build_about_dialog_stylesheet(font_size_preset),
            build_template_selector_stylesheet(),
        ]
    )
