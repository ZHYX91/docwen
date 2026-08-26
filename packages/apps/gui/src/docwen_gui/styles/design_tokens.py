"""
GUI 视觉 Token 定义。

提供统一的尺寸、间距、字号等常量，避免在各组件中硬编码。
"""


class Spacing:
    """全局间距基准（使用前需通过 ``ui_scale`` 进行应用级 UI 缩放）。"""

    XS = 4
    SM = 8
    MD = 12
    LG = 16
    XL = 24


class Typography:
    """全局字体层级（具体字体由 font_utils 提供）。"""

    CAPTION_SIZE = 11
    BODY_SIZE = 12
    CARD_TITLE_SIZE = 13
    SECTION_TITLE_SIZE = 14
    EMPHASIS_TITLE_SIZE = 15
    HERO_SIZE = 16
    PAGE_TITLE_SIZE = 18
    DIALOG_TITLE_SIZE = 20

    @staticmethod
    def resolve(default_size: int, preset: str | None = None) -> int:
        from docwen_gui.font_utils import resolve_typography_size

        return resolve_typography_size(default_size, preset)

    @classmethod
    def qss(cls, default_size: int, preset: str | None = None) -> str:
        return f"{cls.resolve(default_size, preset)}pt"


class Radius:
    """全局圆角基准。"""

    DEFAULT = 6
    SMALL = 4
    MEDIUM = 8
    LARGE = 10
    XL = 12


class Border:
    """全局边框与分隔线基准。"""

    THIN = 1


class Sizing:
    """全局控件尺寸基准。

    与 Spacing（间距）解耦 --- Spacing 控制 margin/padding/gap 等"留白"，
    Sizing 控制 widget 自身的几何尺寸（高度/宽度/图标尺寸等）。
    控件高度统一通过 ``QWidget.setMinimumHeight`` 落地，QSS 不再使用
    ``min-height/max-height`` 控制控件高度，避免 QSS 盒模型与 widget
    geometry 两套语义混用导致的视觉对齐偏差（QSS min-height 不含边框，
    而 setMinimumHeight 含边框）。
    """

    CONTROL_HEIGHT = 32


class SurfaceAlpha:
    """轻量表面与状态叠加透明度。"""

    PANEL_BORDER_LIGHT = 72
    PANEL_BORDER_DARK = 104
    PANEL_BACKGROUND_LIGHT = 244
    PANEL_BACKGROUND_DARK = 232
