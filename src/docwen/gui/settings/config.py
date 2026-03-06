"""
设置界面配置管理模块

本模块统一管理所有设置界面的布局参数、样式常量和配置选项。
采用数据类和枚举提供类型安全的配置管理。

主要组件：
- SectionStyle: 区域样式枚举
- LayoutConfig: 布局参数配置
- DialogConfig: 对话框配置
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from enum import Enum
from typing import Final

logger = logging.getLogger(__name__)


class SectionStyle(Enum):
    """
    设置区域样式枚举

    定义了所有可用的区域框样式，对应ttkbootstrap的bootstyle。
    使用枚举确保样式值的类型安全。
    """

    PRIMARY = "primary"  # 主要：蓝色系
    INFO = "info"  # 信息：青色系
    SUCCESS = "success"  # 成功：绿色系
    WARNING = "warning"  # 警告：橙色系
    DANGER = "danger"  # 危险：红色系
    SECONDARY = "secondary"  # 次要：灰色系


@dataclass(frozen=True)
class LayoutConfig:
    """
    布局参数配置类

    统一管理所有UI布局相关的间距、尺寸参数。
    使用frozen=True确保配置不可变。

    属性：
        section_spacing_top: Section顶部间距（像素）
        section_spacing_bottom: Section底部间距（像素）
        section_padding: Section内部填充（像素）
        widget_spacing: 控件间距（像素）
        label_spacing: 标签与控件间距（像素）
        canvas_padding: 画布边距（像素）
        scrollbar_spacing: Section右侧与滚动条的间距（像素）
    """

    section_spacing_top: int = 15
    section_spacing_bottom: int = 0
    section_padding: int = 10
    widget_spacing: int = 5
    label_spacing: int = 5
    canvas_padding: int = 10
    scrollbar_spacing: int = 10


@dataclass(frozen=True)
class DialogConfig:
    """
    对话框参数配置类

    管理设置对话框窗口的尺寸和行为参数。

    属性：
        default_width: 默认宽度（像素）
        default_height: 默认高度（像素）
        min_width: 最小宽度（像素）
        min_height: 最小高度（像素）
        padding: 对话框内边距（像素）
        button_spacing: 按钮间距（像素）
        status_display_time: 状态消息显示时长（毫秒）
    """

    default_width: int = 700
    default_height: int = 800
    min_width: int = 510
    min_height: int = 750
    padding: int = 15
    button_spacing: int = 5
    status_display_time: int = 3000


# 全局配置实例
LAYOUT_CONFIG: Final[LayoutConfig] = LayoutConfig()
DIALOG_CONFIG: Final[DialogConfig] = DialogConfig()

logger.debug(f"配置管理模块初始化完成：布局={LAYOUT_CONFIG}, 对话框={DIALOG_CONFIG}")
