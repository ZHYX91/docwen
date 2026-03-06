"""
日志设置选项卡模块

实现设置对话框的日志设置选项卡，包含：
- 日志系统开关
- 日志级别设置
- 文件输出配置
- 控制台输出配置

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。
"""

from __future__ import annotations

import logging
import tkinter as tk
from collections.abc import Callable
from typing import Any

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.gui.settings.base_tab import BaseSettingsTab
from docwen.gui.settings.config import SectionStyle
from docwen.i18n import t

logger = logging.getLogger(__name__)


class LoggingTab(BaseSettingsTab):
    """
    日志设置选项卡类

    管理日志系统的所有配置选项。
    包含日志系统、文件输出和控制台输出三个主要区域。
    """

    def __init__(self, parent, config_manager: Any, on_change: Callable[[str, Any], None]):
        """
        初始化日志设置选项卡

        参数：
            parent: 父组件
            config_manager: 配置管理器实例
            on_change: 设置变更回调函数
        """
        super().__init__(parent, config_manager, on_change)
        logger.info("日志设置选项卡初始化完成")

    def _create_interface(self):
        """
        创建选项卡界面

        创建三个主要设置区域：
        1. 日志系统 - 总开关和日志级别
        2. 文件输出 - 文件前缀和保留天数
        3. 控制台输出 - 控制台开关和级别
        """
        logger.debug("开始创建日志设置选项卡界面")

        self._create_logging_system_section()
        self._create_file_output_section()
        self._create_console_output_section()

        logger.debug("日志设置选项卡界面创建完成")

    def _create_logging_system_section(self):
        """
        创建日志系统设置区域

        包含日志系统总开关和全局日志级别设置。

        配置路径：
        - logger_config.logging.enable
        - logger_config.logging.level
        """
        logger.debug("创建日志系统设置区域")

        settings = self.config_manager.get_logging_config()

        # 创建区域框架
        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.logging.system_section"), SectionStyle.PRIMARY
        )

        # 日志系统开关
        self.logging_enabled_var = tk.BooleanVar(value=settings.get("enable", True))
        self.create_checkbox_with_info(
            frame,
            t("settings.logging.enable_logging_label"),
            self.logging_enabled_var,
            t("settings.logging.enable_logging_tooltip"),
            self._on_logging_enabled_changed,
        )

        # 日志级别设置（显示翻译文本，存储英文值）
        # 存储值列表（用于配置文件）
        self._log_level_values = ["debug", "info", "warning", "error"]
        # 显示文本列表（用于下拉框显示）
        log_level_display = [
            t("settings.logging.levels.debug"),
            t("settings.logging.levels.info"),
            t("settings.logging.levels.warning"),
            t("settings.logging.levels.error"),
        ]
        # 创建双向映射
        self._log_level_display_to_value = dict(zip(log_level_display, self._log_level_values, strict=False))
        self._log_level_value_to_display = dict(zip(self._log_level_values, log_level_display, strict=False))

        # 获取当前值对应的显示文本
        current_value = settings.get("level", "info")
        current_display = self._log_level_value_to_display.get(current_value, log_level_display[1])

        self.log_level_var = tk.StringVar(value=current_display)
        self.create_combobox_with_info(
            frame,
            t("settings.logging.log_level_label"),
            self.log_level_var,
            log_level_display,
            t("settings.logging.log_level_tooltip"),
            self._on_log_level_changed,
        )

        logger.debug("日志系统设置区域创建完成")

    def _on_logging_enabled_changed(self):
        """处理日志系统开关变更事件"""
        enabled = self.logging_enabled_var.get()
        logger.info(f"日志系统开关变更: {enabled}")
        self.on_change("enable", enabled)

    def _on_log_level_changed(self, event=None):
        """处理日志级别变更事件（将显示文本转换为存储值）"""
        display_text = self.log_level_var.get()
        # 将显示文本转换为存储值
        level = self._log_level_display_to_value.get(display_text, "info")
        logger.info(f"日志级别变更: {display_text} -> {level}")
        self.on_change("level", level)

    def _create_file_output_section(self):
        """
        创建文件输出设置区域

        包含日志文件名前缀和保留天数设置。

        配置路径：
        - logger_config.logging.file_prefix
        - logger_config.logging.retention_days
        """
        logger.debug("创建文件输出设置区域")

        settings = self.config_manager.get_logging_config()

        # 创建区域框架
        frame = self.create_section_frame(self.scrollable_frame, t("settings.logging.file_section"), SectionStyle.INFO)

        # 文件前缀
        self.file_prefix_var = tk.StringVar(value=settings.get("file_prefix", "docwen"))
        self.create_label_entry_pair(
            frame,
            t("settings.logging.file_prefix_label"),
            self.file_prefix_var,
            t("settings.logging.file_prefix_tooltip"),
        )

        # 绑定失去焦点事件
        for child in frame.winfo_children():
            if isinstance(child, tb.Frame):
                for subchild in child.winfo_children():
                    if isinstance(subchild, tb.Entry):
                        subchild.bind("<FocusOut>", self._on_file_prefix_changed)

        # 保留天数
        self.retention_days_var = tk.StringVar(value=str(settings.get("retention_days", 30)))
        self.create_label_entry_pair(
            frame,
            t("settings.logging.retention_days_label"),
            self.retention_days_var,
            t("settings.logging.retention_days_tooltip"),
        )

        # 绑定失去焦点事件
        for child in frame.winfo_children():
            if isinstance(child, tb.Frame):
                for subchild in child.winfo_children():
                    if isinstance(subchild, tb.Entry) and subchild.cget("textvariable") == str(self.retention_days_var):
                        subchild.bind("<FocusOut>", self._on_retention_days_changed)

        logger.debug("文件输出设置区域创建完成")

    def _on_file_prefix_changed(self, event=None):
        """处理文件前缀变更事件"""
        prefix = self.file_prefix_var.get()
        logger.info(f"文件前缀变更: {prefix}")
        self.on_change("file_prefix", prefix)

    def _on_retention_days_changed(self, event=None):
        """处理保留天数变更事件"""
        try:
            days = int(self.retention_days_var.get())
            logger.info(f"保留天数变更: {days}")
            self.on_change("retention_days", days)
        except ValueError:
            logger.error("保留天数必须是整数")
            settings = self.config_manager.get_logging_config()
            self.retention_days_var.set(str(settings.get("retention_days", 30)))

    def _create_console_output_section(self):
        """
        创建控制台输出设置区域

        包含控制台输出开关和控制台日志级别设置。

        配置路径：
        - logger_config.logging.console_enable
        - logger_config.logging.console_level
        """
        logger.debug("创建控制台输出设置区域")

        settings = self.config_manager.get_logging_config()

        # 创建区域框架
        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.logging.console_section"), SectionStyle.WARNING
        )

        # 控制台输出开关
        self.console_enabled_var = tk.BooleanVar(value=settings.get("console_enable", True))
        self.create_checkbox_with_info(
            frame,
            t("settings.logging.enable_console_label"),
            self.console_enabled_var,
            t("settings.logging.enable_console_tooltip"),
            self._on_console_enabled_changed,
        )

        # 控制台日志级别（显示翻译文本，存储英文值）
        # 复用日志系统区域创建的映射（两个下拉框使用相同的级别选项）
        console_level_display = [
            t("settings.logging.levels.debug"),
            t("settings.logging.levels.info"),
            t("settings.logging.levels.warning"),
            t("settings.logging.levels.error"),
        ]

        # 获取当前值对应的显示文本
        current_value = settings.get("console_level", "debug")
        current_display = self._log_level_value_to_display.get(current_value, console_level_display[0])

        self.console_level_var = tk.StringVar(value=current_display)
        self.create_combobox_with_info(
            frame,
            t("settings.logging.console_level_label"),
            self.console_level_var,
            console_level_display,
            t("settings.logging.console_level_tooltip"),
            self._on_console_level_changed,
        )

        logger.debug("控制台输出设置区域创建完成")

    def _on_console_enabled_changed(self):
        """处理控制台输出开关变更事件"""
        enabled = self.console_enabled_var.get()
        logger.info(f"控制台输出开关变更: {enabled}")
        self.on_change("console_enable", enabled)

    def _on_console_level_changed(self, event=None):
        """处理控制台日志级别变更事件（将显示文本转换为存储值）"""
        display_text = self.console_level_var.get()
        # 将显示文本转换为存储值
        level = self._log_level_display_to_value.get(display_text, "debug")
        logger.info(f"控制台日志级别变更: {display_text} -> {level}")
        self.on_change("console_level", level)

    def get_settings(self) -> dict[str, Any]:
        """
        获取当前选项卡的设置

        将显示文本转换为存储值后返回。

        返回：
            Dict[str, Any]: 当前所有设置项的值（存储格式）
        """
        # 将日志级别显示文本转换为存储值
        level_display = self.log_level_var.get()
        level_value = self._log_level_display_to_value.get(level_display, "info")

        console_level_display = self.console_level_var.get()
        console_level_value = self._log_level_display_to_value.get(console_level_display, "debug")

        settings = {
            "enable": self.logging_enabled_var.get(),
            "level": level_value,
            "file_prefix": self.file_prefix_var.get(),
            "retention_days": int(self.retention_days_var.get()),
            "console_enable": self.console_enabled_var.get(),
            "console_level": console_level_value,
        }

        logger.debug(f"获取日志设置: {settings}")
        return settings

    def apply_settings(self) -> bool:
        """
        应用当前设置到配置文件

        将所有设置项保存到对应的配置路径。

        返回：
            bool: 应用是否成功
        """
        logger.debug("开始应用日志设置到配置文件")

        try:
            settings = self.get_settings()
            success = True

            # 更新配置文件
            for key, value in settings.items():
                result = self.config_manager.update_config_value("logger_config", "logging", key, value)
                if not result:
                    logger.error(f"更新配置失败: logging.{key}")
                    success = False

            if success:
                logger.info("✓ 日志设置已成功应用到配置文件")
            else:
                logger.error("✗ 部分日志设置更新失败")

            return success

        except Exception as e:
            logger.error(f"应用日志设置失败: {e}", exc_info=True)
            return False
