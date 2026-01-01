"""
日志设置选项卡模块

实现设置对话框的日志设置选项卡，包含：
- 日志系统开关
- 日志级别设置
- 文件输出配置
- 控制台输出配置
"""

import logging
import tkinter as tk
from typing import Dict, Any, Callable

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.gui.settings.base_tab import BaseSettingsTab
from gongwen_converter.gui.settings.config import SectionStyle

logger = logging.getLogger(__name__)


class LoggingTab(BaseSettingsTab):
    """
    日志设置选项卡类
    
    管理日志系统的所有配置选项。
    包含日志系统、文件输出和控制台输出三个主要区域。
    """
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
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
            self.scrollable_frame,
            "日志系统",
            SectionStyle.PRIMARY
        )
        
        # 日志系统开关
        self.logging_enabled_var = tk.BooleanVar(value=settings.get("enable", True))
        self.create_checkbox_with_info(
            frame,
            "启用日志系统",
            self.logging_enabled_var,
            "启用或禁用整个日志系统。\n禁用后，应用程序将不会记录任何日志信息。\n建议在生产环境中保持启用状态。",
            self._on_logging_enabled_changed
        )
        
        # 日志级别设置
        log_levels = ["debug", "info", "warning", "error", "critical"]
        self.log_level_var = tk.StringVar(value=settings.get("level", "info"))
        self.create_combobox_with_info(
            frame,
            "日志级别:",
            self.log_level_var,
            log_levels,
            "设置日志记录的详细程度。\n级别从低到高：debug > info > warning > error > critical\n级别越高，记录的日志信息越少。",
            self._on_log_level_changed
        )
        
        logger.debug("日志系统设置区域创建完成")
    
    def _on_logging_enabled_changed(self):
        """处理日志系统开关变更事件"""
        enabled = self.logging_enabled_var.get()
        logger.info(f"日志系统开关变更: {enabled}")
        self.on_change("enable", enabled)
    
    def _on_log_level_changed(self, event=None):
        """处理日志级别变更事件"""
        level = self.log_level_var.get()
        logger.info(f"日志级别变更: {level}")
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
        frame = self.create_section_frame(
            self.scrollable_frame,
            "文件输出",
            SectionStyle.INFO
        )
        
        # 文件前缀
        self.file_prefix_var = tk.StringVar(value=settings.get("file_prefix", "gongwen"))
        self.create_label_entry_pair(
            frame,
            "日志文件名前缀:",
            self.file_prefix_var,
            "设置日志文件的名称前缀"
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
            "日志保留天数:",
            self.retention_days_var,
            "设置日志文件保留的天数，超过此天数的日志将被自动删除"
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
            self.scrollable_frame,
            "控制台输出",
            SectionStyle.WARNING
        )
        
        # 控制台输出开关
        self.console_enabled_var = tk.BooleanVar(value=settings.get("console_enable", True))
        self.create_checkbox_with_info(
            frame,
            "在控制台显示日志",
            self.console_enabled_var,
            "启用后，日志将同时输出到控制台窗口",
            self._on_console_enabled_changed
        )
        
        # 控制台日志级别
        console_levels = ["debug", "info", "warning", "error", "critical"]
        self.console_level_var = tk.StringVar(value=settings.get("console_level", "debug"))
        self.create_combobox_with_info(
            frame,
            "控制台日志级别:",
            self.console_level_var,
            console_levels,
            "设置控制台输出的日志级别，可以与文件日志级别不同",
            self._on_console_level_changed
        )
        
        logger.debug("控制台输出设置区域创建完成")
    
    def _on_console_enabled_changed(self):
        """处理控制台输出开关变更事件"""
        enabled = self.console_enabled_var.get()
        logger.info(f"控制台输出开关变更: {enabled}")
        self.on_change("console_enable", enabled)
    
    def _on_console_level_changed(self, event=None):
        """处理控制台日志级别变更事件"""
        level = self.console_level_var.get()
        logger.info(f"控制台日志级别变更: {level}")
        self.on_change("console_level", level)
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置
        
        返回：
            Dict[str, Any]: 当前所有设置项的值
        """
        settings = {
            "enable": self.logging_enabled_var.get(),
            "level": self.log_level_var.get(),
            "file_prefix": self.file_prefix_var.get(),
            "retention_days": int(self.retention_days_var.get()),
            "console_enable": self.console_enabled_var.get(),
            "console_level": self.console_level_var.get()
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
