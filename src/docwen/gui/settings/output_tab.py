"""
输出设置选项卡模块

实现设置对话框的输出设置选项卡，包含：
- 中间文件保存策略
- 输出目录设置
- 日期子文件夹设置
- 输出行为配置

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。
使用 ConfigCombobox 组件实现配置值与显示文本的分离，
避免语言切换时的映射问题。
"""

import logging
import os
import tkinter as tk
from typing import Dict, Any, Callable

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.gui.settings.base_tab import BaseSettingsTab
from docwen.gui.settings.config import SectionStyle
from docwen.gui.components.config_combobox import ConfigCombobox
from docwen.i18n import t

logger = logging.getLogger(__name__)

# 输出模式的配置值和翻译键映射
OUTPUT_MODE_CONFIG_VALUES = ["source", "custom"]
OUTPUT_MODE_TRANSLATE_KEYS = {
    "source": "settings.output.output_modes.source",
    "custom": "settings.output.output_modes.custom"
}

# 日期格式的配置值和翻译键映射
DATE_FORMAT_CONFIG_VALUES = ["%Y-%m-%d", "%Y%m%d", "%Y年%m月%d日"]
DATE_FORMAT_TRANSLATE_KEYS = {
    "%Y-%m-%d": "settings.output.date_formats.iso",
    "%Y%m%d": "settings.output.date_formats.compact",
    "%Y年%m月%d日": "settings.output.date_formats.chinese"
}


class OutputTab(BaseSettingsTab):
    """
    输出设置选项卡类
    
    管理输出过程相关的所有配置选项。
    包含中间文件、输出目录、日期子文件夹和输出行为四个主要区域。
    """
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
        """初始化输出设置选项卡"""
        super().__init__(parent, config_manager, on_change)
        logger.info("输出设置选项卡初始化完成")
    
    def _create_interface(self):
        """
        创建选项卡界面
        
        创建三个主要设置区域：
        1. 输出目录设置 - 文件输出位置配置
        2. 日期子文件夹设置 - 按日期组织输出文件
        3. 输出行为设置 - 转换完成后的操作行为
        """
        logger.debug("开始创建输出设置选项卡界面")
        
        self._create_intermediate_files_section()
        self._create_output_directory_section()
        self._create_date_subfolder_section()
        self._create_output_behavior_section()
        
        logger.debug("输出设置选项卡界面创建完成")
    
    def _create_intermediate_files_section(self):
        """
        创建中间文件保存设置区域
        
        配置多步转换（如 XLS→XLSX→ODS）时是否保存中间步骤的文件。
        
        配置路径：output_config.intermediate_files.save_to_output
        """
        logger.debug("创建中间文件设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.output.intermediate_section"),
            SectionStyle.PRIMARY
        )
        
        # 说明文本
        desc = tb.Label(
            frame,
            text=t("settings.output.save_intermediate_desc"),
            bootstyle="secondary",
            justify="left"
        )
        desc.pack(anchor="w", pady=(0, 10))
        
        # 获取当前配置
        intermediate_settings = self.config_manager.get_output_intermediate_files_config()
        save_to_output = intermediate_settings.get("save_to_output", True)
        
        # 保存中间文件复选框
        self.save_intermediate_var = tk.BooleanVar(value=save_to_output)
        self.create_checkbox_with_info(
            frame,
            t("settings.output.save_intermediate_label"),
            self.save_intermediate_var,
            t("settings.output.save_intermediate_tooltip"),
            self._on_intermediate_changed
        )
        
        logger.debug("中间文件设置区域创建完成")
    
    def _on_intermediate_changed(self):
        """处理中间文件设置变更"""
        value = self.save_intermediate_var.get()
        logger.info(f"中间文件保存设置变更: {value}")
        self.on_change("save_intermediate_files", value)
    
    def _create_output_directory_section(self):
        """
        创建输出目录设置区域
        
        配置文件输出位置，支持原文件位置和自定义位置两种模式。
        
        配置路径：
        - output_config.directory.mode
        - output_config.directory.custom_path
        """
        logger.debug("创建输出目录设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.output.directory_section"),
            SectionStyle.INFO
        )
        
        # 获取当前配置
        settings = self.config_manager.get_output_directory_settings()
        mode = settings.get("mode", "source")
        custom_path = settings.get("custom_path", "")
        
        # 1. 输出模式选择（使用 ConfigCombobox）
        self.output_mode_combo = self._create_config_combobox(
            frame,
            t("settings.output.output_mode_label"),
            OUTPUT_MODE_CONFIG_VALUES,
            OUTPUT_MODE_TRANSLATE_KEYS,
            mode,
            t("settings.output.output_mode_tooltip"),
            self._on_output_mode_changed
        )
        
        # 2. 自定义路径输入区域（始终显示，根据模式启用/禁用）
        custom_path_label_frame = tb.Frame(frame)
        custom_path_label_frame.pack(fill="x", pady=(10, 5))
        
        tb.Label(custom_path_label_frame, text=t("settings.output.custom_path_label"), bootstyle="secondary").pack(anchor="w")
        
        path_input_frame = tb.Frame(frame)
        path_input_frame.pack(fill="x", pady=(0, 10))
        
        self.custom_path_var = tk.StringVar(value=custom_path)
        self.custom_path_entry = tb.Entry(
            path_input_frame,
            textvariable=self.custom_path_var,
            bootstyle="secondary"
        )
        self.custom_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.browse_btn = tb.Button(
            path_input_frame,
            text=t("common.browse"),
            bootstyle="info-outline",
            width=8,
            command=self._browse_custom_path
        )
        self.browse_btn.pack(side="right")
        
        # 初始化控件状态
        self._update_custom_path_state()
        
        logger.debug("输出目录设置区域创建完成")
    
    def _create_date_subfolder_section(self):
        """
        创建日期子文件夹设置区域
        
        配置是否在输出目录下创建日期子文件夹，以及日期格式。
        
        配置路径：
        - output_config.directory.create_date_subfolder
        - output_config.directory.date_folder_format
        """
        logger.debug("创建日期子文件夹设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.output.date_folder.section"),
            SectionStyle.SUCCESS
        )
        
        # 获取当前配置
        settings = self.config_manager.get_output_directory_settings()
        create_date_subfolder = settings.get("create_date_subfolder", False)
        date_folder_format = settings.get("date_folder_format", "%Y-%m-%d")
        
        # 1. 创建日期子文件夹复选框
        self.date_subfolder_var = tk.BooleanVar(value=create_date_subfolder)
        self.create_checkbox_with_info(
            frame,
            t("settings.output.date_folder.create_label"),
            self.date_subfolder_var,
            t("settings.output.date_folder.create_tooltip"),
            self._on_date_subfolder_changed
        )
        
        # 2. 日期格式设置（使用 ConfigCombobox）
        format_label_frame = tb.Frame(frame)
        format_label_frame.pack(fill="x", pady=(10, 5))
        
        tb.Label(format_label_frame, text=t("settings.output.date_folder.format_label"), bootstyle="secondary").pack(anchor="w")
        
        self.date_format_combo = ConfigCombobox(
            frame,
            config_values=DATE_FORMAT_CONFIG_VALUES,
            translate_keys=DATE_FORMAT_TRANSLATE_KEYS,
            initial_value=date_folder_format,
            on_change=self._on_date_format_changed
        )
        self.date_format_combo.pack(fill="x", pady=(0, 5))
        
        # 初始化控件状态
        self._update_date_format_state()
        
        logger.debug("日期子文件夹设置区域创建完成")
    
    def _create_config_combobox(
        self,
        parent: tk.Widget,
        label_text: str,
        config_values: list,
        translate_keys: dict,
        initial_value: str,
        tooltip: str,
        on_change: Callable[[str], None]
    ) -> ConfigCombobox:
        """
        创建带标签和信息图标的配置下拉框
        
        参数:
            parent: 父组件
            label_text: 标签文本
            config_values: 配置值列表
            translate_keys: 翻译键映射
            initial_value: 初始配置值
            tooltip: 工具提示文本
            on_change: 值变更回调函数
            
        返回:
            ConfigCombobox: 创建的下拉框组件
        """
        from docwen.utils.gui_utils import create_info_icon
        
        # 创建容器
        container = tb.Frame(parent)
        container.pack(fill="x", pady=(0, self.layout_config.widget_spacing))
        
        # 创建标签行
        label_frame = tb.Frame(container)
        label_frame.pack(fill="x", pady=(0, self.layout_config.label_spacing))
        
        label = tb.Label(label_frame, text=label_text, bootstyle="secondary")
        label.pack(side="left")
        
        info = create_info_icon(label_frame, tooltip, "info")
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 创建配置下拉框
        combobox = ConfigCombobox(
            container,
            config_values=config_values,
            translate_keys=translate_keys,
            initial_value=initial_value,
            on_change=on_change
        )
        combobox.pack(fill="x")
        
        return combobox
    
    def _on_output_mode_changed(self, config_value: str):
        """处理输出模式变更"""
        logger.info(f"输出模式变更: {config_value}")
        self.on_change("output_mode", config_value)
        
        # 更新自定义路径控件状态
        self._update_custom_path_state()
    
    def _browse_custom_path(self):
        """浏览并选择自定义输出目录"""
        from tkinter import filedialog
        
        initial_dir = self.custom_path_var.get()
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~")
        
        selected_dir = filedialog.askdirectory(
            title=t("settings.output.browse_title"),
            initialdir=initial_dir
        )
        
        if selected_dir:
            self.custom_path_var.set(selected_dir)
            logger.info(f"选择自定义输出目录: {selected_dir}")
            self.on_change("custom_path", selected_dir)
    
    def _on_date_subfolder_changed(self):
        """处理日期子文件夹开关变更"""
        value = self.date_subfolder_var.get()
        logger.info(f"日期子文件夹设置变更: {value}")
        self.on_change("create_date_subfolder", value)
        
        # 更新日期格式控件状态
        self._update_date_format_state()
    
    def _on_date_format_changed(self, config_value: str):
        """处理日期格式变更"""
        logger.info(f"日期格式变更: {config_value}")
        self.on_change("date_folder_format", config_value)
    
    def _update_custom_path_state(self):
        """更新自定义路径控件的启用/禁用状态"""
        mode = self.output_mode_combo.get_config_value()
        
        if mode == "custom":
            # 启用自定义路径输入
            self.custom_path_entry.configure(state="normal")
            self.browse_btn.configure(state="normal")
            logger.debug("自定义路径控件已启用")
        else:
            # 禁用自定义路径输入
            self.custom_path_entry.configure(state="disabled")
            self.browse_btn.configure(state="disabled")
            logger.debug("自定义路径控件已禁用")
    
    def _update_date_format_state(self):
        """更新日期格式控件的启用/禁用状态"""
        if self.date_subfolder_var.get():
            # 启用日期格式选择
            self.date_format_combo.configure(state="readonly")
            logger.debug("日期格式控件已启用")
        else:
            # 禁用日期格式选择
            self.date_format_combo.configure(state="disabled")
            logger.debug("日期格式控件已禁用")
    
    def _create_output_behavior_section(self):
        """
        创建输出行为设置区域
        
        配置转换完成后的自动操作行为。
        
        配置路径：output_config.behavior.auto_open_folder
        """
        logger.debug("创建输出行为设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.output.behavior.section"),
            SectionStyle.WARNING
        )
        
        # 获取当前配置
        behavior_settings = self.config_manager.get_output_behavior_settings()
        auto_open_folder = behavior_settings.get("auto_open_folder", False)
        
        # 自动打开文件夹设置
        self.auto_open_folder_var = tk.BooleanVar(value=auto_open_folder)
        self.create_checkbox_with_info(
            frame,
            t("settings.output.behavior.auto_open_folder_label"),
            self.auto_open_folder_var,
            t("settings.output.behavior.auto_open_folder_tooltip"),
            self._on_auto_open_changed
        )
        
        logger.debug("输出行为设置区域创建完成")
    
    def _on_auto_open_changed(self):
        """处理自动打开文件夹开关变更"""
        value = self.auto_open_folder_var.get()
        logger.info(f"自动打开文件夹设置变更: {value}")
        self.on_change("auto_open_folder", value)
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置
        
        返回：
            Dict[str, Any]: 当前所有设置项的值
        """
        settings = {
            "save_intermediate_files": self.save_intermediate_var.get(),
            "output_mode": self.output_mode_combo.get_config_value(),
            "custom_path": self.custom_path_var.get(),
            "create_date_subfolder": self.date_subfolder_var.get(),
            "date_folder_format": self.date_format_combo.get_config_value(),
            "auto_open_folder": self.auto_open_folder_var.get()
        }
        
        logger.debug(f"获取输出设置: {settings}")
        return settings
    
    def apply_settings(self) -> bool:
        """
        应用当前设置到配置文件
        
        将所有设置项保存到对应的配置路径。
        
        返回：
            bool: 应用是否成功
        """
        logger.debug("开始应用输出设置到配置文件")
        
        try:
            settings = self.get_settings()
            success = True
            
            # 中间文件设置
            if not self.config_manager.update_config_value(
                "output_config",
                "intermediate_files",
                "save_to_output",
                settings["save_intermediate_files"]
            ):
                success = False
            
            # 输出目录模式
            if not self.config_manager.update_config_value(
                "output_config", 
                "directory", 
                "mode", 
                settings["output_mode"]
            ):
                success = False
            
            # 自定义路径
            if not self.config_manager.update_config_value(
                "output_config", 
                "directory", 
                "custom_path", 
                settings["custom_path"]
            ):
                success = False
            
            # 日期子文件夹开关
            if not self.config_manager.update_config_value(
                "output_config", 
                "directory", 
                "create_date_subfolder", 
                settings["create_date_subfolder"]
            ):
                success = False
            
            # 日期格式
            if not self.config_manager.update_config_value(
                "output_config", 
                "directory", 
                "date_folder_format", 
                settings["date_folder_format"]
            ):
                success = False
            
            # 自动打开文件夹
            if not self.config_manager.update_config_value(
                "output_config", 
                "behavior", 
                "auto_open_folder", 
                settings["auto_open_folder"]
            ):
                success = False
            
            if success:
                logger.info("✓ 输出设置已成功应用到配置文件")
            else:
                logger.error("✗ 部分输出设置更新失败")
            
            return success
            
        except Exception as e:
            logger.error(f"应用输出设置失败: {e}", exc_info=True)
            return False
