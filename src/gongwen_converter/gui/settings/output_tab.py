"""
输出设置选项卡模块

实现设置对话框的输出设置选项卡，包含：
- 中间文件保存策略
- 输出目录设置
- 日期子文件夹设置
- 输出行为配置
"""

import logging
import os
import tkinter as tk
from typing import Dict, Any, Callable

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.gui.settings.base_tab import BaseSettingsTab
from gongwen_converter.gui.settings.config import SectionStyle

logger = logging.getLogger(__name__)


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
            "中间文件保存策略",
            SectionStyle.PRIMARY
        )
        
        # 说明文本
        desc = tb.Label(
            frame,
            text="多步转换时（如 XLS→XLSX→ODS），是否保存中间文件XLSX",
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
            "保存中间文件到输出目录",
            self.save_intermediate_var,
            "启用：保存所有中间步骤文件\n"
            "禁用：只保存最终结果文件\n\n"
            "示例（XLS→XLSX→ODS）：\n"
            "• 启用：保存 XLSX 和 ODS\n"
            "• 禁用：只保存 ODS",
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
            "输出目录设置",
            SectionStyle.INFO
        )
        
        # 获取当前配置
        settings = self.config_manager.get_output_directory_settings()
        mode = settings.get("mode", "source")
        custom_path = settings.get("custom_path", "")
        
        # 1. 输出模式选择
        mode_display = "原文件位置" if mode == "source" else "自定义位置"
        self.output_mode_var = tk.StringVar(value=mode_display)
        self.create_combobox_with_info(
            frame,
            "输出位置:",
            self.output_mode_var,
            ["原文件位置", "自定义位置"],
            "原文件位置：输出到原文件所在文件夹\n"
            "自定义位置：输出到指定文件夹",
            self._on_output_mode_changed
        )
        
        # 2. 自定义路径输入区域（始终显示，根据模式启用/禁用）
        custom_path_label_frame = tb.Frame(frame)
        custom_path_label_frame.pack(fill="x", pady=(10, 5))
        
        tb.Label(custom_path_label_frame, text="自定义路径:", bootstyle="secondary").pack(anchor="w")
        
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
            text="浏览",
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
            "日期子文件夹设置",
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
            "创建日期子文件夹",
            self.date_subfolder_var,
            "在输出目录下创建日期子文件夹\n"
            "例如: 输出目录/2025-01-06/文件.pdf\n"
            "此选项对两种输出模式都有效",
            self._on_date_subfolder_changed
        )
        
        # 2. 日期格式设置（始终显示，根据复选框启用/禁用）
        format_label_frame = tb.Frame(frame)
        format_label_frame.pack(fill="x", pady=(10, 5))
        
        tb.Label(format_label_frame, text="日期格式:", bootstyle="secondary").pack(anchor="w")
        
        # 日期格式映射
        self.date_format_mapping = {
            "2025-01-06": "%Y-%m-%d",
            "20250106": "%Y%m%d",
            "2025年01月06日": "%Y年%m月%d日"
        }
        
        # 获取当前格式的显示值
        current_format_display = None
        for display, fmt in self.date_format_mapping.items():
            if fmt == date_folder_format:
                current_format_display = display
                break
        if not current_format_display:
            current_format_display = "2025-01-06"
        
        self.date_format_var = tk.StringVar(value=current_format_display)
        self.date_format_combo = tb.Combobox(
            frame,
            textvariable=self.date_format_var,
            values=list(self.date_format_mapping.keys()),
            state="readonly",
            bootstyle="secondary"
        )
        self.date_format_combo.pack(fill="x", pady=(0, 5))
        self.date_format_combo.bind("<<ComboboxSelected>>", self._on_date_format_changed)
        
        # 初始化控件状态
        self._update_date_format_state()
        
        logger.debug("日期子文件夹设置区域创建完成")
    
    def _on_output_mode_changed(self, event=None):
        """处理输出模式变更"""
        display_value = self.output_mode_var.get()
        config_value = "source" if display_value == "原文件位置" else "custom"
        logger.info(f"输出模式变更: {display_value} → {config_value}")
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
            title="选择输出目录",
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
    
    def _on_date_format_changed(self, event=None):
        """处理日期格式变更"""
        display_value = self.date_format_var.get()
        config_value = self.date_format_mapping.get(display_value, "%Y-%m-%d")
        logger.info(f"日期格式变更: {display_value} → {config_value}")
        self.on_change("date_folder_format", config_value)
    
    def _update_custom_path_state(self):
        """更新自定义路径控件的启用/禁用状态"""
        mode = self.output_mode_var.get()
        
        if mode == "自定义位置":
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
            "输出行为设置",
            SectionStyle.WARNING
        )
        
        # 获取当前配置
        behavior_settings = self.config_manager.get_output_behavior_settings()
        auto_open_folder = behavior_settings.get("auto_open_folder", False)
        
        # 自动打开文件夹设置
        self.auto_open_folder_var = tk.BooleanVar(value=auto_open_folder)
        self.create_checkbox_with_info(
            frame,
            "转换完成后自动打开输出文件夹",
            self.auto_open_folder_var,
            "转换或校对成功后自动打开输出文件夹并选中最终文件\n"
            "方便快速查看转换结果",
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
        # 转换显示值到配置值
        output_display = self.output_mode_var.get()
        output_config = "source" if output_display == "原文件位置" else "custom"
        
        date_format_display = self.date_format_var.get()
        date_format_config = self.date_format_mapping.get(date_format_display, "%Y-%m-%d")
        
        settings = {
            "save_intermediate_files": self.save_intermediate_var.get(),
            "output_mode": output_config,
            "custom_path": self.custom_path_var.get(),
            "create_date_subfolder": self.date_subfolder_var.get(),
            "date_folder_format": date_format_config,
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
