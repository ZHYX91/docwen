"""
通用设置选项卡模块

实现设置对话框的通用设置选项卡，包含：
- 界面主题选择
- 窗口透明度设置
- 窗口行为配置
"""

import logging
import tkinter as tk
from typing import Dict, Any, Callable

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.gui.settings.base_tab import BaseSettingsTab
from gongwen_converter.gui.settings.config import SectionStyle

logger = logging.getLogger(__name__)


class GeneralTab(BaseSettingsTab):
    """
    通用设置选项卡类
    
    管理应用程序的通用界面和窗口行为设置。
    包含主题、透明度和窗口行为三个主要区域。
    """
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
        """
        初始化通用设置选项卡
        
        参数：
            parent: 父组件
            config_manager: 配置管理器实例
            on_change: 设置变更回调函数
        """
        # 调用基类初始化（基类会自动调用_create_interface）
        super().__init__(parent, config_manager, on_change)
        logger.info("通用设置选项卡初始化完成")
    
    def _create_interface(self):
        """
        创建选项卡界面
        
        创建三个主要设置区域：
        1. 主题设置 - 界面主题选择和预览
        2. 透明度设置 - 窗口透明效果配置
        3. 窗口行为 - 窗口位置、大小和默认模式
        """
        logger.debug("开始创建通用设置选项卡界面")
        
        self._create_theme_section()
        self._create_transparency_section()
        self._create_window_behavior_section()
        self._create_template_section()
        
        logger.debug("通用设置选项卡界面创建完成")
    
    def _create_theme_section(self):
        """
        创建主题设置区域
        
        包含主题选择下拉框和主题预览示例。
        主题变更会在保存时应用，而非实时预览。
        
        配置路径：gui_config.theme.default_theme
        """
        logger.debug("创建主题设置区域")
        
        # 创建区域框架
        frame = self.create_section_frame(
            self.scrollable_frame,
            "界面主题",
            SectionStyle.PRIMARY
        )
        
        # 获取主题映射
        available_themes, self.theme_mapping = self._get_available_themes()
        
        # 获取当前主题并转换为中文显示
        current_theme = self.config_manager.get_default_theme()
        english_to_chinese = {v: k for k, v in self.theme_mapping.items()}
        current_theme_chinese = english_to_chinese.get(current_theme, current_theme)
        
        # 创建主题选择下拉框
        self.theme_var = tk.StringVar(value=current_theme_chinese)
        self.create_combobox_with_info(
            frame,
            "选择应用程序主题:",
            self.theme_var,
            available_themes,
            "选择应用程序的视觉主题。\n不同的主题提供不同的颜色方案和界面风格。\n更改主题会在保存后生效。",
            self._on_theme_changed
        )
        
        # 主题预览区域
        preview_label = tb.Label(
            frame,
            text="主题预览:",
            bootstyle="secondary"
        )
        preview_label.pack(anchor="w", pady=(10, 5))
        
        preview_frame = tb.Frame(frame, bootstyle="default", height=60)
        preview_frame.pack(fill="x", pady=(0, 5))
        preview_frame.pack_propagate(True)
        
        # 添加预览控件
        self._create_theme_preview(preview_frame)
        
        logger.debug("主题设置区域创建完成")
    
    def _get_available_themes(self):
        """
        获取可用主题列表和映射
        
        返回：
            tuple: (中文主题名称列表, 中文->英文映射字典)
        """
        logger.debug("获取可用主题列表")
        
        # 主题名称映射：英文 -> 中文
        theme_name_mapping = {
            "cerculean": "蔚蓝主题",
            "cosmo": "宇宙主题",
            "cyborg": "赛博主题",
            "darkly": "暗黑主题",
            "flatly": "扁平主题",
            "journal": "杂志主题",
            "litera": "文学主题",
            "lumen": "光感主题",
            "minty": "薄荷主题",
            "morph": "渐变主题",
            "pulse": "脉冲主题",
            "sandstone": "砂岩主题",
            "simplex": "极简主题",
            "solar": "阳光主题",
            "superhero": "超英主题",
            "united": "联合主题",
            "vapor": "蒸汽主题",
            "yeti": "雪人主题"
        }
        
        # 创建反向映射：中文 -> 英文
        reverse_mapping = {chinese: english for english, chinese in theme_name_mapping.items()}
        
        # 返回排序后的中文名称列表
        chinese_themes = sorted(reverse_mapping.keys())
        
        logger.debug(f"可用主题数量: {len(chinese_themes)}")
        return chinese_themes, reverse_mapping
    
    def _create_theme_preview(self, parent: tb.Frame):
        """
        创建主题预览示例
        
        包含示例按钮和标签，用于预览当前主题效果。
        
        参数：
            parent: 预览框架
        """
        logger.debug("创建主题预览")
        
        # 示例按钮
        sample_button = tb.Button(
            parent,
            text="示例按钮",
            bootstyle="primary",
            width=10
        )
        sample_button.pack(side="left", padx=10, pady=10)
        
        # 示例标签
        sample_label = tb.Label(
            parent,
            text="示例文本",
            bootstyle="secondary"
        )
        sample_label.pack(side="right", padx=10, pady=10)
        
        logger.debug("主题预览创建完成")
    
    def _on_theme_changed(self, event=None):
        """
        处理主题变更事件
        
        主题变更不会立即应用，只记录变更。
        实际应用发生在用户点击"应用"或"确定"按钮时。
        """
        new_theme_chinese = self.theme_var.get()
        new_theme_english = self.theme_mapping.get(new_theme_chinese, new_theme_chinese)
        
        logger.info(f"主题已选择: {new_theme_chinese} -> {new_theme_english}（将在保存时应用）")
        self.on_change("default_theme", new_theme_english)
    
    def _create_template_section(self):
        """
        创建模板设置区域
        
        配置MD文件的默认模板类型选择。
        
        配置路径：gui_config.template.md_default_template
        """
        logger.debug("创建模板设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "模板设置",
            SectionStyle.SUCCESS
        )
        
        # 获取当前配置
        current_template_type = self.config_manager.get_default_md_template_type()
        current_display = "文档模板(DOCX)" if current_template_type == "docx" else "表格模板(XLSX)"
        
        # 模板类型选择
        self.template_type_var = tk.StringVar(value=current_display)
        self.create_combobox_with_info(
            frame,
            "MD文件默认模板类型:",
            self.template_type_var,
            ["文档模板(DOCX)", "表格模板(XLSX)"],
            "设置MD文件的默认模板选择\n• DOCX：文档模板\n• XLSX：表格模板",
            self._on_template_changed
        )
        
        logger.debug("模板设置区域创建完成")
    
    def _on_template_changed(self, event=None):
        """处理模板类型变更"""
        display = self.template_type_var.get()
        config_value = "docx" if display == "文档模板(DOCX)" else "xlsx"
        logger.info(f"模板类型变更: {display} → {config_value}")
        self.on_change("md_default_template", config_value)
    
    def _create_transparency_section(self):
        """
        创建透明度设置区域
        
        包含透明度开关和滑块控件。
        透明度范围：75%-99%
        
        配置路径：
        - gui_config.transparency.enabled
        - gui_config.transparency.default_value
        """
        logger.debug("创建透明度设置区域")
        
        # 创建区域框架
        frame = self.create_section_frame(
            self.scrollable_frame,
            "窗口透明度",
            SectionStyle.INFO
        )
        
        # 透明度启用开关
        self.transparency_enabled_var = tk.BooleanVar(
            value=self.config_manager.is_transparency_enabled()
        )
        self.create_checkbox_with_info(
            frame,
            "启用窗口透明度",
            self.transparency_enabled_var,
            "启用窗口透明效果。\n启用后可以调整窗口的透明度级别，\n让应用程序窗口呈现半透明效果。",
            self._on_transparency_toggled
        )
        
        # 透明度滑块
        slider_frame = tb.Frame(frame)
        slider_frame.pack(fill="x", pady=(0, 5))
        
        transparency_value = self.config_manager.get_transparency_value()
        self.transparency_var = tk.DoubleVar(value=transparency_value)
        
        transparency_slider = tb.Scale(
            slider_frame,
            from_=0.75,
            to=0.99,
            orient="horizontal",
            variable=self.transparency_var,
            bootstyle="primary",
            command=self._on_transparency_changed
        )
        transparency_slider.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        # 透明度值标签
        transparency_percent = int(transparency_value * 100)
        self.transparency_value_label = tb.Label(
            slider_frame,
            text=f"{transparency_percent}%",
            width=4,
            bootstyle="secondary"
        )
        self.transparency_value_label.pack(side="right")
        
        # 初始化控件状态
        self._update_transparency_controls_state()
        
        logger.debug("透明度设置区域创建完成")
    
    def _on_transparency_toggled(self):
        """
        处理透明度开关变更事件
        
        更新透明度控件的启用状态。
        如果启用，立即应用当前透明度值作为预览。
        """
        enabled = self.transparency_enabled_var.get()
        logger.info(f"透明度启用状态变更: {enabled}")
        
        self.on_change("transparency_enabled", enabled)
        self._update_transparency_controls_state()
        
        # 立即应用到对话框（预览效果）
        try:
            root = self.winfo_toplevel()
            if enabled:
                root.attributes('-alpha', self.transparency_var.get())
            else:
                root.attributes('-alpha', 1.0)
            logger.debug(f"对话框透明度已{'启用' if enabled else '禁用'}")
        except Exception as e:
            logger.error(f"设置对话框透明度失败: {e}")
    
    def _update_transparency_controls_state(self):
        """
        更新透明度控件的启用/禁用状态
        
        根据透明度开关状态，启用或禁用滑块和值标签。
        """
        enabled = self.transparency_enabled_var.get()
        state = "normal" if enabled else "disabled"
        
        # 更新滑块状态
        for child in self.transparency_value_label.master.winfo_children():
            if isinstance(child, tb.Scale):
                child.configure(state=state)
        
        # 更新值标签状态
        self.transparency_value_label.configure(state=state)
        
        logger.debug(f"透明度控件状态已更新: {state}")
    
    def _on_transparency_changed(self, value):
        """
        处理透明度滑块变更事件
        
        更新透明度值标签，并立即应用到对话框作为预览。
        
        参数：
            value: 透明度值（0.75-0.99）
        """
        transparency_value = float(value)
        transparency_value = max(0.75, min(0.99, transparency_value))
        
        logger.info(f"透明度变更: {transparency_value}")
        self.on_change("default_transparency", transparency_value)
        
        # 更新值标签
        transparency_percent = int(transparency_value * 100)
        self.transparency_value_label.configure(text=f"{transparency_percent}%")
        
        # 立即应用到对话框（预览效果）
        if self.transparency_enabled_var.get():
            try:
                root = self.winfo_toplevel()
                root.attributes('-alpha', transparency_value)
                logger.debug(f"对话框透明度已设置为: {transparency_value}")
            except Exception as e:
                logger.error(f"设置对话框透明度失败: {e}")
    
    def _create_window_behavior_section(self):
        """
        创建窗口行为设置区域
        
        包含窗口位置记忆、自动居中和默认模式设置。
        
        配置路径：
        - gui_config.window.remember_gui_state
        - gui_config.window.auto_center
        - gui_config.window.default_mode
        """
        logger.debug("创建窗口行为设置区域")
        
        # 创建区域框架
        frame = self.create_section_frame(
            self.scrollable_frame,
            "窗口行为",
            SectionStyle.WARNING
        )
        
        # 记住窗口位置和大小
        self.remember_state_var = tk.BooleanVar(
            value=self.config_manager.should_remember_gui_state()
        )
        self.create_checkbox_with_info(
            frame,
            "记住窗口位置和大小",
            self.remember_state_var,
            "记住应用程序窗口的位置和尺寸。\n下次启动时，窗口会自动恢复到上次关闭时的状态。\n这有助于保持您习惯的工作环境。",
            self._on_remember_state_changed
        )
        
        # 启动时自动居中窗口
        self.auto_center_var = tk.BooleanVar(
            value=self.config_manager.should_auto_center()
        )
        self.create_checkbox_with_info(
            frame,
            "启动时自动居中窗口",
            self.auto_center_var,
            "应用程序启动时自动将窗口居中显示在屏幕上。\n如果同时启用了“记住窗口位置”，此设置将优先于记住的位置。",
            self._on_auto_center_changed
        )
        
        # 默认文件处理模式
        mode_options = ["单文件模式", "批量模式"]
        current_mode = self.config_manager.get_default_mode()
        current_display_value = "单文件模式" if current_mode == "single" else "批量模式"
        
        self.default_mode_var = tk.StringVar(value=current_display_value)
        self.create_combobox_with_info(
            frame,
            "默认文件处理模式:",
            self.default_mode_var,
            mode_options,
            "设置应用程序启动时的默认文件处理模式。\n单文件模式：每次处理一个文件\n批量模式：可以同时处理多个文件",
            self._on_default_mode_changed
        )
        
        logger.debug("窗口行为设置区域创建完成")
    
    def _on_remember_state_changed(self):
        """处理记住窗口状态变更事件"""
        remember = self.remember_state_var.get()
        logger.info(f"记住窗口状态设置变更: {remember}")
        self.on_change("remember_gui_state", remember)
    
    def _on_auto_center_changed(self):
        """处理自动居中变更事件"""
        auto_center = self.auto_center_var.get()
        logger.info(f"自动居中设置变更: {auto_center}")
        self.on_change("auto_center", auto_center)
    
    def _on_default_mode_changed(self, event=None):
        """
        处理默认模式变更事件
        
        将显示值转换为配置值并通知变更。
        """
        new_mode_display = self.default_mode_var.get()
        new_mode_config = "single" if new_mode_display == "单文件模式" else "batch"
        
        logger.info(f"默认模式变更: {new_mode_display} -> {new_mode_config}")
        self.on_change("default_mode", new_mode_config)
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置
        
        返回：
            Dict[str, Any]: 当前所有设置项的值
        """
        # 转换主题名称
        current_theme_chinese = self.theme_var.get()
        current_theme_english = self.theme_mapping.get(current_theme_chinese, current_theme_chinese)
        
        # 转换模板类型
        template_display = self.template_type_var.get()
        template_config = "docx" if template_display == "文档模板(DOCX)" else "xlsx"
        
        # 转换默认模式
        mode_display = self.default_mode_var.get()
        mode_config = "single" if mode_display == "单文件模式" else "batch"
        
        settings = {
            "default_theme": current_theme_english,
            "md_default_template": template_config,
            "transparency_enabled": self.transparency_enabled_var.get(),
            "default_transparency": self.transparency_var.get(),
            "remember_gui_state": self.remember_state_var.get(),
            "auto_center": self.auto_center_var.get(),
            "default_mode": mode_config
        }
        
        logger.debug(f"获取通用设置: {settings}")
        return settings
    
    def apply_settings(self) -> bool:
        """
        应用当前设置到配置文件
        
        将所有设置项保存到对应的配置路径。
        
        返回：
            bool: 应用是否成功
        """
        logger.debug("开始应用通用设置到配置文件")
        
        try:
            settings = self.get_settings()
            success = True
            
            # 更新配置文件
            for key, value in settings.items():
                if key == "default_theme":
                    result = self.config_manager.update_config_value("gui_config", "theme", "default_theme", value)
                elif key == "md_default_template":
                    result = self.config_manager.update_config_value("gui_config", "template", "md_default_template", value)
                elif key == "transparency_enabled":
                    result = self.config_manager.update_config_value("gui_config", "transparency", "enabled", value)
                elif key == "default_transparency":
                    result = self.config_manager.update_config_value("gui_config", "transparency", "default_value", value)
                elif key in ["remember_gui_state", "auto_center", "default_mode"]:
                    result = self.config_manager.update_config_value("gui_config", "window", key, value)
                else:
                    logger.warning(f"未知的配置键: {key}")
                    continue
                
                if not result:
                    logger.error(f"更新配置失败: {key}")
                    success = False
            
            if success:
                logger.info("✓ 通用设置已成功应用到配置文件")
            else:
                logger.error("✗ 部分通用设置更新失败")
            
            return success
            
        except Exception as e:
            logger.error(f"应用通用设置失败: {e}", exc_info=True)
            return False
