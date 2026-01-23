"""
通用设置选项卡模块

实现设置对话框的通用设置选项卡，包含：
- 语言和地区设置
- 界面主题选择
- 窗口透明度设置
- 窗口行为配置

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。语言切换后需要重启应用才能生效。
使用 ConfigCombobox 组件实现配置值与显示文本的分离，
避免语言切换时的映射问题。
"""

import logging
import tkinter as tk
from typing import Dict, Any, Callable

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.gui.settings.base_tab import BaseSettingsTab
from docwen.gui.settings.config import SectionStyle
from docwen.gui.components.config_combobox import ConfigCombobox
from docwen.i18n import t, get_available_locales, get_current_locale, set_locale

logger = logging.getLogger(__name__)

# ========== 配置值常量和翻译键映射 ==========

# 主题选项（18个主题）
THEME_CONFIG_VALUES = [
    "cerculean", "cosmo", "cyborg", "darkly", "flatly", "journal",
    "litera", "lumen", "minty", "morph", "pulse", "sandstone",
    "simplex", "solar", "superhero", "united", "vapor", "yeti"
]
THEME_TRANSLATE_KEYS = {theme: f"settings.general.themes.{theme}" for theme in THEME_CONFIG_VALUES}

# 默认模式选项
DEFAULT_MODE_CONFIG_VALUES = ["single", "batch"]
DEFAULT_MODE_TRANSLATE_KEYS = {
    "single": "settings.general.window.single_mode",
    "batch": "settings.general.window.batch_mode"
}


class GeneralTab(BaseSettingsTab):
    """
    通用设置选项卡类
    
    管理应用程序的通用界面和窗口行为设置。
    包含语言、主题、透明度和窗口行为四个主要区域。
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
        
        创建四个主要设置区域：
        1. 语言设置 - 界面语言选择
        2. 主题设置 - 界面主题选择和预览
        3. 透明度设置 - 窗口透明效果配置
        4. 窗口行为 - 窗口位置、大小和默认模式
        """
        logger.debug("开始创建通用设置选项卡界面")
        
        self._create_language_section()
        self._create_theme_section()
        self._create_transparency_section()
        self._create_window_behavior_section()
        
        logger.debug("通用设置选项卡界面创建完成")
    
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
    
    def _create_language_section(self):
        """
        创建语言设置区域
        
        包含语言选择下拉框和重启提示标签。
        语言切换后需要重启应用才能生效。
        
        注意：语言选择使用 i18n 模块的动态数据，不使用 ConfigCombobox。
        
        配置路径：gui_config.language.locale
        """
        from docwen.utils.gui_utils import create_info_icon
        
        logger.debug("创建语言设置区域")
        
        # 创建区域框架
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.general.language_section"),
            SectionStyle.SUCCESS
        )
        
        # 创建标签行
        label_frame = tb.Frame(frame)
        label_frame.pack(fill="x", pady=(0, self.layout_config.label_spacing))
        
        label = tb.Label(label_frame, text=t("settings.general.language_label"), bootstyle="secondary")
        label.pack(side="left")
        
        info = create_info_icon(
            label_frame,
            t("settings.general.language_tooltip"),
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 获取可用语言列表
        available_locales = get_available_locales()
        locale_names = [loc["native_name"] for loc in available_locales]
        
        # 建立语言名称到代码的映射
        self.locale_name_to_code = {loc["native_name"]: loc["code"] for loc in available_locales}
        self.locale_code_to_name = {loc["code"]: loc["native_name"] for loc in available_locales}
        
        # 获取当前语言并转换为显示名称
        current_locale = get_current_locale()
        current_display = self.locale_code_to_name.get(current_locale, locale_names[0])
        
        # 创建语言下拉框（已启用）
        self.language_var = tk.StringVar(value=current_display)
        combobox = tb.Combobox(
            frame,
            textvariable=self.language_var,
            values=locale_names,
            state="readonly",
            bootstyle="secondary"
        )
        combobox.pack(fill="x")
        combobox.bind("<<ComboboxSelected>>", self._on_language_changed)
        
        # 添加重启提示
        hint_label = tb.Label(
            frame,
            text=t("settings.general.language_restart_hint"),
            bootstyle="warning",
            font=(None, 9)
        )
        hint_label.pack(anchor="w", pady=(5, 0))
        
        logger.debug("语言设置区域创建完成")
    
    def _on_language_changed(self, event=None):
        """
        处理语言变更事件
        
        记录语言变更，实际保存在 apply_settings 时执行。
        语言切换后需要重启应用才能生效。
        """
        new_language_name = self.language_var.get()
        new_locale_code = self.locale_name_to_code.get(new_language_name, "zh_CN")
        
        logger.info(f"语言已选择: {new_language_name} -> {new_locale_code}（将在保存并重启后生效）")
        self.on_change("locale", new_locale_code)
    
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
            t("settings.general.theme_section"),
            SectionStyle.PRIMARY
        )
        
        # 获取当前主题
        current_theme = self.config_manager.get_default_theme()
        
        # 创建主题选择下拉框（使用 ConfigCombobox）
        self.theme_combo = self._create_config_combobox(
            frame,
            t("settings.general.theme_label"),
            THEME_CONFIG_VALUES,
            THEME_TRANSLATE_KEYS,
            current_theme,
            t("settings.general.theme_tooltip"),
            self._on_theme_changed
        )
        
        # 主题预览区域
        preview_label = tb.Label(
            frame,
            text=t("settings.general.theme_preview"),
            bootstyle="secondary"
        )
        preview_label.pack(anchor="w", pady=(10, 5))
        
        preview_frame = tb.Frame(frame, bootstyle="default", height=60)
        preview_frame.pack(fill="x", pady=(0, 5))
        preview_frame.pack_propagate(True)
        
        # 添加预览控件
        self._create_theme_preview(preview_frame)
        
        logger.debug("主题设置区域创建完成")
    
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
            text=t("settings.general.sample_button"),
            bootstyle="primary",
            width=14
        )
        sample_button.pack(side="left", padx=10, pady=10)
        
        # 示例标签
        sample_label = tb.Label(
            parent,
            text=t("settings.general.sample_text"),
            bootstyle="secondary"
        )
        sample_label.pack(side="right", padx=10, pady=10)
        
        logger.debug("主题预览创建完成")
    
    def _on_theme_changed(self, new_theme: str):
        """
        处理主题变更事件
        
        主题变更会立即应用，提供实时预览效果。
        
        参数：
            new_theme: 新选择的主题配置值
        """
        logger.info(f"主题变更: {new_theme}（即时生效）")
        self.on_change("default_theme", new_theme)
        
        # 即时应用主题（实时预览）
        try:
            self.style.theme_use(new_theme)
            logger.debug(f"主题已切换为: {new_theme}")
        except Exception as e:
            logger.error(f"切换主题失败: {e}")
    
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
            t("settings.general.transparency.section"),
            SectionStyle.INFO
        )
        
        # 透明度启用开关
        self.transparency_enabled_var = tk.BooleanVar(
            value=self.config_manager.is_transparency_enabled()
        )
        self.create_checkbox_with_info(
            frame,
            t("settings.general.transparency.enable"),
            self.transparency_enabled_var,
            t("settings.general.transparency.tooltip"),
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
        
        # 透明度值标签（宽度设为5以确保能显示100%）
        transparency_percent = int(transparency_value * 100)
        self.transparency_value_label = tb.Label(
            slider_frame,
            text=f"{transparency_percent}%",
            width=5,
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
            t("settings.general.window.section"),
            SectionStyle.WARNING
        )
        
        # 记住窗口位置和大小
        self.remember_state_var = tk.BooleanVar(
            value=self.config_manager.should_remember_gui_state()
        )
        self.create_checkbox_with_info(
            frame,
            t("settings.general.window.remember_state"),
            self.remember_state_var,
            t("settings.general.window.remember_state_tooltip"),
            self._on_remember_state_changed
        )
        
        # 启动时自动居中窗口
        self.auto_center_var = tk.BooleanVar(
            value=self.config_manager.should_auto_center()
        )
        self.create_checkbox_with_info(
            frame,
            t("settings.general.window.auto_center"),
            self.auto_center_var,
            t("settings.general.window.auto_center_tooltip"),
            self._on_auto_center_changed
        )
        
        # 默认文件处理模式（使用 ConfigCombobox）
        current_mode = self.config_manager.get_default_mode()
        
        self.default_mode_combo = self._create_config_combobox(
            frame,
            t("settings.general.window.default_mode_label"),
            DEFAULT_MODE_CONFIG_VALUES,
            DEFAULT_MODE_TRANSLATE_KEYS,
            current_mode,
            t("settings.general.window.default_mode_tooltip"),
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
    
    def _on_default_mode_changed(self, new_mode: str):
        """
        处理默认模式变更事件
        
        参数：
            new_mode: 新选择的模式配置值
        """
        logger.info(f"默认模式变更: {new_mode}")
        self.on_change("default_mode", new_mode)
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置
        
        返回：
            Dict[str, Any]: 当前所有设置项的值
        """
        # 转换语言名称为代码
        language_display = self.language_var.get()
        locale_code = self.locale_name_to_code.get(language_display, "zh_CN")
        
        settings = {
            "locale": locale_code,
            "default_theme": self.theme_combo.get_config_value(),
            "transparency_enabled": self.transparency_enabled_var.get(),
            "default_transparency": self.transparency_var.get(),
            "remember_gui_state": self.remember_state_var.get(),
            "auto_center": self.auto_center_var.get(),
            "default_mode": self.default_mode_combo.get_config_value()
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
                if key == "locale":
                    # 使用 set_locale 保存语言设置
                    result = set_locale(value)
                elif key == "default_theme":
                    result = self.config_manager.update_config_value("gui_config", "theme", "default_theme", value)
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
