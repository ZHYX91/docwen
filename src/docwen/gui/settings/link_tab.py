"""
链接设置选项卡模块

实现设置对话框的链接设置选项卡，包含：
- Markdown链接格式设置（生成MD时的格式）
- 非嵌入链接处理（MD转文档/表格时）
- 嵌入链接处理（MD转文档/表格时）

注：路径解析和错误处理等高级设置请直接编辑link_config.toml文件

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。
使用 ConfigCombobox 组件实现配置值与显示文本的分离，
避免语言切换时的映射问题。
"""

from __future__ import annotations

import logging
import tkinter as tk
from typing import Dict, Any, Callable

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.gui.settings.base_tab import BaseSettingsTab
from docwen.gui.settings.config import SectionStyle
from docwen.gui.components.config_combobox import ConfigCombobox
from docwen.utils.dpi_utils import scale
from docwen.i18n import t

logger = logging.getLogger(__name__)

# 图片链接样式的配置值和翻译键映射
IMAGE_STYLE_CONFIG_VALUES = ["markdown_embed", "markdown_link", "wiki_embed", "wiki_link"]
IMAGE_STYLE_TRANSLATE_KEYS = {
    "markdown_embed": "settings.link.image_styles.md_embed",
    "markdown_link": "settings.link.image_styles.md_link",
    "wiki_embed": "settings.link.image_styles.wiki_embed",
    "wiki_link": "settings.link.image_styles.wiki_link"
}

# MD文件链接样式的配置值和翻译键映射
MD_FILE_STYLE_CONFIG_VALUES = ["markdown_link", "wiki_embed", "wiki_link"]
MD_FILE_STYLE_TRANSLATE_KEYS = {
    "markdown_link": "settings.link.md_file_styles.md_link",
    "wiki_embed": "settings.link.md_file_styles.wiki_embed",
    "wiki_link": "settings.link.md_file_styles.wiki_link"
}

# 链接处理模式的配置值和翻译键映射
LINK_MODE_CONFIG_VALUES = ["keep", "extract_text", "remove"]
LINK_MODE_TRANSLATE_KEYS = {
    "keep": "settings.link.modes.keep",
    "extract_text": "settings.link.modes.extract_text",
    "remove": "settings.link.modes.remove"
}

# 嵌入链接处理模式（多一个 embed 选项）
EMBED_MODE_CONFIG_VALUES = ["keep", "extract_text", "remove", "embed"]
EMBED_MODE_TRANSLATE_KEYS = {
    "keep": "settings.link.modes.keep",
    "extract_text": "settings.link.modes.extract_text",
    "remove": "settings.link.modes.remove",
    "embed": "settings.link.modes.embed"
}


class LinkTab(BaseSettingsTab):
    """
    链接设置选项卡类
    
    管理链接格式和处理相关的配置选项。
    所有配置对MD的生成和转换都有效（文档、表格、版式、图片OCR等）。
    """
    
    def __init__(self, parent, config_manager: Any, on_change: Callable[[str, Any], None]):
        """初始化链接设置选项卡"""
        super().__init__(parent, config_manager, on_change)
        logger.info("链接设置选项卡初始化完成")
    
    def _create_interface(self):
        """
        创建选项卡界面
        
        创建三个主要区域：
        1. Markdown链接格式
        2. 非嵌入链接处理
        3. 嵌入链接处理
        
        注：高级设置（路径解析、错误处理）需要直接编辑配置文件
        """
        logger.debug("开始创建链接设置选项卡界面")
        
        self._create_markdown_link_format_section()
        self._create_non_embed_links_section()
        self._create_embed_links_section()
        
        logger.debug("链接设置选项卡界面创建完成")
    
    def _create_markdown_link_format_section(self):
        """
        创建Markdown链接格式设置区域（两列布局）
        
        配置生成MD时的链接格式（适用于所有生成MD的场景）
        
        配置路径：link_config.format.*
        """
        logger.debug("创建Markdown链接格式设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.link.format_section"),
            SectionStyle.PRIMARY
        )
        
        # 说明文本
        desc = tb.Label(
            frame,
            text=t("settings.link.format_desc"),
            bootstyle="secondary",
            wraplength=scale(450)
        )
        desc.pack(anchor="w", pady=(0, 10))
        
        # 获取当前配置
        try:
            settings = self.config_manager.get_markdown_link_style_settings()
        except Exception as e:
            logger.warning(f"读取链接配置失败，使用默认值: {e}")
            settings = {
                "image_link_style": "wiki_embed",
                "md_file_link_style": "wiki_embed"
            }
        
        # 创建两列容器
        columns_frame = tb.Frame(frame)
        columns_frame.pack(fill="x")
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        
        left_column = tb.Frame(columns_frame)
        left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        right_column = tb.Frame(columns_frame)
        right_column.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        
        # 左列 - 图片链接样式下拉框
        self.image_link_style_combo = self._create_config_combobox(
            left_column,
            t("settings.link.image_link_style_label"),
            IMAGE_STYLE_CONFIG_VALUES,
            IMAGE_STYLE_TRANSLATE_KEYS,
            settings.get("image_link_style", "wiki_embed"),
            t("settings.link.image_link_style_tooltip"),
            lambda v: self.on_change("image_link_style", v)
        )
        
        # 右列 - MD文件链接样式下拉框
        self.md_file_link_style_combo = self._create_config_combobox(
            right_column,
            t("settings.link.md_file_link_style_label"),
            MD_FILE_STYLE_CONFIG_VALUES,
            MD_FILE_STYLE_TRANSLATE_KEYS,
            settings.get("md_file_link_style", "wiki_embed"),
            t("settings.link.md_file_link_style_tooltip"),
            lambda v: self.on_change("md_file_link_style", v)
        )
        
        logger.debug("Markdown链接格式设置区域创建完成")
    
    def _create_non_embed_links_section(self):
        """
        创建非嵌入链接处理区域（两列布局）
        
        配置MD转文档/表格时如何处理普通链接
        
        配置路径：link_config.non_embed_links.*
        """
        logger.debug("创建非嵌入链接处理区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.link.non_embed_section"),
            SectionStyle.SUCCESS
        )
        
        # 说明文本
        desc = tb.Label(
            frame,
            text=t("settings.link.non_embed_desc"),
            bootstyle="secondary",
            wraplength=scale(450)
        )
        desc.pack(anchor="w", pady=(0, 10))
        
        # 创建两列容器
        columns_frame = tb.Frame(frame)
        columns_frame.pack(fill="x")
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        
        left_column = tb.Frame(columns_frame)
        left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        right_column = tb.Frame(columns_frame)
        right_column.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        
        # 左列 - Wiki链接处理方式
        self.wiki_link_mode_combo = self._create_config_combobox(
            left_column,
            t("settings.link.wiki_link_mode"),
            LINK_MODE_CONFIG_VALUES,
            LINK_MODE_TRANSLATE_KEYS,
            self.config_manager.get_wiki_link_mode(),
            t("settings.link.wiki_link_tooltip"),
            lambda v: self.on_change("wiki_mode", v)
        )
        
        # 右列 - Markdown链接处理方式
        self.markdown_link_mode_combo = self._create_config_combobox(
            right_column,
            t("settings.link.markdown_link_mode"),
            LINK_MODE_CONFIG_VALUES,
            LINK_MODE_TRANSLATE_KEYS,
            self.config_manager.get_markdown_link_mode(),
            t("settings.link.markdown_link_tooltip"),
            lambda v: self.on_change("markdown_mode", v)
        )
        
        logger.debug("非嵌入链接处理区域创建完成")
    
    def _create_embed_links_section(self):
        """
        创建嵌入链接处理区域（混合布局）
        
        配置MD转文档/表格时如何处理嵌入链接
        
        配置路径：link_config.embed_links.* 和 link_config.embedding.*
        """
        logger.debug("创建嵌入链接处理区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.link.embed_section"),
            SectionStyle.WARNING
        )
        
        # 说明文本
        desc = tb.Label(
            frame,
            text=t("settings.link.embed_desc"),
            bootstyle="secondary",
            wraplength=scale(450)
        )
        desc.pack(anchor="w", pady=(0, 10))
        
        # === 统一的 2行×2列 容器 ===
        grid_frame = tb.Frame(frame)
        grid_frame.pack(fill="x")
        grid_frame.columnconfigure(0, weight=1, uniform="col")
        grid_frame.columnconfigure(1, weight=1, uniform="col")
        
        # 创建4个单元格容器
        cell_00 = tb.Frame(grid_frame)  # 第1行左列
        cell_00.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        cell_01 = tb.Frame(grid_frame)  # 第1行右列
        cell_01.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        
        cell_10 = tb.Frame(grid_frame)  # 第2行左列
        cell_10.grid(row=1, column=0, sticky="nsew", padx=(0, 5))
        
        cell_11 = tb.Frame(grid_frame)  # 第2行右列
        cell_11.grid(row=1, column=1, sticky="nsew", padx=(5, 0))
        
        # (0,0) Wiki嵌入图片处理方式
        self.wiki_embed_image_mode_combo = self._create_config_combobox(
            cell_00,
            t("settings.link.wiki_embed_image_mode"),
            EMBED_MODE_CONFIG_VALUES,
            EMBED_MODE_TRANSLATE_KEYS,
            self.config_manager.get_wiki_embed_image_mode(),
            t("settings.link.wiki_embed_image_tooltip"),
            lambda v: self.on_change("wiki_image_mode", v)
        )
        
        # (0,1) Markdown嵌入图片处理方式
        self.markdown_embed_image_mode_combo = self._create_config_combobox(
            cell_01,
            t("settings.link.markdown_embed_image_mode"),
            EMBED_MODE_CONFIG_VALUES,
            EMBED_MODE_TRANSLATE_KEYS,
            self.config_manager.get_markdown_embed_image_mode(),
            t("settings.link.markdown_embed_image_tooltip"),
            lambda v: self.on_change("markdown_image_mode", v)
        )
        
        # (1,0) 嵌入MD文件处理方式
        self.embed_md_file_mode_combo = self._create_config_combobox(
            cell_10,
            t("settings.link.embed_md_file_mode"),
            EMBED_MODE_CONFIG_VALUES,
            EMBED_MODE_TRANSLATE_KEYS,
            self.config_manager.get_embed_md_file_mode(),
            t("settings.link.embed_md_file_tooltip"),
            lambda v: self.on_change("md_file_mode", v)
        )
        
        # (1,1) 最大嵌入深度
        max_depth = self.config_manager.get_max_embed_depth()
        self.max_depth_var = tk.StringVar(value=str(max_depth))
        self._create_spinbox_in_column(
            cell_11,
            t("settings.link.max_depth_label"),
            self.max_depth_var,
            1, 10,
            t("settings.link.max_depth_tooltip"),
            self._on_max_depth_changed
        )
        
        logger.debug("嵌入链接处理区域创建完成")
    
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
        在指定列中创建带标签和信息图标的配置下拉框
        
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
    
    def _create_spinbox_in_column(
        self,
        parent: tk.Widget,
        label_text: str,
        variable: tk.StringVar,
        from_val: int,
        to_val: int,
        tooltip: str,
        command=None
    ) -> tb.Frame:
        """
        在指定列中创建带标签和信息图标的Spinbox
        
        参数:
            parent: 父组件
            label_text: 标签文本
            variable: 绑定的StringVar变量
            from_val: 最小值
            to_val: 最大值
            tooltip: 工具提示文本
            command: 值改变时的回调函数
            
        返回:
            tb.Frame: 包含标签和Spinbox的容器框架
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
        
        # 创建Spinbox
        spinbox = tb.Spinbox(
            container,
            from_=from_val,
            to=to_val,
            textvariable=variable,
            bootstyle="secondary",
            command=command
        )
        spinbox.pack(fill="x")
        
        # 绑定键盘事件
        if command:
            spinbox.bind('<Return>', lambda e: command())
            spinbox.bind('<FocusOut>', lambda e: command())
        
        return container
    
    def _on_max_depth_changed(self):
        """处理最大嵌入深度变更"""
        try:
            value = int(self.max_depth_var.get())
            # 限制在1-10之间
            value = max(1, min(10, value))
            self.max_depth_var.set(str(value))
            logger.info(f"最大嵌入深度变更: {value}")
            self.on_change("max_depth", value)
        except ValueError:
            logger.warning("无效的最大深度值，使用默认值3")
            self.max_depth_var.set("3")
            self.on_change("max_depth", 3)
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置
        
        返回：
            Dict[str, Any]: 当前所有设置项的值
        """
        settings = {
            # Markdown链接格式（直接获取配置值，无需映射）
            "image_link_style": self.image_link_style_combo.get_config_value(),
            "md_file_link_style": self.md_file_link_style_combo.get_config_value(),
            # 非嵌入链接处理
            "wiki_mode": self.wiki_link_mode_combo.get_config_value(),
            "markdown_mode": self.markdown_link_mode_combo.get_config_value(),
            # 嵌入链接处理
            "wiki_image_mode": self.wiki_embed_image_mode_combo.get_config_value(),
            "markdown_image_mode": self.markdown_embed_image_mode_combo.get_config_value(),
            "md_file_mode": self.embed_md_file_mode_combo.get_config_value(),
            "max_depth": int(self.max_depth_var.get())
        }
        
        logger.debug(f"获取链接设置: {settings}")
        return settings
    
    def apply_settings(self) -> bool:
        """
        应用当前设置到配置文件
        
        将所有设置项保存到对应的配置路径。
        
        返回：
            bool: 应用是否成功
        """
        logger.debug("开始应用链接设置到配置文件")
        
        try:
            settings = self.get_settings()
            success = True
            
            # === Markdown链接格式 (format) ===
            if not self.config_manager.update_config_value(
                "link_config", "format", "image_link_style", settings["image_link_style"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "link_config", "format", "md_file_link_style", settings["md_file_link_style"]
            ):
                success = False
            
            # === 非嵌入链接处理 (non_embed_links) ===
            if not self.config_manager.update_config_value(
                "link_config", "non_embed_links", "wiki_mode", settings["wiki_mode"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "link_config", "non_embed_links", "markdown_mode", settings["markdown_mode"]
            ):
                success = False
            
            # === 嵌入链接处理 (embed_links) ===
            if not self.config_manager.update_config_value(
                "link_config", "embed_links", "wiki_image_mode", settings["wiki_image_mode"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "link_config", "embed_links", "markdown_image_mode", settings["markdown_image_mode"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "link_config", "embed_links", "md_file_mode", settings["md_file_mode"]
            ):
                success = False
            
            # === 嵌入详细配置 (embedding) ===
            if not self.config_manager.update_config_value(
                "link_config", "embedding", "max_depth", settings["max_depth"]
            ):
                success = False
            
            if success:
                logger.info("✓ 链接设置已成功应用到配置文件")
            else:
                logger.error("✗ 部分链接设置更新失败")
            
            return success
            
        except Exception as e:
            logger.error(f"应用链接设置失败: {e}", exc_info=True)
            return False
