"""
链接设置选项卡模块

实现设置对话框的链接设置选项卡，包含：
- Markdown链接格式设置（生成MD时的格式）
- 非嵌入链接处理（MD转文档/表格时）
- 嵌入链接处理（MD转文档/表格时）

注：路径解析和错误处理等高级设置请直接编辑link_config.toml文件
"""

import logging
import tkinter as tk
from typing import Dict, Any, Callable

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.gui.settings.base_tab import BaseSettingsTab
from gongwen_converter.gui.settings.config import SectionStyle
from gongwen_converter.utils.dpi_utils import scale

logger = logging.getLogger(__name__)

# 图片链接样式映射（配置值 <-> 显示值）
IMAGE_STYLE_OPTIONS = ["Markdown嵌入显示", "Markdown链接", "Wiki嵌入显示", "Wiki链接"]
IMAGE_STYLE_CONFIG_MAP = {
    "Markdown嵌入显示": "markdown_embed",
    "Markdown链接": "markdown_link",
    "Wiki嵌入显示": "wiki_embed",
    "Wiki链接": "wiki_link"
}
IMAGE_STYLE_DISPLAY_MAP = {v: k for k, v in IMAGE_STYLE_CONFIG_MAP.items()}

# MD文件链接样式映射（配置值 <-> 显示值）
MD_FILE_STYLE_OPTIONS = ["Markdown链接", "Wiki嵌入显示", "Wiki链接"]
MD_FILE_STYLE_CONFIG_MAP = {
    "Markdown链接": "markdown_link",
    "Wiki嵌入显示": "wiki_embed",
    "Wiki链接": "wiki_link"
}
MD_FILE_STYLE_DISPLAY_MAP = {v: k for k, v in MD_FILE_STYLE_CONFIG_MAP.items()}


class LinkTab(BaseSettingsTab):
    """
    链接设置选项卡类
    
    管理链接格式和处理相关的配置选项。
    所有配置对MD的生成和转换都有效（文档、表格、版式、图片OCR等）。
    """
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
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
            "Markdown链接格式",
            SectionStyle.PRIMARY
        )
        
        # 说明文本
        desc = tb.Label(
            frame,
            text="设置生成的Markdown文件中的链接格式",
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
        image_style = settings.get("image_link_style", "wiki_embed")
        image_style_display = IMAGE_STYLE_DISPLAY_MAP.get(image_style, "Wiki嵌入显示")
        self.image_link_style_var = tk.StringVar(value=image_style_display)
        self._create_column_combobox(
            left_column,
            "图片链接格式:",
            self.image_link_style_var,
            IMAGE_STYLE_OPTIONS,
            "设置图片在MD文件中的链接格式\n\n"
            "• Markdown嵌入显示: ![alt](image.png)\n"
            "• Markdown链接: [alt](image.png)\n"
            "• Wiki嵌入显示: ![[image.png]] (Obsidian)\n"
            "• Wiki链接: [[image.png]] (Obsidian)",
            self._on_image_link_style_changed
        )
        
        # 右列 - MD文件链接样式下拉框
        md_style = settings.get("md_file_link_style", "wiki_embed")
        md_style_display = MD_FILE_STYLE_DISPLAY_MAP.get(md_style, "Wiki嵌入显示")
        self.md_file_link_style_var = tk.StringVar(value=md_style_display)
        self._create_column_combobox(
            right_column,
            "MD文件链接格式:",
            self.md_file_link_style_var,
            MD_FILE_STYLE_OPTIONS,
            "设置MD文件在MD文件中的链接格式\n\n"
            "• Markdown链接: [file.md](file.md)\n"
            "• Wiki嵌入显示: ![[file.md]] (Obsidian)\n"
            "• Wiki链接: [[file.md]] (Obsidian)\n\n"
            "注：Markdown格式不支持嵌入MD文件内容",
            self._on_md_file_link_style_changed
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
            "非嵌入链接处理",
            SectionStyle.SUCCESS
        )
        
        # 说明文本
        desc = tb.Label(
            frame,
            text="设置Markdown转文档/表格时，如何处理普通链接（不带!前缀的链接）",
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
        wiki_mode = self.config_manager.get_wiki_link_mode()
        wiki_mode_display = self._mode_to_display(wiki_mode)
        self.wiki_link_mode_var = tk.StringVar(value=wiki_mode_display)
        self._create_column_combobox(
            left_column,
            "Wiki链接处理:",
            self.wiki_link_mode_var,
            ["保留原样", "提取文本", "完全移除"],
            "处理 [[link]] 或 [[link|text]] 格式\n\n"
            "保留原样: [[link|text]]\n"
            "提取文本: text（去除标记）\n"
            "完全移除: 删除整个链接",
            self._on_wiki_link_mode_changed
        )
        
        # 右列 - Markdown链接处理方式
        markdown_mode = self.config_manager.get_markdown_link_mode()
        markdown_mode_display = self._mode_to_display(markdown_mode)
        self.markdown_link_mode_var = tk.StringVar(value=markdown_mode_display)
        self._create_column_combobox(
            right_column,
            "Markdown链接处理:",
            self.markdown_link_mode_var,
            ["保留原样", "提取文本", "完全移除"],
            "处理 [text](url) 格式\n\n"
            "保留原样: [text](url)\n"
            "提取文本: text（去除标记）\n"
            "完全移除: 删除整个链接",
            self._on_markdown_link_mode_changed
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
            "嵌入链接处理",
            SectionStyle.WARNING
        )
        
        # 说明文本
        desc = tb.Label(
            frame,
            text="设置Markdown转文档/表格时，如何处理嵌入链接（带!前缀的链接）",
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
        wiki_image_mode = self.config_manager.get_wiki_embed_image_mode()
        wiki_image_mode_display = self._embed_mode_to_display(wiki_image_mode)
        self.wiki_embed_image_mode_var = tk.StringVar(value=wiki_image_mode_display)
        self._create_column_combobox(
            cell_00,
            "Wiki嵌入图片处理:",
            self.wiki_embed_image_mode_var,
            ["保留原样", "提取文本", "完全移除", "插入内容"],
            "处理 ![[image.png]] 格式\n\n"
            "保留原样: ![[image.png]]\n"
            "提取文本: image.png\n"
            "完全移除: 删除链接\n"
            "插入内容: 将图片插入到文档中",
            self._on_wiki_embed_image_mode_changed
        )
        
        # (0,1) Markdown嵌入图片处理方式
        md_image_mode = self.config_manager.get_markdown_embed_image_mode()
        md_image_mode_display = self._embed_mode_to_display(md_image_mode)
        self.markdown_embed_image_mode_var = tk.StringVar(value=md_image_mode_display)
        self._create_column_combobox(
            cell_01,
            "Markdown嵌入图片处理:",
            self.markdown_embed_image_mode_var,
            ["保留原样", "提取文本", "完全移除", "插入内容"],
            "处理 ![alt](image.png) 格式\n\n"
            "保留原样: ![alt](image.png)\n"
            "提取文本: image.png\n"
            "完全移除: 删除链接\n"
            "插入内容: 将图片插入到文档中",
            self._on_markdown_embed_image_mode_changed
        )
        
        # (1,0) 嵌入MD文件处理方式
        embed_md_mode = self.config_manager.get_embed_md_file_mode()
        embed_md_mode_display = self._embed_mode_to_display(embed_md_mode)
        self.embed_md_file_mode_var = tk.StringVar(value=embed_md_mode_display)
        self._create_column_combobox(
            cell_10,
            "Wiki嵌入的其他Markdown文件处理:",
            self.embed_md_file_mode_var,
            ["保留原样", "提取文本", "完全移除", "插入内容"],
            "处理 ![[file.md]] 格式（仅Wiki支持）\n\n"
            "保留原样: ![[file.md]]\n"
            "提取文本: file.md\n"
            "完全移除: 删除链接\n"
            "插入内容: 递归处理后插入",
            self._on_embed_md_file_mode_changed
        )
        
        # (1,1) 最大嵌入深度
        max_depth = self.config_manager.get_max_embed_depth()
        self.max_depth_var = tk.StringVar(value=str(max_depth))
        self._create_spinbox_in_column(
            cell_11,
            "最大嵌入深度:",
            self.max_depth_var,
            1, 10,
            "MD文件递归嵌入的最大深度\n"
            "防止无限递归和性能问题\n\n"
            "例如: A嵌入B，B嵌入C，C嵌入D...\n"
            "推荐值: 3-5",
            self._on_max_depth_changed
        )
        
        logger.debug("嵌入链接处理区域创建完成")
    
    def _create_column_combobox(
        self,
        parent: tk.Widget,
        label_text: str,
        variable: tk.StringVar,
        values: list,
        tooltip: str,
        command=None,
        disabled: bool = False
    ) -> tb.Frame:
        """
        在指定列中创建带标签和信息图标的下拉框
        
        参数:
            parent: 父组件（左列或右列）
            label_text: 标签文本
            variable: 绑定的StringVar变量
            values: 下拉选项列表
            tooltip: 工具提示文本
            command: 选择改变时的回调函数
            disabled: 是否禁用下拉框
            
        返回:
            tb.Frame: 包含标签和下拉框的容器框架
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
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
        
        # 创建下拉框
        combobox = tb.Combobox(
            container,
            textvariable=variable,
            values=values,
            state="disabled" if disabled else "readonly",
            bootstyle="secondary"
        )
        combobox.pack(fill="x")
        
        # 绑定事件
        if command and not disabled:
            combobox.bind("<<ComboboxSelected>>", command)
        
        return container
    
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
            parent: 父组件（左列或右列）
            label_text: 标签文本
            variable: 绑定的StringVar变量
            from_val: 最小值
            to_val: 最大值
            tooltip: 工具提示文本
            command: 值改变时的回调函数
            
        返回:
            tb.Frame: 包含标签和Spinbox的容器框架
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
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
    
    # === 辅助方法：显示值和配置值转换 ===
    
    def _mode_to_display(self, mode: str) -> str:
        """将配置模式转换为显示文本"""
        mapping = {
            "keep": "保留原样",
            "extract_text": "提取文本",
            "remove": "完全移除"
        }
        return mapping.get(mode, "提取文本")
    
    def _display_to_mode(self, display: str) -> str:
        """将显示文本转换为配置模式"""
        mapping = {
            "保留原样": "keep",
            "提取文本": "extract_text",
            "完全移除": "remove"
        }
        return mapping.get(display, "extract_text")
    
    def _embed_mode_to_display(self, mode: str) -> str:
        """将嵌入配置模式转换为显示文本"""
        mapping = {
            "keep": "保留原样",
            "extract_text": "提取文本",
            "remove": "完全移除",
            "embed": "插入内容"
        }
        return mapping.get(mode, "插入内容")
    
    def _display_to_embed_mode(self, display: str) -> str:
        """将显示文本转换为嵌入配置模式"""
        mapping = {
            "保留原样": "keep",
            "提取文本": "extract_text",
            "完全移除": "remove",
            "插入内容": "embed"
        }
        return mapping.get(display, "embed")
    
    # === 事件处理方法 ===
    
    def _on_image_link_style_changed(self, event=None):
        """处理图片链接样式变更"""
        display = self.image_link_style_var.get()
        config_value = IMAGE_STYLE_CONFIG_MAP.get(display, "wiki_embed")
        logger.info(f"图片链接样式变更: {display} → {config_value}")
        self.on_change("image_link_style", config_value)
    
    def _on_md_file_link_style_changed(self, event=None):
        """处理MD文件链接样式变更"""
        display = self.md_file_link_style_var.get()
        config_value = MD_FILE_STYLE_CONFIG_MAP.get(display, "wiki_embed")
        logger.info(f"MD文件链接样式变更: {display} → {config_value}")
        self.on_change("md_file_link_style", config_value)
    
    def _on_wiki_link_mode_changed(self, event=None):
        """处理Wiki链接处理方式变更"""
        display = self.wiki_link_mode_var.get()
        config_value = self._display_to_mode(display)
        logger.info(f"Wiki链接处理方式变更: {display} → {config_value}")
        self.on_change("wiki_mode", config_value)
    
    def _on_markdown_link_mode_changed(self, event=None):
        """处理Markdown链接处理方式变更"""
        display = self.markdown_link_mode_var.get()
        config_value = self._display_to_mode(display)
        logger.info(f"Markdown链接处理方式变更: {display} → {config_value}")
        self.on_change("markdown_mode", config_value)
    
    def _on_wiki_embed_image_mode_changed(self, event=None):
        """处理Wiki嵌入图片处理方式变更"""
        display = self.wiki_embed_image_mode_var.get()
        config_value = self._display_to_embed_mode(display)
        logger.info(f"Wiki嵌入图片处理方式变更: {display} → {config_value}")
        self.on_change("wiki_image_mode", config_value)
    
    def _on_markdown_embed_image_mode_changed(self, event=None):
        """处理Markdown嵌入图片处理方式变更"""
        display = self.markdown_embed_image_mode_var.get()
        config_value = self._display_to_embed_mode(display)
        logger.info(f"Markdown嵌入图片处理方式变更: {display} → {config_value}")
        self.on_change("markdown_image_mode", config_value)
    
    def _on_embed_md_file_mode_changed(self, event=None):
        """处理嵌入MD文件处理方式变更"""
        display = self.embed_md_file_mode_var.get()
        config_value = self._display_to_embed_mode(display)
        logger.info(f"嵌入MD文件处理方式变更: {display} → {config_value}")
        self.on_change("embed_md_file_mode", config_value)
    
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
    
    # === 配置获取和应用方法 ===
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置
        
        返回：
            Dict[str, Any]: 当前所有设置项的值
        """
        # 图片链接样式
        image_style_display = self.image_link_style_var.get()
        image_style_config = IMAGE_STYLE_CONFIG_MAP.get(image_style_display, "wiki_embed")
        
        # MD文件链接样式
        md_style_display = self.md_file_link_style_var.get()
        md_style_config = MD_FILE_STYLE_CONFIG_MAP.get(md_style_display, "wiki_embed")
        
        settings = {
            # Markdown链接格式
            "image_link_style": image_style_config,
            "md_file_link_style": md_style_config,
            # 非嵌入链接处理
            "wiki_mode": self._display_to_mode(self.wiki_link_mode_var.get()),
            "markdown_mode": self._display_to_mode(self.markdown_link_mode_var.get()),
            # 嵌入链接处理（拆分为Wiki和Markdown两种图片处理）
            "wiki_image_mode": self._display_to_embed_mode(self.wiki_embed_image_mode_var.get()),
            "markdown_image_mode": self._display_to_embed_mode(self.markdown_embed_image_mode_var.get()),
            "md_file_mode": self._display_to_embed_mode(self.embed_md_file_mode_var.get()),
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
