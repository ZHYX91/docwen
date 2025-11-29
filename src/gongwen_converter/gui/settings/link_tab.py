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

logger = logging.getLogger(__name__)


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
        创建Markdown链接格式设置区域
        
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
            text="设置生成MD时的链接格式（适用于文档、表格、版式、图片OCR等所有场景）",
            bootstyle="secondary",
            wraplength=450,
            justify="left"
        )
        desc.pack(anchor="w", pady=(0, 10))
        
        # 获取当前配置
        try:
            settings = self.config_manager.get_markdown_link_style_settings()
        except Exception as e:
            logger.warning(f"读取链接配置失败，使用默认值: {e}")
            settings = {
                "image_link_format": "wiki",
                "image_embed": True,
                "md_file_link_format": "wiki",
                "md_file_embed": True
            }
        
        # 图片链接格式
        image_format = settings.get("image_link_format", "wiki")
        image_format_display = "Wiki链接" if image_format == "wiki" else "Markdown链接"
        self.image_link_format_var = tk.StringVar(value=image_format_display)
        self.create_combobox_with_info(
            frame,
            "图片链接格式:",
            self.image_link_format_var,
            ["Markdown链接", "Wiki链接"],
            "Markdown格式: ![image.png](image.png) 或 [image.png](image.png)\n"
            "Wiki格式: ![[image.png]] 或 [[image.png]]（适用于Obsidian等）",
            self._on_image_link_format_changed
        )
        
        # 图片是否嵌入
        self.image_embed_var = tk.BooleanVar(value=settings.get("image_embed", True))
        self.create_checkbox_with_info(
            frame,
            "图片嵌入显示",
            self.image_embed_var,
            "启用: 显示图片内容（带!前缀）\n"
            "禁用: 仅作为链接（不带!前缀）\n\n"
            "Markdown: ![](url) vs [](url)\n"
            "Wiki: ![[url]] vs [[url]]",
            self._on_image_embed_changed
        )
        
        # MD文件链接格式
        md_format = settings.get("md_file_link_format", "wiki")
        md_format_display = "Wiki链接" if md_format == "wiki" else "Markdown链接"
        self.md_file_link_format_var = tk.StringVar(value=md_format_display)
        self.create_combobox_with_info(
            frame,
            "MD文件链接格式:",
            self.md_file_link_format_var,
            ["Markdown链接", "Wiki链接"],
            "Markdown格式: [file.md](file.md)（固定为链接，无嵌入模式）\n"
            "Wiki格式: ![[file.md]] 或 [[file.md]]（可嵌入或链接）",
            self._on_md_file_link_format_changed
        )
        
        # MD文件是否嵌入（仅Wiki有效）
        self.md_file_embed_var = tk.BooleanVar(value=settings.get("md_file_embed", True))
        self.create_checkbox_with_info(
            frame,
            "MD文件嵌入显示（仅Wiki）",
            self.md_file_embed_var,
            "启用: ![[file.md]] 嵌入文件内容\n"
            "禁用: [[file.md]] 仅作为链接\n\n"
            "⚠️ 仅当选择Wiki链接格式时有效",
            self._on_md_file_embed_changed
        )
        
        logger.debug("Markdown链接格式设置区域创建完成")
    
    def _create_non_embed_links_section(self):
        """
        创建非嵌入链接处理区域
        
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
            text="设置MD转文档/表格时如何处理普通链接（不带!前缀的链接）",
            bootstyle="secondary",
            wraplength=450,
            justify="left"
        )
        desc.pack(anchor="w", pady=(0, 10))
        
        # Wiki链接处理方式
        wiki_mode = self.config_manager.get_wiki_link_mode()
        wiki_mode_display = self._mode_to_display(wiki_mode)
        self.wiki_link_mode_var = tk.StringVar(value=wiki_mode_display)
        self.create_combobox_with_info(
            frame,
            "Wiki链接处理方式:",
            self.wiki_link_mode_var,
            ["保留原样", "提取文本", "完全移除"],
            "处理 [[link]] 或 [[link|text]] 格式的链接\n\n"
            "保留原样: [[link|text]]\n"
            "提取文本: text（去除链接标记）\n"
            "完全移除: 删除整个链接",
            self._on_wiki_link_mode_changed
        )
        
        # Markdown链接处理方式
        markdown_mode = self.config_manager.get_markdown_link_mode()
        markdown_mode_display = self._mode_to_display(markdown_mode)
        self.markdown_link_mode_var = tk.StringVar(value=markdown_mode_display)
        self.create_combobox_with_info(
            frame,
            "Markdown链接处理方式:",
            self.markdown_link_mode_var,
            ["保留原样", "提取文本", "完全移除"],
            "处理 [text](url) 格式的链接\n\n"
            "保留原样: [text](url)\n"
            "提取文本: text（去除链接标记）\n"
            "完全移除: 删除整个链接",
            self._on_markdown_link_mode_changed
        )
        
        logger.debug("非嵌入链接处理区域创建完成")
    
    def _create_embed_links_section(self):
        """
        创建嵌入链接处理区域
        
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
            text="设置MD转文档/表格时如何处理嵌入链接（带!前缀的链接）",
            bootstyle="secondary",
            wraplength=450,
            justify="left"
        )
        desc.pack(anchor="w", pady=(0, 10))
        
        # 嵌入功能总开关
        self.embedding_enabled_var = tk.BooleanVar(
            value=self.config_manager.is_embedding_enabled()
        )
        self.create_checkbox_with_info(
            frame,
            "启用嵌入功能",
            self.embedding_enabled_var,
            "启用: 根据下面的设置处理嵌入链接\n"
            "禁用: 嵌入链接按非嵌入方式处理（使用上面的设置）",
            self._on_embedding_enabled_changed
        )
        
        # 嵌入图片处理方式
        embed_image_mode = self.config_manager.get_embed_image_mode()
        embed_image_mode_display = self._embed_mode_to_display(embed_image_mode)
        self.embed_image_mode_var = tk.StringVar(value=embed_image_mode_display)
        self.create_combobox_with_info(
            frame,
            "嵌入图片处理方式:",
            self.embed_image_mode_var,
            ["保留原样", "提取文本", "完全移除", "插入内容"],
            "处理 ![[image.png]] 或 ![](image.png) 格式\n\n"
            "保留原样: ![[image.png]]\n"
            "提取文本: image.png\n"
            "完全移除: 删除链接\n"
            "插入内容: 将图片实际插入到文档中",
            self._on_embed_image_mode_changed
        )
        
        # 嵌入MD文件处理方式
        embed_md_mode = self.config_manager.get_embed_md_file_mode()
        embed_md_mode_display = self._embed_mode_to_display(embed_md_mode)
        self.embed_md_file_mode_var = tk.StringVar(value=embed_md_mode_display)
        self.create_combobox_with_info(
            frame,
            "嵌入MD文件处理方式:",
            self.embed_md_file_mode_var,
            ["保留原样", "提取文本", "完全移除", "插入内容"],
            "处理 ![[file.md]] 格式（仅Wiki支持）\n\n"
            "保留原样: ![[file.md]]\n"
            "提取文本: file.md\n"
            "完全移除: 删除链接\n"
            "插入内容: 读取文件内容并递归处理后插入",
            self._on_embed_md_file_mode_changed
        )
        
        # 最大嵌入深度
        max_depth = self.config_manager.get_max_embed_depth()
        self.max_depth_var = tk.StringVar(value=str(max_depth))
        
        depth_container = tb.Frame(frame)
        depth_container.pack(fill="x", pady=(0, self.layout_config.widget_spacing))
        
        depth_label_frame = tb.Frame(depth_container)
        depth_label_frame.pack(fill="x", pady=(0, self.layout_config.label_spacing))
        
        depth_label = tb.Label(
            depth_label_frame,
            text="最大嵌入深度:",
            bootstyle="secondary"
        )
        depth_label.pack(side="left")
        
        from gongwen_converter.utils.gui_utils import create_info_icon
        depth_info = create_info_icon(
            depth_label_frame,
            "MD文件递归嵌入的最大深度\n"
            "防止无限递归和性能问题\n\n"
            "例如: A嵌入B，B嵌入C，C嵌入D...\n"
            "推荐值: 3-5",
            "info"
        )
        depth_info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        depth_spinbox = tb.Spinbox(
            depth_container,
            from_=1,
            to=10,
            textvariable=self.max_depth_var,
            bootstyle="secondary",
            width=10,
            command=self._on_max_depth_changed
        )
        depth_spinbox.pack(anchor="w")
        # 绑定键盘事件
        depth_spinbox.bind('<Return>', lambda e: self._on_max_depth_changed())
        depth_spinbox.bind('<FocusOut>', lambda e: self._on_max_depth_changed())
        
        logger.debug("嵌入链接处理区域创建完成")
    
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
    
    def _on_image_link_format_changed(self, event=None):
        """处理图片链接格式变更"""
        display = self.image_link_format_var.get()
        config_value = "wiki" if display == "Wiki链接" else "markdown"
        logger.info(f"图片链接格式变更: {display} → {config_value}")
        self.on_change("image_link_format", config_value)
    
    def _on_image_embed_changed(self):
        """处理图片嵌入设置变更"""
        value = self.image_embed_var.get()
        logger.info(f"图片嵌入显示设置变更: {value}")
        self.on_change("image_embed", value)
    
    def _on_md_file_link_format_changed(self, event=None):
        """处理MD文件链接格式变更"""
        display = self.md_file_link_format_var.get()
        config_value = "wiki" if display == "Wiki链接" else "markdown"
        logger.info(f"MD文件链接格式变更: {display} → {config_value}")
        self.on_change("md_file_link_format", config_value)
    
    def _on_md_file_embed_changed(self):
        """处理MD文件嵌入设置变更"""
        value = self.md_file_embed_var.get()
        logger.info(f"MD文件嵌入显示设置变更: {value}")
        self.on_change("md_file_embed", value)
    
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
    
    def _on_embedding_enabled_changed(self):
        """处理嵌入功能开关变更"""
        value = self.embedding_enabled_var.get()
        logger.info(f"嵌入功能启用状态变更: {value}")
        self.on_change("embedding_enabled", value)
    
    def _on_embed_image_mode_changed(self, event=None):
        """处理嵌入图片处理方式变更"""
        display = self.embed_image_mode_var.get()
        config_value = self._display_to_embed_mode(display)
        logger.info(f"嵌入图片处理方式变更: {display} → {config_value}")
        self.on_change("embed_image_mode", config_value)
    
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
        # 图片链接格式
        image_format_display = self.image_link_format_var.get()
        image_format_config = "wiki" if image_format_display == "Wiki链接" else "markdown"
        
        # MD文件链接格式
        md_format_display = self.md_file_link_format_var.get()
        md_format_config = "wiki" if md_format_display == "Wiki链接" else "markdown"
        
        settings = {
            # Markdown链接格式
            "image_link_format": image_format_config,
            "image_embed": self.image_embed_var.get(),
            "md_file_link_format": md_format_config,
            "md_file_embed": self.md_file_embed_var.get(),
            # 非嵌入链接处理
            "wiki_mode": self._display_to_mode(self.wiki_link_mode_var.get()),
            "markdown_mode": self._display_to_mode(self.markdown_link_mode_var.get()),
            # 嵌入链接处理
            "embedding_enabled": self.embedding_enabled_var.get(),
            "embed_image_mode": self._display_to_embed_mode(self.embed_image_mode_var.get()),
            "embed_md_file_mode": self._display_to_embed_mode(self.embed_md_file_mode_var.get()),
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
                "link_config", "format", "image_link_format", settings["image_link_format"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "link_config", "format", "image_embed", settings["image_embed"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "link_config", "format", "md_file_link_format", settings["md_file_link_format"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "link_config", "format", "md_file_embed", settings["md_file_embed"]
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
                "link_config", "embed_links", "enabled", settings["embedding_enabled"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "link_config", "embed_links", "image_mode", settings["embed_image_mode"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "link_config", "embed_links", "md_file_mode", settings["embed_md_file_mode"]
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
