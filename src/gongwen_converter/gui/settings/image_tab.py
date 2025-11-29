"""
图片设置选项卡模块

实现设置对话框的图片设置选项卡，包含：
- 图片提取默认设置（文档、表格、版式、图片文件）
"""

import logging
import tkinter as tk
from typing import Dict, Any, Callable

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.gui.settings.base_tab import BaseSettingsTab
from gongwen_converter.gui.settings.config import SectionStyle

logger = logging.getLogger(__name__)


class ImageTab(BaseSettingsTab):
    """
    图片设置选项卡类
    
    管理图片提取相关的配置选项。
    包含各种文件类型转MD时的图片提取默认设置。
    """
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
        """初始化图片设置选项卡"""
        super().__init__(parent, config_manager, on_change)
        logger.info("图片设置选项卡初始化完成")
    
    def _create_interface(self):
        """
        创建选项卡界面
        
        创建图片提取默认设置区域
        """
        logger.debug("开始创建图片设置选项卡界面")
        
        self._create_extraction_defaults_section()
        
        logger.debug("图片设置选项卡界面创建完成")
    
    def _create_extraction_defaults_section(self):
        """
        创建图片提取默认设置区域
        
        配置各种文件类型转MD时的图片提取默认选项。
        
        配置路径：image_config.extraction_defaults.*
        """
        logger.debug("创建图片提取默认设置区域")
        
        # 获取当前配置
        try:
            docx_keep = self.config_manager.get_docx_to_md_keep_images()
            docx_ocr = self.config_manager.get_docx_to_md_enable_ocr()
            xlsx_keep = self.config_manager.get_xlsx_to_md_keep_images()
            xlsx_ocr = self.config_manager.get_xlsx_to_md_enable_ocr()
            layout_keep = self.config_manager.get_layout_to_md_keep_images()
            layout_ocr = self.config_manager.get_layout_to_md_enable_ocr()
            image_keep = self.config_manager.get_image_to_md_keep_images()
            image_ocr = self.config_manager.get_image_to_md_enable_ocr()
        except Exception as e:
            logger.warning(f"读取图片配置失败，使用默认值: {e}")
            docx_keep, docx_ocr = True, False
            xlsx_keep, xlsx_ocr = True, False
            layout_keep, layout_ocr = True, False
            image_keep, image_ocr = True, True
        
        # 文档转MD设置
        doc_frame = self.create_section_frame(
            self.scrollable_frame,
            "文档文件 (DOCX/DOC)",
            SectionStyle.PRIMARY
        )
        
        self.docx_keep_var = tk.BooleanVar(value=docx_keep)
        self.create_checkbox_with_info(
            doc_frame,
            '默认勾选"提取图片"',
            self.docx_keep_var,
            '启用：DOCX转MD时默认勾选"提取图片"\n'
            '禁用：DOCX转MD时默认不勾选"提取图片"',
            self._on_docx_keep_changed
        )
        
        self.docx_ocr_var = tk.BooleanVar(value=docx_ocr)
        self.create_checkbox_with_info(
            doc_frame,
            '默认勾选"图片文字识别"',
            self.docx_ocr_var,
            '启用：DOCX转MD时默认勾选"图片文字识别"\n'
            '禁用：DOCX转MD时默认不勾选"图片文字识别"',
            self._on_docx_ocr_changed
        )
        
        # 表格转MD设置
        xlsx_frame = self.create_section_frame(
            self.scrollable_frame,
            "表格文件 (XLSX/XLS/CSV)",
            SectionStyle.SUCCESS
        )
        
        self.xlsx_keep_var = tk.BooleanVar(value=xlsx_keep)
        self.create_checkbox_with_info(
            xlsx_frame,
            '默认勾选"提取图片"',
            self.xlsx_keep_var,
            '启用：XLSX转MD时默认勾选"提取图片"\n'
            '禁用：XLSX转MD时默认不勾选"提取图片"',
            self._on_xlsx_keep_changed
        )
        
        self.xlsx_ocr_var = tk.BooleanVar(value=xlsx_ocr)
        self.create_checkbox_with_info(
            xlsx_frame,
            '默认勾选"图片文字识别"',
            self.xlsx_ocr_var,
            '启用：XLSX转MD时默认勾选"图片文字识别"\n'
            '禁用：XLSX转MD时默认不勾选"图片文字识别"',
            self._on_xlsx_ocr_changed
        )
        
        # 版式文件转MD设置
        layout_frame = self.create_section_frame(
            self.scrollable_frame,
            "版式文件 (PDF/OFD)",
            SectionStyle.WARNING
        )
        
        self.layout_keep_var = tk.BooleanVar(value=layout_keep)
        self.create_checkbox_with_info(
            layout_frame,
            '默认勾选"提取图片"',
            self.layout_keep_var,
            '启用：PDF/OFD转MD时默认勾选"提取图片"\n'
            '禁用：PDF/OFD转MD时默认不勾选"提取图片"',
            self._on_layout_keep_changed
        )
        
        self.layout_ocr_var = tk.BooleanVar(value=layout_ocr)
        self.create_checkbox_with_info(
            layout_frame,
            '默认勾选"图片文字识别"',
            self.layout_ocr_var,
            '启用：PDF/OFD转MD时默认勾选"图片文字识别"\n'
            '禁用：PDF/OFD转MD时默认不勾选"图片文字识别"',
            self._on_layout_ocr_changed
        )
        
        # 图片文件转MD设置
        image_frame = self.create_section_frame(
            self.scrollable_frame,
            "图片文件 (JPG/PNG)",
            SectionStyle.DANGER
        )
        
        self.image_keep_var = tk.BooleanVar(value=image_keep)
        self.create_checkbox_with_info(
            image_frame,
            '默认勾选"提取图片"',
            self.image_keep_var,
            '启用：图片转MD时默认勾选"提取图片"\n'
            '禁用：图片转MD时默认不勾选"提取图片"',
            self._on_image_keep_changed
        )
        
        self.image_ocr_var = tk.BooleanVar(value=image_ocr)
        self.create_checkbox_with_info(
            image_frame,
            '默认勾选"图片文字识别"',
            self.image_ocr_var,
            '启用：图片转MD时默认勾选"图片文字识别"\n'
            '禁用：图片转MD时默认不勾选"图片文字识别"',
            self._on_image_ocr_changed
        )
        
        logger.debug("图片提取默认设置区域创建完成")
    
    def _on_docx_keep_changed(self):
        """处理文档提取图片设置变更"""
        value = self.docx_keep_var.get()
        logger.info(f"文档提取图片默认设置变更: {value}")
        self.on_change("docx_to_md_keep_images", value)
    
    def _on_docx_ocr_changed(self):
        """处理文档OCR设置变更"""
        value = self.docx_ocr_var.get()
        logger.info(f"文档OCR默认设置变更: {value}")
        self.on_change("docx_to_md_enable_ocr", value)
    
    def _on_xlsx_keep_changed(self):
        """处理表格提取图片设置变更"""
        value = self.xlsx_keep_var.get()
        logger.info(f"表格提取图片默认设置变更: {value}")
        self.on_change("xlsx_to_md_keep_images", value)
    
    def _on_xlsx_ocr_changed(self):
        """处理表格OCR设置变更"""
        value = self.xlsx_ocr_var.get()
        logger.info(f"表格OCR默认设置变更: {value}")
        self.on_change("xlsx_to_md_enable_ocr", value)
    
    def _on_layout_keep_changed(self):
        """处理版式提取图片设置变更"""
        value = self.layout_keep_var.get()
        logger.info(f"版式提取图片默认设置变更: {value}")
        self.on_change("layout_to_md_keep_images", value)
    
    def _on_layout_ocr_changed(self):
        """处理版式OCR设置变更"""
        value = self.layout_ocr_var.get()
        logger.info(f"版式OCR默认设置变更: {value}")
        self.on_change("layout_to_md_enable_ocr", value)
    
    def _on_image_keep_changed(self):
        """处理图片提取图片设置变更"""
        value = self.image_keep_var.get()
        logger.info(f"图片提取图片默认设置变更: {value}")
        self.on_change("image_to_md_keep_images", value)
    
    def _on_image_ocr_changed(self):
        """处理图片OCR设置变更"""
        value = self.image_ocr_var.get()
        logger.info(f"图片OCR默认设置变更: {value}")
        self.on_change("image_to_md_enable_ocr", value)
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置
        
        返回：
            Dict[str, Any]: 当前所有设置项的值
        """
        settings = {
            "docx_to_md_keep_images": self.docx_keep_var.get(),
            "docx_to_md_enable_ocr": self.docx_ocr_var.get(),
            "xlsx_to_md_keep_images": self.xlsx_keep_var.get(),
            "xlsx_to_md_enable_ocr": self.xlsx_ocr_var.get(),
            "layout_to_md_keep_images": self.layout_keep_var.get(),
            "layout_to_md_enable_ocr": self.layout_ocr_var.get(),
            "image_to_md_keep_images": self.image_keep_var.get(),
            "image_to_md_enable_ocr": self.image_ocr_var.get()
        }
        
        logger.debug(f"获取图片设置: {settings}")
        return settings
    
    def apply_settings(self) -> bool:
        """
        应用当前设置到配置文件
        
        将所有设置项保存到对应的配置路径。
        
        返回：
            bool: 应用是否成功
        """
        logger.debug("开始应用图片设置到配置文件")
        
        try:
            settings = self.get_settings()
            success = True
            
            # 文档转MD设置
            if not self.config_manager.update_config_value(
                "image_config",
                "extraction_defaults",
                "docx_to_md_keep_images",
                settings["docx_to_md_keep_images"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "image_config",
                "extraction_defaults",
                "docx_to_md_enable_ocr",
                settings["docx_to_md_enable_ocr"]
            ):
                success = False
            
            # 表格转MD设置
            if not self.config_manager.update_config_value(
                "image_config",
                "extraction_defaults",
                "xlsx_to_md_keep_images",
                settings["xlsx_to_md_keep_images"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "image_config",
                "extraction_defaults",
                "xlsx_to_md_enable_ocr",
                settings["xlsx_to_md_enable_ocr"]
            ):
                success = False
            
            # 版式转MD设置
            if not self.config_manager.update_config_value(
                "image_config",
                "extraction_defaults",
                "layout_to_md_keep_images",
                settings["layout_to_md_keep_images"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "image_config",
                "extraction_defaults",
                "layout_to_md_enable_ocr",
                settings["layout_to_md_enable_ocr"]
            ):
                success = False
            
            # 图片转MD设置
            if not self.config_manager.update_config_value(
                "image_config",
                "extraction_defaults",
                "image_to_md_keep_images",
                settings["image_to_md_keep_images"]
            ):
                success = False
            
            if not self.config_manager.update_config_value(
                "image_config",
                "extraction_defaults",
                "image_to_md_enable_ocr",
                settings["image_to_md_enable_ocr"]
            ):
                success = False
            
            if success:
                logger.info("✓ 图片设置已成功应用到配置文件")
            else:
                logger.error("✗ 部分图片设置更新失败")
            
            return success
            
        except Exception as e:
            logger.error(f"应用图片设置失败: {e}", exc_info=True)
            return False
