"""
图片设置选项卡模块

实现设置对话框的图片设置选项卡，包含：
- 提取/OCR默认设置
- 压缩选项默认设置
- PDF/TIFF选项默认设置
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
    
    管理图片文件相关的所有配置选项。
    包含提取/OCR设置、压缩选项和PDF/TIFF选项。
    """
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
        """初始化图片设置选项卡"""
        # 联动逻辑标志位和状态记录
        self._updating_image_options: bool = False
        self._image_last_image_state: bool = False
        self._image_last_ocr_state: bool = False
        
        super().__init__(parent, config_manager, on_change)
        logger.info("图片设置选项卡初始化完成")
    
    def _create_interface(self):
        """创建选项卡界面"""
        logger.debug("开始创建图片设置选项卡界面")
        
        self._create_extraction_section()
        self._create_compress_options_section()
        self._create_pdf_tiff_options_section()
        
        logger.debug("图片设置选项卡界面创建完成")
    
    def _create_extraction_section(self):
        """创建提取/OCR设置区域"""
        logger.debug("创建提取/OCR设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "图片提取/OCR设置",
            SectionStyle.DANGER
        )
        
        # 获取当前配置
        try:
            image_keep = self.config_manager.get_image_to_md_keep_images()
            image_ocr = self.config_manager.get_image_to_md_enable_ocr()
        except Exception as e:
            logger.warning(f"读取图片配置失败: {e}")
            image_keep, image_ocr = True, True
        
        # 提取图片
        self.image_keep_var = tk.BooleanVar(value=image_keep)
        self.create_checkbox_with_info(
            frame,
            '默认勾选"提取图片"',
            self.image_keep_var,
            '图片转MD时默认勾选"提取图片"',
            self._on_image_keep_changed
        )
        
        # OCR识别
        self.image_ocr_var = tk.BooleanVar(value=image_ocr)
        self.create_checkbox_with_info(
            frame,
            '默认勾选"图片文字识别"',
            self.image_ocr_var,
            '图片转MD时默认勾选"图片文字识别"',
            self._on_image_ocr_changed
        )
        
        # 初始化状态记录
        self._image_last_image_state = image_keep
        self._image_last_ocr_state = image_ocr
    
    def _create_compress_options_section(self):
        """创建压缩选项设置区域"""
        logger.debug("创建压缩选项设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "压缩选项设置",
            SectionStyle.INFO
        )
        
        # 获取当前配置
        try:
            compress_mode = self.config_manager.get_image_compress_mode()
            size_limit = self.config_manager.get_image_size_limit()
            size_unit = self.config_manager.get_image_size_unit()
        except Exception as e:
            logger.warning(f"读取压缩配置失败: {e}")
            compress_mode, size_limit, size_unit = "lossless", 200, "KB"
        
        # 说明文本
        desc_label = tb.Label(
            frame,
            text="设置图片格式转换时的默认压缩模式",
            bootstyle="secondary",
            wraplength=400
        )
        desc_label.pack(anchor="w", pady=(0, 10))
        
        # 压缩模式单选按钮
        self.compress_mode_var = tk.StringVar(value=compress_mode)
        
        radio_frame = tb.Frame(frame)
        radio_frame.pack(fill="x", pady=(10, 0))
        
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 最高质量
        lossless_frame = tb.Frame(radio_frame)
        lossless_frame.pack(fill="x", pady=(0, 8))
        
        tb.Radiobutton(
            lossless_frame,
            text="最高质量",
            variable=self.compress_mode_var,
            value="lossless",
            command=self._on_compress_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            lossless_frame,
            "无损或高质量转换（quality=95）",
            "info"
        ).pack(side="left", padx=(5, 0))
        
        # 限制大小
        limit_frame = tb.Frame(radio_frame)
        limit_frame.pack(fill="x")
        
        tb.Radiobutton(
            limit_frame,
            text="限制文件大小",
            variable=self.compress_mode_var,
            value="limit_size",
            command=self._on_compress_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            limit_frame,
            "通过调整quality参数压缩至目标大小",
            "info"
        ).pack(side="left", padx=(5, 0))
        
        # 大小限制输入
        input_frame = tb.Frame(frame)
        input_frame.pack(fill="x", pady=(10, 0))
        
        tb.Label(
            input_frame,
            text="默认大小上限：",
            font=(self.small_font, self.small_size),
            bootstyle="secondary"
        ).pack(side="left")
        
        self.size_limit_var = tk.IntVar(value=size_limit)
        tb.Entry(
            input_frame,
            textvariable=self.size_limit_var,
            width=10
        ).pack(side="left", padx=(5, 5))
        
        self.size_unit_var = tk.StringVar(value=size_unit)
        tb.Combobox(
            input_frame,
            textvariable=self.size_unit_var,
            values=["KB", "MB"],
            state="readonly",
            width=5
        ).pack(side="left")
    
    def _create_pdf_tiff_options_section(self):
        """创建PDF/TIFF选项设置区域"""
        logger.debug("创建PDF/TIFF选项设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "PDF/TIFF选项设置",
            SectionStyle.WARNING
        )
        
        # 获取当前配置
        try:
            pdf_quality = self.config_manager.get_image_pdf_quality()
            tiff_mode = self.config_manager.get_image_tiff_mode()
        except Exception as e:
            logger.warning(f"读取PDF/TIFF配置失败: {e}")
            pdf_quality, tiff_mode = "original", "smart"
        
        # PDF质量设置
        pdf_label = tb.Label(
            frame,
            text="图片转PDF默认尺寸",
            font=(self.small_font, self.small_size, "bold")
        )
        pdf_label.pack(anchor="w", pady=(0, 10))
        
        self.pdf_quality_var = tk.StringVar(value=pdf_quality)
        
        pdf_radio_frame = tb.Frame(frame)
        pdf_radio_frame.pack(fill="x", pady=(0, 15))
        
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 原图嵌入
        tb.Radiobutton(
            pdf_radio_frame,
            text="原图嵌入",
            variable=self.pdf_quality_var,
            value="original",
            command=self._on_pdf_quality_changed,
            bootstyle="primary"
        ).pack(side="left", padx=(0, 15))
        
        # 适合A4
        tb.Radiobutton(
            pdf_radio_frame,
            text="适合A4",
            variable=self.pdf_quality_var,
            value="a4",
            command=self._on_pdf_quality_changed,
            bootstyle="primary"
        ).pack(side="left", padx=(0, 15))
        
        # 适合A3
        tb.Radiobutton(
            pdf_radio_frame,
            text="适合A3",
            variable=self.pdf_quality_var,
            value="a3",
            command=self._on_pdf_quality_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        # TIFF模式设置
        tiff_label = tb.Label(
            frame,
            text="合并为TIFF默认模式",
            font=(self.small_font, self.small_size, "bold")
        )
        tiff_label.pack(anchor="w", pady=(0, 10))
        
        self.tiff_mode_var = tk.StringVar(value=tiff_mode)
        
        tiff_radio_frame = tb.Frame(frame)
        tiff_radio_frame.pack(fill="x")
        
        # 保留透明
        smart_frame = tb.Frame(tiff_radio_frame)
        smart_frame.pack(fill="x", pady=(0, 8))
        
        tb.Radiobutton(
            smart_frame,
            text="保留透明",
            variable=self.tiff_mode_var,
            value="smart",
            command=self._on_tiff_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            smart_frame,
            "保留源图片的透明背景（如果有）",
            "info"
        ).pack(side="left", padx=(5, 0))
        
        # 不保留透明
        rgb_frame = tb.Frame(tiff_radio_frame)
        rgb_frame.pack(fill="x")
        
        tb.Radiobutton(
            rgb_frame,
            text="不保留透明",
            variable=self.tiff_mode_var,
            value="RGB",
            command=self._on_tiff_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            rgb_frame,
            "转为RGB模式，透明背景变为白色",
            "info"
        ).pack(side="left", padx=(5, 0))
    
    # ========== 事件处理 ==========
    
    def _on_image_keep_changed(self):
        """
        处理提取图片设置变更
        
        实现联动逻辑：
        1. 勾选OCR时，自动勾选"提取图片"
        2. 取消"提取图片"时，自动取消OCR
        """
        if self._updating_image_options:
            return
        
        try:
            self._updating_image_options = True
            
            extract_image = self.image_keep_var.get()
            extract_ocr = self.image_ocr_var.get()
            
            # 检测状态变化
            image_changed = (extract_image != self._image_last_image_state)
            
            # 场景：用户取消提取图片（从有到无）
            if not extract_image and self._image_last_image_state and image_changed:
                if extract_ocr:
                    logger.debug("图片设置：取消提取图片，自动取消OCR")
                    self.image_ocr_var.set(False)
            
            # 更新状态记录
            self._image_last_image_state = self.image_keep_var.get()
            self._image_last_ocr_state = self.image_ocr_var.get()
            
            # 通知配置变更
            value = self.image_keep_var.get()
            logger.info(f"图片提取图片设置变更: {value}")
            self.on_change("to_md_keep_images", value)
        
        finally:
            self._updating_image_options = False
    
    def _on_image_ocr_changed(self):
        """
        处理OCR设置变更
        
        实现联动逻辑：
        1. 勾选OCR时，自动勾选"提取图片"
        2. 取消"提取图片"时，自动取消OCR
        """
        if self._updating_image_options:
            return
        
        try:
            self._updating_image_options = True
            
            extract_image = self.image_keep_var.get()
            extract_ocr = self.image_ocr_var.get()
            
            # 检测状态变化
            ocr_changed = (extract_ocr != self._image_last_ocr_state)
            
            # 场景：用户勾选OCR（从无到有）
            if extract_ocr and not self._image_last_ocr_state and ocr_changed:
                if not extract_image:
                    logger.debug("图片设置：勾选OCR，自动勾选提取图片")
                    self.image_keep_var.set(True)
            
            # 更新状态记录
            self._image_last_image_state = self.image_keep_var.get()
            self._image_last_ocr_state = self.image_ocr_var.get()
            
            # 通知配置变更
            value = self.image_ocr_var.get()
            logger.info(f"图片OCR设置变更: {value}")
            self.on_change("to_md_enable_ocr", value)
        
        finally:
            self._updating_image_options = False
    
    def _on_compress_mode_changed(self):
        """处理压缩模式变更"""
        value = self.compress_mode_var.get()
        logger.info(f"压缩模式设置变更: {value}")
        self.on_change("compress_mode", value)
    
    def _on_pdf_quality_changed(self):
        """处理PDF质量变更"""
        value = self.pdf_quality_var.get()
        logger.info(f"PDF质量设置变更: {value}")
        self.on_change("pdf_quality", value)
    
    def _on_tiff_mode_changed(self):
        """处理TIFF模式变更"""
        value = self.tiff_mode_var.get()
        logger.info(f"TIFF模式设置变更: {value}")
        self.on_change("tiff_mode", value)
    
    # ========== 配置获取和应用 ==========
    
    def get_settings(self) -> Dict[str, Any]:
        """获取当前设置"""
        settings = {
            "to_md_keep_images": self.image_keep_var.get(),
            "to_md_enable_ocr": self.image_ocr_var.get(),
            "compress_mode": self.compress_mode_var.get(),
            "size_limit": self.size_limit_var.get(),
            "size_unit": self.size_unit_var.get(),
            "pdf_quality": self.pdf_quality_var.get(),
            "tiff_mode": self.tiff_mode_var.get()
        }
        logger.debug(f"获取图片设置: {settings}")
        return settings
    
    def apply_settings(self) -> bool:
        """应用设置到配置文件"""
        logger.debug("开始应用图片设置")
        
        try:
            settings = self.get_settings()
            success = True
            
            # 保存所有设置到conversion_defaults.toml
            for key, value in settings.items():
                if not self.config_manager.update_config_value(
                    "conversion_defaults", "image", key, value
                ):
                    success = False
            
            if success:
                logger.info("✓ 图片设置已成功应用")
            else:
                logger.error("✗ 部分图片设置更新失败")
            
            return success
            
        except Exception as e:
            logger.error(f"应用图片设置失败: {e}", exc_info=True)
            return False
