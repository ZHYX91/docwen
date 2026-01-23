"""
版式设置选项卡模块

实现设置对话框的版式设置选项卡，包含：
- 提取/OCR默认设置
- 渲染DPI默认设置
- PDF转换软件优先级

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。
"""

import logging
import tkinter as tk
from dataclasses import dataclass
from typing import Dict, Any, Callable, List, Optional

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.gui.settings.base_tab import BaseSettingsTab
from docwen.gui.settings.config import SectionStyle
from docwen.utils.dpi_utils import scale
from docwen.i18n import t

logger = logging.getLogger(__name__)


@dataclass
class SoftwareInfo:
    """软件信息数据类"""
    software_id: str
    display_name: str
    priority: int
    is_selected: bool = False


class LayoutTab(BaseSettingsTab):
    """
    版式设置选项卡类
    
    管理版式文件相关的所有配置选项。
    包含提取/OCR设置、渲染DPI设置和软件优先级。
    """
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
        """初始化版式设置选项卡"""
        # 软件名称映射
        self.software_names = {
            "msoffice_word": "Word",
            "libreoffice": "LibreOffice"
        }
        
        # 软件列表
        self.pdf_to_office_software: List[SoftwareInfo] = []
        
        # 选中软件
        self.selected_software: Optional[SoftwareInfo] = None
        self.selected_category: Optional[str] = None
        
        # 卡片容器
        self.pdf_to_office_cards_frame: Optional[tb.Frame] = None
        
        # 联动逻辑标志位和状态记录
        self._updating_layout_options: bool = False
        self._layout_last_image_state: bool = False
        self._layout_last_ocr_state: bool = False
        
        # 加载配置
        self._load_settings_data(config_manager)
        
        super().__init__(parent, config_manager, on_change)
        logger.info("版式设置选项卡初始化完成")
    
    def _load_settings_data(self, config_manager):
        """加载配置数据"""
        try:
            pdf_to_office_priority = config_manager.get_special_conversion_priority("pdf_to_office")
            self.pdf_to_office_software = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id, sw_id), i + 1)
                for i, sw_id in enumerate(pdf_to_office_priority)
            ]
            
            # 默认选中第一个
            if self.pdf_to_office_software:
                self.pdf_to_office_software[0].is_selected = True
                self.selected_software = self.pdf_to_office_software[0]
                self.selected_category = 'pdf_to_office'
                
        except Exception as e:
            logger.error(f"加载版式配置失败: {e}")
    
    def _create_interface(self):
        """创建选项卡界面"""
        logger.debug("开始创建版式设置选项卡界面")
        
        self._create_extraction_section()
        self._create_dpi_section()
        self._create_software_priority_section()
        
        logger.debug("版式设置选项卡界面创建完成")
    
    def _post_initialize(self):
        """初始化后处理"""
        self._refresh_category('pdf_to_office')
        super()._post_initialize()
    
    def _create_extraction_section(self):
        """创建提取/OCR设置区域"""
        logger.debug("创建提取/OCR设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.layout.extraction_section"),
            SectionStyle.DANGER
        )
        
        # 获取当前配置
        try:
            layout_keep = self.config_manager.get_layout_to_md_keep_images()
            layout_ocr = self.config_manager.get_layout_to_md_enable_ocr()
        except Exception as e:
            logger.warning(f"读取版式配置失败: {e}")
            layout_keep, layout_ocr = True, False
        
        # 提取图片
        self.layout_keep_var = tk.BooleanVar(value=layout_keep)
        self.create_checkbox_with_info(
            frame,
            t("settings.layout.keep_images"),
            self.layout_keep_var,
            t("settings.layout.keep_images_tooltip"),
            self._on_layout_keep_changed
        )
        
        # OCR识别
        self.layout_ocr_var = tk.BooleanVar(value=layout_ocr)
        self.create_checkbox_with_info(
            frame,
            t("settings.layout.enable_ocr"),
            self.layout_ocr_var,
            t("settings.layout.enable_ocr_tooltip"),
            self._on_layout_ocr_changed
        )
        
        # 初始化状态记录
        self._layout_last_image_state = layout_keep
        self._layout_last_ocr_state = layout_ocr
    
    def _create_dpi_section(self):
        """创建渲染DPI设置区域"""
        logger.debug("创建渲染DPI设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.layout.dpi_section"),
            SectionStyle.INFO
        )
        
        # 获取当前配置
        try:
            render_dpi = self.config_manager.get_layout_render_dpi()
        except Exception as e:
            logger.warning(f"读取渲染DPI失败: {e}")
            render_dpi = 300
        
        # 说明文本
        desc_label = tb.Label(
            frame,
            text=t("settings.layout.dpi_desc"),
            bootstyle="secondary",
            wraplength=400
        )
        desc_label.pack(anchor="w", pady=(0, 10))
        
        # 单选按钮变量
        self.render_dpi_var = tk.IntVar(value=render_dpi)
        
        # 单选按钮容器
        radio_frame = tb.Frame(frame)
        radio_frame.pack(fill="x", pady=(10, 0))
        
        # 导入信息图标
        from docwen.utils.gui_utils import create_info_icon
        
        # 150 DPI
        dpi_150_frame = tb.Frame(radio_frame)
        dpi_150_frame.pack(fill="x", pady=(0, 8))
        
        tb.Radiobutton(
            dpi_150_frame,
            text=t("settings.layout.dpi_min"),
            variable=self.render_dpi_var,
            value=150,
            command=self._on_dpi_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            dpi_150_frame,
            t("settings.layout.dpi_min_tooltip"),
            "info"
        ).pack(side="left", padx=(5, 0))
        
        # 300 DPI
        dpi_300_frame = tb.Frame(radio_frame)
        dpi_300_frame.pack(fill="x", pady=(0, 8))
        
        tb.Radiobutton(
            dpi_300_frame,
            text=t("settings.layout.dpi_medium"),
            variable=self.render_dpi_var,
            value=300,
            command=self._on_dpi_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            dpi_300_frame,
            t("settings.layout.dpi_medium_tooltip"),
            "info"
        ).pack(side="left", padx=(5, 0))
        
        # 600 DPI
        dpi_600_frame = tb.Frame(radio_frame)
        dpi_600_frame.pack(fill="x")
        
        tb.Radiobutton(
            dpi_600_frame,
            text=t("settings.layout.dpi_high"),
            variable=self.render_dpi_var,
            value=600,
            command=self._on_dpi_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            dpi_600_frame,
            t("settings.layout.dpi_high_tooltip"),
            "info"
        ).pack(side="left", padx=(5, 0))
    
    def _create_software_priority_section(self):
        """创建软件优先级区域"""
        logger.debug("创建软件优先级区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.layout.software_section"),
            SectionStyle.SUCCESS
        )
        
        # PDF转文档
        pdf_label = tb.Label(
            frame,
            text=t("settings.layout.pdf_to_doc_label"),
            font=(self.small_font, self.small_size, "bold")
        )
        pdf_label.pack(anchor="w", pady=(0, 10))
        
        self.pdf_to_office_cards_frame = tb.Frame(frame)
        self.pdf_to_office_cards_frame.pack(fill="x")
    
    # ========== 软件优先级方法 ==========
    
    def _create_software_card_for_category(self, parent, software_info: SoftwareInfo, category: str, index: int, list_len: int):
        """创建软件卡片（使用基类方法）"""
        self.create_software_card(
            parent=parent,
            software_id=software_info.software_id,
            display_name=software_info.display_name,
            is_selected=software_info.is_selected,
            index=index,
            list_len=list_len,
            on_select=lambda: self._select_software(software_info, category),
            on_move=self._move_software
        )
    
    def _select_software(self, software_info: SoftwareInfo, category: str):
        """选中软件"""
        for sw in self.pdf_to_office_software:
            sw.is_selected = False
        
        software_info.is_selected = True
        self.selected_software = software_info
        self.selected_category = category
        
        self._refresh_category('pdf_to_office')
    
    def _refresh_category(self, category: str):
        """刷新指定类别"""
        if category == 'pdf_to_office':
            cards_frame = self.pdf_to_office_cards_frame
            software_list = self.pdf_to_office_software
        else:
            return
        
        for widget in cards_frame.winfo_children():
            widget.destroy()
        
        list_len = len(software_list)
        for i, sw in enumerate(software_list):
            self._create_software_card_for_category(cards_frame, sw, category, i, list_len)
    
    def _move_software(self, direction: str):
        """移动软件位置"""
        if not self.selected_software or not self.selected_category:
            return
        
        software_list = self.pdf_to_office_software
        index = software_list.index(self.selected_software)
        
        if direction == 'left' and index > 0:
            software_list[index], software_list[index - 1] = software_list[index - 1], software_list[index]
        elif direction == 'right' and index < len(software_list) - 1:
            software_list[index], software_list[index + 1] = software_list[index + 1], software_list[index]
        
        for i, sw in enumerate(software_list):
            sw.priority = i + 1
        
        self._refresh_category('pdf_to_office')
    
    # ========== 事件处理 ==========
    
    def _on_layout_keep_changed(self):
        """
        处理提取图片设置变更
        
        实现联动逻辑：
        1. 勾选OCR时，自动勾选"提取图片"
        2. 取消"提取图片"时，自动取消OCR
        """
        if self._updating_layout_options:
            return
        
        try:
            self._updating_layout_options = True
            
            extract_image = self.layout_keep_var.get()
            extract_ocr = self.layout_ocr_var.get()
            
            # 检测状态变化
            image_changed = (extract_image != self._layout_last_image_state)
            
            # 场景：用户取消提取图片（从有到无）
            if not extract_image and self._layout_last_image_state and image_changed:
                if extract_ocr:
                    logger.debug("版式设置：取消提取图片，自动取消OCR")
                    self.layout_ocr_var.set(False)
            
            # 更新状态记录
            self._layout_last_image_state = self.layout_keep_var.get()
            self._layout_last_ocr_state = self.layout_ocr_var.get()
            
            # 通知配置变更
            value = self.layout_keep_var.get()
            logger.info(f"版式提取图片设置变更: {value}")
            self.on_change("to_md_keep_images", value)
        
        finally:
            self._updating_layout_options = False
    
    def _on_layout_ocr_changed(self):
        """
        处理OCR设置变更
        
        实现联动逻辑：
        1. 勾选OCR时，自动勾选"提取图片"
        2. 取消"提取图片"时，自动取消OCR
        """
        if self._updating_layout_options:
            return
        
        try:
            self._updating_layout_options = True
            
            extract_image = self.layout_keep_var.get()
            extract_ocr = self.layout_ocr_var.get()
            
            # 检测状态变化
            ocr_changed = (extract_ocr != self._layout_last_ocr_state)
            
            # 场景：用户勾选OCR（从无到有）
            if extract_ocr and not self._layout_last_ocr_state and ocr_changed:
                if not extract_image:
                    logger.debug("版式设置：勾选OCR，自动勾选提取图片")
                    self.layout_keep_var.set(True)
            
            # 更新状态记录
            self._layout_last_image_state = self.layout_keep_var.get()
            self._layout_last_ocr_state = self.layout_ocr_var.get()
            
            # 通知配置变更
            value = self.layout_ocr_var.get()
            logger.info(f"版式OCR设置变更: {value}")
            self.on_change("to_md_enable_ocr", value)
        
        finally:
            self._updating_layout_options = False
    
    def _on_dpi_changed(self):
        """处理DPI设置变更"""
        value = self.render_dpi_var.get()
        logger.info(f"渲染DPI设置变更: {value}")
        self.on_change("render_dpi", value)
    
    # ========== 配置获取和应用 ==========
    
    def get_settings(self) -> Dict[str, Any]:
        """获取当前设置"""
        settings = {
            "to_md_keep_images": self.layout_keep_var.get(),
            "to_md_enable_ocr": self.layout_ocr_var.get(),
            "render_dpi": self.render_dpi_var.get(),
            "pdf_to_office": [sw.software_id for sw in self.pdf_to_office_software]
        }
        logger.debug(f"获取版式设置: {settings}")
        return settings
    
    def apply_settings(self) -> bool:
        """应用设置到配置文件"""
        logger.debug("开始应用版式设置")
        
        try:
            settings = self.get_settings()
            success = True
            
            # 保存到conversion_defaults.toml
            conversion_defaults_keys = ["to_md_keep_images", "to_md_enable_ocr", "render_dpi"]
            for key in conversion_defaults_keys:
                if not self.config_manager.update_config_value(
                    "conversion_defaults", "layout", key, settings[key]
                ):
                    success = False
            
            # 保存软件优先级
            special_conversions = {
                "odt": self.config_manager.get_special_conversion_priority("odt"),
                "ods": self.config_manager.get_special_conversion_priority("ods"),
                "pdf_to_office": settings["pdf_to_office"],
                "document_to_pdf": self.config_manager.get_document_to_pdf_priority(),
                "spreadsheet_to_pdf": self.config_manager.get_spreadsheet_to_pdf_priority()
            }
            
            if not self.config_manager.update_config_section(
                "software_priority", "special_conversions", special_conversions
            ):
                success = False
            
            if success:
                logger.info("✓ 版式设置已成功应用")
            else:
                logger.error("✗ 部分版式设置更新失败")
            
            return success
            
        except Exception as e:
            logger.error(f"应用版式设置失败: {e}", exc_info=True)
            return False
