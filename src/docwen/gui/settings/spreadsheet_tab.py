"""
表格设置选项卡模块

实现设置对话框的表格设置选项卡，包含：
- 提取/OCR默认设置
- 汇总模式默认设置
- 表格处理软件优先级
"""

from __future__ import annotations

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


class SpreadsheetTab(BaseSettingsTab):
    """
    表格设置选项卡类
    
    管理表格文件相关的所有配置选项。
    包含提取/OCR设置、汇总模式设置和软件优先级。
    """
    
    def __init__(self, parent, config_manager: Any, on_change: Callable[[str, Any], None]):
        """初始化表格设置选项卡"""
        # 软件名称映射（使用国际化）
        self.software_names = {
            "wps_spreadsheets": t("settings.spreadsheet.software.wps_spreadsheets"),
            "msoffice_excel": t("settings.spreadsheet.software.excel"),
            "libreoffice": t("settings.spreadsheet.software.libreoffice")
        }
        
        # 软件列表
        self.spreadsheet_processors: List[SoftwareInfo] = []
        self.ods_software: List[SoftwareInfo] = []
        self.spreadsheet_to_pdf_software: List[SoftwareInfo] = []
        
        # 选中软件
        self.selected_software: Optional[SoftwareInfo] = None
        self.selected_category: Optional[str] = None
        
        # 卡片容器
        self.spreadsheet_cards_frame: Optional[tb.Frame] = None
        self.ods_cards_frame: Optional[tb.Frame] = None
        self.spreadsheet_to_pdf_cards_frame: Optional[tb.Frame] = None
        
        # 联动逻辑标志位和状态记录
        self._updating_sheet_options: bool = False
        self._sheet_last_image_state: bool = False
        self._sheet_last_ocr_state: bool = False
        
        # 加载配置
        self._load_settings_data(config_manager)
        
        super().__init__(parent, config_manager, on_change)
        logger.info("表格设置选项卡初始化完成")
    
    def _load_settings_data(self, config_manager):
        """加载配置数据"""
        try:
            # 加载软件优先级
            spreadsheet_processors = config_manager.get_spreadsheet_processors_priority()
            self.spreadsheet_processors = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id) or sw_id, i + 1)
                for i, sw_id in enumerate(spreadsheet_processors)
            ]
            
            ods_priority = config_manager.get_special_conversion_priority("ods")
            self.ods_software = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id) or sw_id, i + 1)
                for i, sw_id in enumerate(ods_priority)
            ]
            
            sheet_to_pdf_priority = config_manager.get_spreadsheet_to_pdf_priority()
            self.spreadsheet_to_pdf_software = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id) or sw_id, i + 1)
                for i, sw_id in enumerate(sheet_to_pdf_priority)
            ]
            
            # 默认选中第一个
            if self.spreadsheet_processors:
                self.spreadsheet_processors[0].is_selected = True
                self.selected_software = self.spreadsheet_processors[0]
                self.selected_category = 'spreadsheet'
                
        except Exception as e:
            logger.error(f"加载表格配置失败: {e}")
    
    def _create_interface(self):
        """创建选项卡界面"""
        logger.debug("开始创建表格设置选项卡界面")
        
        self._create_extraction_section()
        self._create_merge_mode_section()
        self._create_software_priority_section()
        
        logger.debug("表格设置选项卡界面创建完成")
    
    def _post_initialize(self):
        """初始化后处理"""
        self._refresh_all_categories()
        super()._post_initialize()
    
    def _create_extraction_section(self):
        """创建提取/OCR设置区域"""
        logger.debug("创建提取/OCR设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.spreadsheet.extraction_section"),
            SectionStyle.DANGER
        )
        
        # 获取当前配置
        try:
            sheet_keep = self.config_manager.get_xlsx_to_md_keep_images()
            sheet_ocr = self.config_manager.get_xlsx_to_md_enable_ocr()
        except Exception as e:
            logger.warning(f"读取表格配置失败: {e}")
            sheet_keep, sheet_ocr = True, False
        
        # 提取图片
        self.sheet_keep_var = tk.BooleanVar(value=sheet_keep)
        self.create_checkbox_with_info(
            frame,
            t("settings.spreadsheet.keep_images"),
            self.sheet_keep_var,
            t("settings.spreadsheet.keep_images_tooltip"),
            self._on_sheet_keep_changed
        )
        
        # OCR识别
        self.sheet_ocr_var = tk.BooleanVar(value=sheet_ocr)
        self.create_checkbox_with_info(
            frame,
            t("settings.spreadsheet.enable_ocr"),
            self.sheet_ocr_var,
            t("settings.spreadsheet.enable_ocr_tooltip"),
            self._on_sheet_ocr_changed
        )
        
        # 初始化状态记录
        self._sheet_last_image_state = sheet_keep
        self._sheet_last_ocr_state = sheet_ocr
    
    def _create_merge_mode_section(self):
        """创建汇总模式设置区域"""
        logger.debug("创建汇总模式设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.spreadsheet.merge_mode_section"),
            SectionStyle.INFO
        )
        
        # 获取当前配置
        try:
            merge_mode = self.config_manager.get_spreadsheet_merge_mode()
        except Exception as e:
            logger.warning(f"读取汇总模式失败: {e}")
            merge_mode = 3
        
        # 说明文本
        desc_label = tb.Label(
            frame,
            text=t("settings.spreadsheet.merge_mode_desc"),
            bootstyle="secondary",
            wraplength=400
        )
        desc_label.pack(anchor="w", pady=(0, 10))
        
        # 单选按钮变量
        self.merge_mode_var = tk.IntVar(value=merge_mode)
        
        # 单选按钮容器
        radio_frame = tb.Frame(frame)
        radio_frame.pack(fill="x", pady=(10, 0))
        
        # 导入信息图标
        from docwen.utils.gui_utils import create_info_icon
        
        # 按行汇总
        row_frame = tb.Frame(radio_frame)
        row_frame.pack(fill="x", pady=(0, 8))
        
        tb.Radiobutton(
            row_frame,
            text=t("settings.spreadsheet.merge_modes.by_row"),
            variable=self.merge_mode_var,
            value=1,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            row_frame,
            t("settings.spreadsheet.merge_modes.by_row_tooltip"),
            "info"
        ).pack(side="left", padx=(5, 0))
        
        # 按列汇总
        col_frame = tb.Frame(radio_frame)
        col_frame.pack(fill="x", pady=(0, 8))
        
        tb.Radiobutton(
            col_frame,
            text=t("settings.spreadsheet.merge_modes.by_column"),
            variable=self.merge_mode_var,
            value=2,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            col_frame,
            t("settings.spreadsheet.merge_modes.by_column_tooltip"),
            "info"
        ).pack(side="left", padx=(5, 0))
        
        # 按单元格汇总
        cell_frame = tb.Frame(radio_frame)
        cell_frame.pack(fill="x")
        
        tb.Radiobutton(
            cell_frame,
            text=t("settings.spreadsheet.merge_modes.by_cell"),
            variable=self.merge_mode_var,
            value=3,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            cell_frame,
            t("settings.spreadsheet.merge_modes.by_cell_tooltip"),
            "info"
        ).pack(side="left", padx=(5, 0))
    
    def _create_software_priority_section(self):
        """创建软件优先级区域"""
        logger.debug("创建软件优先级区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.spreadsheet.software_section"),
            SectionStyle.SUCCESS
        )
        
        # 表格处理软件
        sheet_label = tb.Label(
            frame,
            text=t("settings.spreadsheet.spreadsheet_processors_label"),
            font=(self.small_font, self.small_size, "bold")
        )
        sheet_label.pack(anchor="w", pady=(0, 10))
        
        self.spreadsheet_cards_frame = tb.Frame(frame)
        self.spreadsheet_cards_frame.pack(fill="x", pady=(0, 15))
        
        # ODS格式转换
        ods_label = tb.Label(
            frame,
            text=t("settings.spreadsheet.ods_conversion_label"),
            font=(self.small_font, self.small_size, "bold")
        )
        ods_label.pack(anchor="w", pady=(0, 10))
        
        self.ods_cards_frame = tb.Frame(frame)
        self.ods_cards_frame.pack(fill="x", pady=(0, 15))
        
        # 表格转PDF
        pdf_label = tb.Label(
            frame,
            text=t("settings.spreadsheet.spreadsheet_to_pdf_label"),
            font=(self.small_font, self.small_size, "bold")
        )
        pdf_label.pack(anchor="w", pady=(0, 10))
        
        self.spreadsheet_to_pdf_cards_frame = tb.Frame(frame)
        self.spreadsheet_to_pdf_cards_frame.pack(fill="x")
    
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
        for sw in self.spreadsheet_processors:
            sw.is_selected = False
        for sw in self.ods_software:
            sw.is_selected = False
        for sw in self.spreadsheet_to_pdf_software:
            sw.is_selected = False
        
        software_info.is_selected = True
        self.selected_software = software_info
        self.selected_category = category
        
        self._refresh_all_categories()
    
    def _refresh_all_categories(self):
        """刷新所有类别"""
        self._refresh_category('spreadsheet')
        self._refresh_category('ods')
        self._refresh_category('spreadsheet_to_pdf')
    
    def _refresh_category(self, category: str):
        """刷新指定类别"""
        if category == 'spreadsheet':
            cards_frame = self.spreadsheet_cards_frame
            software_list = self.spreadsheet_processors
        elif category == 'ods':
            cards_frame = self.ods_cards_frame
            software_list = self.ods_software
        elif category == 'spreadsheet_to_pdf':
            cards_frame = self.spreadsheet_to_pdf_cards_frame
            software_list = self.spreadsheet_to_pdf_software
        else:
            return

        if not cards_frame:
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
        
        if self.selected_category == 'spreadsheet':
            software_list = self.spreadsheet_processors
        elif self.selected_category == 'ods':
            software_list = self.ods_software
        elif self.selected_category == 'spreadsheet_to_pdf':
            software_list = self.spreadsheet_to_pdf_software
        else:
            return
        
        index = software_list.index(self.selected_software)
        
        if direction == 'left' and index > 0:
            software_list[index], software_list[index - 1] = software_list[index - 1], software_list[index]
        elif direction == 'right' and index < len(software_list) - 1:
            software_list[index], software_list[index + 1] = software_list[index + 1], software_list[index]
        
        for i, sw in enumerate(software_list):
            sw.priority = i + 1
        
        self._refresh_category(self.selected_category)
    
    # ========== 事件处理 ==========
    
    def _on_sheet_keep_changed(self):
        """
        处理提取图片设置变更
        
        实现联动逻辑：
        1. 勾选OCR时，自动勾选"提取图片"
        2. 取消"提取图片"时，自动取消OCR
        """
        if self._updating_sheet_options:
            return
        
        try:
            self._updating_sheet_options = True
            
            extract_image = self.sheet_keep_var.get()
            extract_ocr = self.sheet_ocr_var.get()
            
            # 检测状态变化
            image_changed = (extract_image != self._sheet_last_image_state)
            
            # 场景：用户取消提取图片（从有到无）
            if not extract_image and self._sheet_last_image_state and image_changed:
                if extract_ocr:
                    logger.debug("表格设置：取消提取图片，自动取消OCR")
                    self.sheet_ocr_var.set(False)
            
            # 更新状态记录
            self._sheet_last_image_state = self.sheet_keep_var.get()
            self._sheet_last_ocr_state = self.sheet_ocr_var.get()
            
            # 通知配置变更
            value = self.sheet_keep_var.get()
            logger.info(f"表格提取图片设置变更: {value}")
            self.on_change("to_md_keep_images", value)
        
        finally:
            self._updating_sheet_options = False
    
    def _on_sheet_ocr_changed(self):
        """
        处理OCR设置变更
        
        实现联动逻辑：
        1. 勾选OCR时，自动勾选"提取图片"
        2. 取消"提取图片"时，自动取消OCR
        """
        if self._updating_sheet_options:
            return
        
        try:
            self._updating_sheet_options = True
            
            extract_image = self.sheet_keep_var.get()
            extract_ocr = self.sheet_ocr_var.get()
            
            # 检测状态变化
            ocr_changed = (extract_ocr != self._sheet_last_ocr_state)
            
            # 场景：用户勾选OCR（从无到有）
            if extract_ocr and not self._sheet_last_ocr_state and ocr_changed:
                if not extract_image:
                    logger.debug("表格设置：勾选OCR，自动勾选提取图片")
                    self.sheet_keep_var.set(True)
            
            # 更新状态记录
            self._sheet_last_image_state = self.sheet_keep_var.get()
            self._sheet_last_ocr_state = self.sheet_ocr_var.get()
            
            # 通知配置变更
            value = self.sheet_ocr_var.get()
            logger.info(f"表格OCR设置变更: {value}")
            self.on_change("to_md_enable_ocr", value)
        
        finally:
            self._updating_sheet_options = False
    
    def _on_merge_mode_changed(self):
        """处理汇总模式变更"""
        value = self.merge_mode_var.get()
        logger.info(f"汇总模式设置变更: {value}")
        self.on_change("merge_mode", value)
    
    # ========== 配置获取和应用 ==========
    
    def get_settings(self) -> Dict[str, Any]:
        """获取当前设置"""
        settings = {
            "to_md_keep_images": self.sheet_keep_var.get(),
            "to_md_enable_ocr": self.sheet_ocr_var.get(),
            "merge_mode": self.merge_mode_var.get(),
            "spreadsheet_processors": [sw.software_id for sw in self.spreadsheet_processors],
            "ods": [sw.software_id for sw in self.ods_software],
            "spreadsheet_to_pdf": [sw.software_id for sw in self.spreadsheet_to_pdf_software]
        }
        logger.debug(f"获取表格设置: {settings}")
        return settings
    
    def apply_settings(self) -> bool:
        """应用设置到配置文件"""
        logger.debug("开始应用表格设置")
        
        try:
            settings = self.get_settings()
            success = True
            
            # 保存到conversion_defaults.toml
            conversion_defaults_keys = ["to_md_keep_images", "to_md_enable_ocr", "merge_mode"]
            for key in conversion_defaults_keys:
                if not self.config_manager.update_config_value(
                    "conversion_defaults", "spreadsheet", key, settings[key]
                ):
                    success = False
            
            # 保存软件优先级
            if not self.config_manager.update_config_value(
                "software_priority", "default_priority", "spreadsheet_processors", settings["spreadsheet_processors"]
            ):
                success = False
            
            special_conversions = {
                "odt": self.config_manager.get_special_conversion_priority("odt"),
                "ods": settings["ods"],
                "pdf_to_office": self.config_manager.get_special_conversion_priority("pdf_to_office"),
                "document_to_pdf": self.config_manager.get_document_to_pdf_priority(),
                "spreadsheet_to_pdf": settings["spreadsheet_to_pdf"]
            }
            
            if not self.config_manager.update_config_section(
                "software_priority", "special_conversions", special_conversions
            ):
                success = False
            
            if success:
                logger.info("✓ 表格设置已成功应用")
            else:
                logger.error("✗ 部分表格设置更新失败")
            
            return success
            
        except Exception as e:
            logger.error(f"应用表格设置失败: {e}", exc_info=True)
            return False
