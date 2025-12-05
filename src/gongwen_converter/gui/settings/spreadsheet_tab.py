"""
表格设置选项卡模块

实现设置对话框的表格设置选项卡，包含：
- 提取/OCR默认设置
- 汇总模式默认设置
- 表格处理软件优先级
"""

import logging
import tkinter as tk
from dataclasses import dataclass
from typing import Dict, Any, Callable, List, Optional

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.gui.settings.base_tab import BaseSettingsTab
from gongwen_converter.gui.settings.config import SectionStyle
from gongwen_converter.utils.dpi_utils import scale

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
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
        """初始化表格设置选项卡"""
        # 软件名称映射
        self.software_names = {
            "wps_spreadsheets": "WPS表格",
            "msoffice_excel": "Excel",
            "libreoffice": "LibreOffice"
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
                SoftwareInfo(sw_id, self.software_names.get(sw_id, sw_id), i + 1)
                for i, sw_id in enumerate(spreadsheet_processors)
            ]
            
            ods_priority = config_manager.get_special_conversion_priority("ods")
            self.ods_software = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id, sw_id), i + 1)
                for i, sw_id in enumerate(ods_priority)
            ]
            
            sheet_to_pdf_priority = config_manager.get_spreadsheet_to_pdf_priority()
            self.spreadsheet_to_pdf_software = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id, sw_id), i + 1)
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
            "提取/OCR设置",
            SectionStyle.PRIMARY
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
            '默认勾选"提取图片"',
            self.sheet_keep_var,
            '表格转MD时默认勾选"提取图片"',
            self._on_sheet_keep_changed
        )
        
        # OCR识别
        self.sheet_ocr_var = tk.BooleanVar(value=sheet_ocr)
        self.create_checkbox_with_info(
            frame,
            '默认勾选"图片文字识别"',
            self.sheet_ocr_var,
            '表格转MD时默认勾选"图片文字识别"',
            self._on_sheet_ocr_changed
        )
    
    def _create_merge_mode_section(self):
        """创建汇总模式设置区域"""
        logger.debug("创建汇总模式设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "汇总模式设置",
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
            text="设置表格汇总的默认模式",
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
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 按行汇总
        row_frame = tb.Frame(radio_frame)
        row_frame.pack(fill="x", pady=(0, 8))
        
        tb.Radiobutton(
            row_frame,
            text="按行汇总",
            variable=self.merge_mode_var,
            value=1,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            row_frame,
            "将多个表格逐行合并至基准表格",
            "info"
        ).pack(side="left", padx=(5, 0))
        
        # 按列汇总
        col_frame = tb.Frame(radio_frame)
        col_frame.pack(fill="x", pady=(0, 8))
        
        tb.Radiobutton(
            col_frame,
            text="按列汇总",
            variable=self.merge_mode_var,
            value=2,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            col_frame,
            "将多个表格逐列合并至基准表格",
            "info"
        ).pack(side="left", padx=(5, 0))
        
        # 按单元格汇总
        cell_frame = tb.Frame(radio_frame)
        cell_frame.pack(fill="x")
        
        tb.Radiobutton(
            cell_frame,
            text="按单元格汇总",
            variable=self.merge_mode_var,
            value=3,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        ).pack(side="left")
        
        create_info_icon(
            cell_frame,
            "对应位置的单元格数据相加",
            "info"
        ).pack(side="left", padx=(5, 0))
    
    def _create_software_priority_section(self):
        """创建软件优先级区域"""
        logger.debug("创建软件优先级区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "软件优先级",
            SectionStyle.SUCCESS
        )
        
        # 表格处理软件
        sheet_label = tb.Label(
            frame,
            text="表格处理软件优先级",
            font=(self.small_font, self.small_size, "bold")
        )
        sheet_label.pack(anchor="w", pady=(0, 10))
        
        self.spreadsheet_cards_frame = tb.Frame(frame)
        self.spreadsheet_cards_frame.pack(fill="x", pady=(0, 15))
        
        # ODS格式转换
        ods_label = tb.Label(
            frame,
            text="ODS格式转换",
            font=(self.small_font, self.small_size, "bold")
        )
        ods_label.pack(anchor="w", pady=(0, 10))
        
        self.ods_cards_frame = tb.Frame(frame)
        self.ods_cards_frame.pack(fill="x", pady=(0, 15))
        
        # 表格转PDF
        pdf_label = tb.Label(
            frame,
            text="表格转PDF",
            font=(self.small_font, self.small_size, "bold")
        )
        pdf_label.pack(anchor="w", pady=(0, 10))
        
        self.spreadsheet_to_pdf_cards_frame = tb.Frame(frame)
        self.spreadsheet_to_pdf_cards_frame.pack(fill="x")
    
    # ========== 软件优先级方法 ==========
    
    def _create_software_card(self, parent, software_info: SoftwareInfo, category: str, index: int, list_len: int):
        """创建软件卡片"""
        card = tk.Frame(parent)
        
        try:
            style = tb.Style.get_instance()
            parent_bg = style.colors.bg
            
            if software_info.is_selected:
                primary_color = style.colors.primary
                card.configure(
                    bg=parent_bg,
                    highlightbackground=primary_color,
                    highlightcolor=primary_color,
                    highlightthickness=scale(2)
                )
            else:
                card.configure(
                    bg=parent_bg,
                    highlightbackground="#DDDDDD",
                    highlightcolor="#DDDDDD",
                    highlightthickness=scale(1)
                )
        except:
            card.configure(bg='SystemButtonFace', highlightthickness=scale(1 if not software_info.is_selected else 2))
        
        card.pack(side="left", padx=scale(3))
        
        if software_info.is_selected:
            if index > 0:
                left_btn = tk.Button(card, text="◀", width=1, command=lambda: self._move_software('left'))
                try:
                    primary_color = tb.Style.get_instance().colors.primary
                    left_btn.configure(
                        bg=primary_color,
                        fg="white",
                        activebackground=primary_color,
                        activeforeground="white",
                        relief=tk.FLAT,
                        borderwidth=0,
                        padx=scale(0),
                        pady=scale(6),
                        cursor="hand2",
                        font=(self.small_font, self.small_size)
                    )
                except:
                    pass
                left_btn.pack(side="left")
            
            name_label = tb.Label(
                card,
                text=software_info.display_name,
                font=(self.small_font, self.small_size),
                bootstyle="default",
                anchor=tk.CENTER,
                cursor="hand2",
                width=12
            )
            name_label.pack(side="left", pady=(scale(6), scale(6)))
            
            if index < list_len - 1:
                right_btn = tk.Button(card, text="▶", width=1, command=lambda: self._move_software('right'))
                try:
                    primary_color = tb.Style.get_instance().colors.primary
                    right_btn.configure(
                        bg=primary_color,
                        fg="white",
                        activebackground=primary_color,
                        activeforeground="white",
                        relief=tk.FLAT,
                        borderwidth=0,
                        padx=scale(0),
                        pady=scale(6),
                        cursor="hand2",
                        font=(self.small_font, self.small_size)
                    )
                except:
                    pass
                right_btn.pack(side="left")
            
            name_label.bind("<Button-1>", lambda e, si=software_info, cat=category: self._select_software(si, cat))
        else:
            name_label = tb.Label(
                card,
                text=software_info.display_name,
                font=(self.small_font, self.small_size),
                bootstyle="default",
                anchor=tk.CENTER,
                cursor="hand2",
                width=12
            )
            name_label.pack(pady=(scale(8), scale(8)))
            
            card.bind("<Button-1>", lambda e, si=software_info, cat=category: self._select_software(si, cat))
            name_label.bind("<Button-1>", lambda e, si=software_info, cat=category: self._select_software(si, cat))
    
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
        
        for widget in cards_frame.winfo_children():
            widget.destroy()
        
        list_len = len(software_list)
        for i, sw in enumerate(software_list):
            self._create_software_card(cards_frame, sw, category, i, list_len)
    
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
        """处理提取图片设置变更"""
        value = self.sheet_keep_var.get()
        logger.info(f"表格提取图片设置变更: {value}")
        self.on_change("to_md_keep_images", value)
    
    def _on_sheet_ocr_changed(self):
        """处理OCR设置变更"""
        value = self.sheet_ocr_var.get()
        logger.info(f"表格OCR设置变更: {value}")
        self.on_change("to_md_enable_ocr", value)
    
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
            
            # 保存到file_defaults.toml
            file_defaults_keys = ["to_md_keep_images", "to_md_enable_ocr", "merge_mode"]
            for key in file_defaults_keys:
                if not self.config_manager.update_config_value(
                    "file_defaults", "spreadsheet", key, settings[key]
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
