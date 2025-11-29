"""
软件优先级设置选项卡模块

实现设置对话框的软件优先级选项卡，包含：
- 默认软件优先级（文档处理、表格处理）
- 特殊格式转换优先级（ODT、ODS、PDF转换）
- 可视化软件卡片交互界面
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
    """
    软件信息数据类
    
    封装单个软件的完整信息，用于UI显示和配置管理。
    
    属性：
        software_id: 软件配置ID（如 "wps_writer"）
        display_name: 显示名称（如 "WPS文字"）
        priority: 优先级（1表示最高优先级）
        is_selected: 是否被选中（用于UI交互）
    """
    software_id: str
    display_name: str
    priority: int
    is_selected: bool = False


class SoftwareTab(BaseSettingsTab):
    """
    软件优先级设置选项卡类
    
    管理文档处理和表格处理软件的优先级顺序。
    包含默认优先级设置和特殊格式转换优先级两个主要区域。
    """

    def __init__(self, parent, config_manager: Any, on_change: Callable[[str, Any], None]):
        """初始化软件优先级设置选项卡"""
        # 软件名称映射（配置ID -> 显示名称）
        self.software_names = {
            "wps_writer": "WPS文字",
            "msoffice_word": "Word",
            "libreoffice": "LibreOffice",
            "wps_spreadsheets": "WPS表格",
            "msoffice_excel": "Excel"
        }

        # 软件列表存储
        self.word_processors: List[SoftwareInfo] = []
        self.spreadsheet_processors: List[SoftwareInfo] = []
        self.odt_software: List[SoftwareInfo] = []
        self.ods_software: List[SoftwareInfo] = []
        self.pdf_to_office_software: List[SoftwareInfo] = []
        self.document_to_pdf_software: List[SoftwareInfo] = []
        self.spreadsheet_to_pdf_software: List[SoftwareInfo] = []
        
        # 全局选中的软件（整个页面只有一个选中）
        self.selected_software: Optional[SoftwareInfo] = None
        self.selected_category: Optional[str] = None

        # 卡片容器引用（稍后创建）
        self.word_cards_frame: Optional[tb.Frame] = None
        self.spreadsheet_cards_frame: Optional[tb.Frame] = None
        self.odt_cards_frame: Optional[tb.Frame] = None
        self.ods_cards_frame: Optional[tb.Frame] = None
        self.pdf_to_office_cards_frame: Optional[tb.Frame] = None
        self.document_to_pdf_cards_frame: Optional[tb.Frame] = None
        self.spreadsheet_to_pdf_cards_frame: Optional[tb.Frame] = None

        # 加载当前配置（在基类初始化之前）
        self._load_current_settings_data(config_manager)

        # 调用基类初始化（会自动调用 _create_interface）
        super().__init__(parent, config_manager, on_change)
        
        logger.info("软件优先级设置选项卡初始化完成")

    def _load_current_settings_data(self, config_manager):
        """
        加载当前配置数据
        
        在界面创建之前加载所有软件优先级配置。
        """
        logger.debug("加载软件优先级配置数据")

        try:
            # 获取当前配置
            default_priority = config_manager.get_default_priority_config()
            special_conversions = config_manager.get_special_conversions_config()

            # 加载文档处理软件优先级
            word_processors = default_priority.get("word_processors", ["wps_writer", "msoffice_word", "libreoffice"])
            self.word_processors = []
            for i, software_id in enumerate(word_processors):
                display_name = self.software_names.get(software_id, software_id)
                sw_info = SoftwareInfo(software_id, display_name, i + 1)
                self.word_processors.append(sw_info)

            # 加载表格处理软件优先级
            spreadsheet_processors = default_priority.get("spreadsheet_processors", ["wps_spreadsheets", "msoffice_excel", "libreoffice"])
            self.spreadsheet_processors = []
            for i, software_id in enumerate(spreadsheet_processors):
                display_name = self.software_names.get(software_id, software_id)
                sw_info = SoftwareInfo(software_id, display_name, i + 1)
                self.spreadsheet_processors.append(sw_info)

            # 加载特殊转换配置 - ODT
            odt_priority = special_conversions.get("odt", ["msoffice_word", "libreoffice"])
            self.odt_software = []
            for i, software_id in enumerate(odt_priority):
                display_name = self.software_names.get(software_id, software_id)
                sw_info = SoftwareInfo(software_id, display_name, i + 1)
                self.odt_software.append(sw_info)

            # 加载特殊转换配置 - ODS
            ods_priority = special_conversions.get("ods", ["msoffice_excel", "libreoffice"])
            self.ods_software = []
            for i, software_id in enumerate(ods_priority):
                display_name = self.software_names.get(software_id, software_id)
                sw_info = SoftwareInfo(software_id, display_name, i + 1)
                self.ods_software.append(sw_info)

            # 加载特殊转换配置 - PDF转Office
            pdf_to_office_priority = special_conversions.get("pdf_to_office", ["msoffice_word", "libreoffice"])
            self.pdf_to_office_software = []
            for i, software_id in enumerate(pdf_to_office_priority):
                display_name = self.software_names.get(software_id, software_id)
                sw_info = SoftwareInfo(software_id, display_name, i + 1)
                self.pdf_to_office_software.append(sw_info)

            # 加载特殊转换配置 - 文档转PDF
            document_to_pdf_priority = special_conversions.get("document_to_pdf", ["wps_writer", "msoffice_word", "libreoffice"])
            self.document_to_pdf_software = []
            for i, software_id in enumerate(document_to_pdf_priority):
                display_name = self.software_names.get(software_id, software_id)
                sw_info = SoftwareInfo(software_id, display_name, i + 1)
                self.document_to_pdf_software.append(sw_info)

            # 加载特殊转换配置 - 表格转PDF
            spreadsheet_to_pdf_priority = special_conversions.get("spreadsheet_to_pdf", ["wps_spreadsheets", "msoffice_excel", "libreoffice"])
            self.spreadsheet_to_pdf_software = []
            for i, software_id in enumerate(spreadsheet_to_pdf_priority):
                display_name = self.software_names.get(software_id, software_id)
                sw_info = SoftwareInfo(software_id, display_name, i + 1)
                self.spreadsheet_to_pdf_software.append(sw_info)

            # 默认选中第1行第1个软件（文档处理软件的第1个）
            if self.word_processors:
                self.word_processors[0].is_selected = True
                self.selected_software = self.word_processors[0]
                self.selected_category = 'word'

            logger.debug("软件优先级配置数据加载完成")

        except Exception as e:
            logger.error(f"加载软件优先级配置失败: {str(e)}")

    def _create_interface(self):
        """
        创建选项卡界面
        
        创建三个主要设置区域：
        1. 默认优先级配置 - 文档处理和表格处理软件优先级
        2. 特殊转换配置 - ODT、ODS和PDF转换的软件优先级
        3. 软件说明 - 使用说明和软件介绍
        """
        logger.debug("开始创建软件优先级设置界面")

        self._create_default_priority_section()
        self._create_special_conversions_section()
        self._create_software_info_section()

        logger.debug("软件优先级设置界面创建完成")

    def _create_default_priority_section(self):
        """
        创建默认优先级配置区域
        
        包含文档处理软件和表格处理软件的优先级设置。
        
        配置路径：
        - software_priority.default_priority.word_processors
        - software_priority.default_priority.spreadsheet_processors
        """
        logger.debug("创建默认优先级配置区域")

        frame = self.create_section_frame(
            self.scrollable_frame,
            "默认软件优先级",
            SectionStyle.PRIMARY
        )

        # 文档处理软件优先级
        doc_label = tb.Label(
            frame,
            text="文档处理软件优先级",
            font=(self.small_font, self.small_size, "bold")
        )
        doc_label.pack(anchor="w", pady=(0, 10))

        # 文档处理软件卡片容器
        self.word_cards_frame = tb.Frame(frame)
        self.word_cards_frame.pack(fill="x", pady=(0, 15))

        # 表格处理软件优先级
        sheet_label = tb.Label(
            frame,
            text="表格处理软件优先级",
            font=(self.small_font, self.small_size, "bold")
        )
        sheet_label.pack(anchor="w", pady=(0, 10))

        # 表格处理软件卡片容器
        self.spreadsheet_cards_frame = tb.Frame(frame)
        self.spreadsheet_cards_frame.pack(fill="x")

        logger.debug("默认优先级配置区域创建完成")

    def _create_special_conversions_section(self):
        """
        创建特殊转换配置区域
        
        包含ODT、ODS和PDF转换的软件优先级设置。
        
        配置路径：
        - software_priority.special_conversions.odt
        - software_priority.special_conversions.ods
        - software_priority.special_conversions.pdf_to_office
        - software_priority.special_conversions.office_to_pdf
        """
        logger.debug("创建特殊转换配置区域")

        frame = self.create_section_frame(
            self.scrollable_frame,
            "特殊格式转换优先级",
            SectionStyle.INFO
        )

        # ODT格式转换
        odt_label = tb.Label(
            frame,
            text="ODT格式转换",
            font=(self.small_font, self.small_size, "bold")
        )
        odt_label.pack(anchor="w", pady=(0, 10))

        self.odt_cards_frame = tb.Frame(frame)
        self.odt_cards_frame.pack(fill="x", pady=(0, 15))

        # ODS格式转换
        ods_label = tb.Label(
            frame,
            text="ODS格式转换",
            font=(self.small_font, self.small_size, "bold")
        )
        ods_label.pack(anchor="w", pady=(0, 10))

        self.ods_cards_frame = tb.Frame(frame)
        self.ods_cards_frame.pack(fill="x", pady=(0, 15))

        # PDF转文档
        pdf_to_office_label = tb.Label(
            frame,
            text="PDF转文档",
            font=(self.small_font, self.small_size, "bold")
        )
        pdf_to_office_label.pack(anchor="w", pady=(0, 10))

        self.pdf_to_office_cards_frame = tb.Frame(frame)
        self.pdf_to_office_cards_frame.pack(fill="x", pady=(0, 15))

        # 文档转PDF
        document_to_pdf_label = tb.Label(
            frame,
            text="文档转PDF",
            font=(self.small_font, self.small_size, "bold")
        )
        document_to_pdf_label.pack(anchor="w", pady=(0, 10))

        self.document_to_pdf_cards_frame = tb.Frame(frame)
        self.document_to_pdf_cards_frame.pack(fill="x", pady=(0, 15))

        # 表格转PDF
        spreadsheet_to_pdf_label = tb.Label(
            frame,
            text="表格转PDF",
            font=(self.small_font, self.small_size, "bold")
        )
        spreadsheet_to_pdf_label.pack(anchor="w", pady=(0, 10))

        self.spreadsheet_to_pdf_cards_frame = tb.Frame(frame)
        self.spreadsheet_to_pdf_cards_frame.pack(fill="x")

        logger.debug("特殊转换配置区域创建完成")

    def _create_software_info_section(self):
        """
        创建软件说明区域
        
        提供软件功能说明和使用指南。
        """
        logger.debug("创建软件说明区域")

        frame = self.create_section_frame(
            self.scrollable_frame,
            "软件说明",
            SectionStyle.SECONDARY
        )

        # 说明文本
        info_text = """
软件优先级决定了格式转换时调用外部软件的先后顺序。
系统会按照从左到右的优先级顺序，尝试使用软件，直到找到可用的软件为止。
        """

        info_label = tb.Label(
            frame,
            text=info_text.strip(),
            font=(self.small_font, self.small_size),
            justify="left"
        )
        info_label.pack(anchor="w", fill="x", expand=True)

        # 动态调整换行宽度以适应父容器
        def update_wraplength(event):
            """根据父容器宽度动态更新换行宽度"""
            container_width = event.width
            # 减去一些边距以确保文本不会紧贴边框
            wraplength = max(container_width - 20, 100)  # 最小宽度100像素
            info_label.configure(wraplength=wraplength)

        # 绑定到父容器的尺寸变化事件
        frame.bind("<Configure>", update_wraplength)

        logger.debug("软件说明区域创建完成")

    def _post_initialize(self):
        """
        初始化后处理
        
        先创建所有卡片，再绑定滚轮事件，确保卡片能正确响应滚轮。
        """
        logger.debug("执行软件优先级选项卡后置处理")
        
        # 先刷新所有类别的UI显示（创建所有卡片）
        self._refresh_all_categories()
        
        # 再调用基类的后置处理（绑定滚轮事件）
        super()._post_initialize()
        
        logger.debug("软件优先级选项卡后置处理完成")

    def _create_software_card(self, parent, software_info: SoftwareInfo, category: str, index: int, list_len: int):
        """
        创建软件卡片（自定义UI实现）
        
        创建一个带淡色边框的软件卡片，选中时显示移动按钮。
        
        参数：
            parent: 父组件
            software_info: 软件信息对象
            category: 类别标识（用于选中回调）
            index: 当前软件在列表中的索引
            list_len: 列表总长度
        """
        # 统一使用 tk.Frame 支持边框
        card = tk.Frame(parent)
        
        try:
            # 获取主题颜色
            style = tb.Style.get_instance()
            parent_bg = style.colors.bg
            
            if software_info.is_selected:
                # 选中状态：彩色加粗边框
                primary_color = style.colors.primary
                card.configure(
                    bg=parent_bg,
                    highlightbackground=primary_color,
                    highlightcolor=primary_color,
                    highlightthickness=scale(2)
                )
            else:
                # 未选中状态：淡灰色细边框
                card.configure(
                    bg=parent_bg,
                    highlightbackground="#DDDDDD",
                    highlightcolor="#DDDDDD",
                    highlightthickness=scale(1)
                )
        except Exception as e:
            logger.debug(f"设置边框颜色失败，使用默认: {e}")
            card.configure(
                bg='SystemButtonFace',
                highlightthickness=scale(1 if not software_info.is_selected else 2)
            )
        
        card.pack(side="left", padx=scale(3))
        
        # 根据是否选中创建不同的内容
        if software_info.is_selected:
            # 选中状态：按钮 + 名称 + 按钮（水平布局）
            
            # 左移按钮（如果不是第一个）
            if index > 0:
                left_btn = tk.Button(
                    card,
                    text="◀",
                    width=1,
                    command=lambda: self._move_software('left')
                )
                try:
                    style = tb.Style.get_instance()
                    primary_color = style.colors.primary
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
                except Exception as e:
                    logger.debug(f"设置按钮样式失败: {e}")
                    left_btn.configure(relief=tk.FLAT, borderwidth=0)
                left_btn.pack(side="left")
            
            # 软件名称标签
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
            
            # 右移按钮（如果不是最后一个）
            if index < list_len - 1:
                right_btn = tk.Button(
                    card,
                    text="▶",
                    width=1,
                    command=lambda: self._move_software('right')
                )
                try:
                    style = tb.Style.get_instance()
                    primary_color = style.colors.primary
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
                except Exception as e:
                    logger.debug(f"设置按钮样式失败: {e}")
                    right_btn.configure(relief=tk.FLAT, borderwidth=0)
                right_btn.pack(side="left")
            
            # 绑定点击事件到名称标签
            name_label.bind("<Button-1>", lambda e, si=software_info, cat=category: self._select_software(si, cat))
        else:
            # 未选中状态：只显示名称
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
            
            # 绑定点击事件
            card.bind("<Button-1>", lambda e, si=software_info, cat=category: self._select_software(si, cat))
            name_label.bind("<Button-1>", lambda e, si=software_info, cat=category: self._select_software(si, cat))

    def _select_software(self, software_info: SoftwareInfo, category: str):
        """
        选中指定软件（全局单选）
        
        清除所有软件的选中状态，然后选中指定软件。
        
        参数：
            software_info: 要选中的软件信息
            category: 类别标识
        """
        logger.debug(f"选中软件: {software_info.display_name} (类别: {category})")
        
        # 清除所有类别中所有软件的选中状态
        for sw in self.word_processors:
            sw.is_selected = False
        for sw in self.spreadsheet_processors:
            sw.is_selected = False
        for sw in self.odt_software:
            sw.is_selected = False
        for sw in self.ods_software:
            sw.is_selected = False
        for sw in self.pdf_to_office_software:
            sw.is_selected = False
        for sw in self.document_to_pdf_software:
            sw.is_selected = False
        for sw in self.spreadsheet_to_pdf_software:
            sw.is_selected = False
        
        # 选中当前软件
        software_info.is_selected = True
        self.selected_software = software_info
        self.selected_category = category
        
        # 刷新所有类别的显示
        self._refresh_all_categories()

    def _refresh_all_categories(self):
        """刷新所有类别的UI显示"""
        self._refresh_category('word')
        self._refresh_category('spreadsheet')
        self._refresh_category('odt')
        self._refresh_category('ods')
        self._refresh_category('pdf_to_office')
        self._refresh_category('document_to_pdf')
        self._refresh_category('spreadsheet_to_pdf')

    def _refresh_category(self, category: str):
        """
        刷新指定类别的UI显示
        
        清空容器并重新创建所有软件卡片。
        
        参数：
            category: 类别标识
        """
        # 根据类别获取对应的容器和列表
        if category == 'word':
            cards_frame = self.word_cards_frame
            software_list = self.word_processors
        elif category == 'spreadsheet':
            cards_frame = self.spreadsheet_cards_frame
            software_list = self.spreadsheet_processors
        elif category == 'odt':
            cards_frame = self.odt_cards_frame
            software_list = self.odt_software
        elif category == 'ods':
            cards_frame = self.ods_cards_frame
            software_list = self.ods_software
        elif category == 'pdf_to_office':
            cards_frame = self.pdf_to_office_cards_frame
            software_list = self.pdf_to_office_software
        elif category == 'document_to_pdf':
            cards_frame = self.document_to_pdf_cards_frame
            software_list = self.document_to_pdf_software
        elif category == 'spreadsheet_to_pdf':
            cards_frame = self.spreadsheet_to_pdf_cards_frame
            software_list = self.spreadsheet_to_pdf_software
        else:
            return
        
        # 清空容器
        for widget in cards_frame.winfo_children():
            widget.destroy()
        
        # 重新创建卡片
        list_len = len(software_list)
        for i, sw in enumerate(software_list):
            self._create_software_card(cards_frame, sw, category, i, list_len)

    def _move_software(self, direction: str):
        """
        移动软件位置
        
        在列表中交换选中软件的位置，然后刷新UI。
        
        参数：
            direction: 移动方向（'left' 或 'right'）
        """
        if not self.selected_software or not self.selected_category:
            return
        
        category = self.selected_category
        
        # 根据类别获取对应的列表
        if category == 'word':
            software_list = self.word_processors
        elif category == 'spreadsheet':
            software_list = self.spreadsheet_processors
        elif category == 'odt':
            software_list = self.odt_software
        elif category == 'ods':
            software_list = self.ods_software
        elif category == 'pdf_to_office':
            software_list = self.pdf_to_office_software
        elif category == 'office_to_pdf':
            software_list = self.office_to_pdf_software
        else:
            return
        
        # 获取当前索引
        index = software_list.index(self.selected_software)
        
        # 根据方向移动
        if direction == 'left' and index > 0:
            # 左移（交换位置）
            software_list[index], software_list[index - 1] = software_list[index - 1], software_list[index]
        elif direction == 'right' and index < len(software_list) - 1:
            # 右移（交换位置）
            software_list[index], software_list[index + 1] = software_list[index + 1], software_list[index]
        
        # 更新优先级编号
        for i, sw in enumerate(software_list):
            sw.priority = i + 1
        
        # 只刷新当前类别的显示
        self._refresh_category(category)
        
        logger.debug(f"移动软件完成: {category} - {direction}")

    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置
        
        返回：
            Dict[str, Any]: 当前所有软件优先级设置
        """
        settings = {
            "default_priority": {
                "word_processors": [sw.software_id for sw in self.word_processors],
                "spreadsheet_processors": [sw.software_id for sw in self.spreadsheet_processors]
            },
            "special_conversions": {
                "odt": [sw.software_id for sw in self.odt_software],
                "ods": [sw.software_id for sw in self.ods_software],
                "pdf_to_office": [sw.software_id for sw in self.pdf_to_office_software],
                "document_to_pdf": [sw.software_id for sw in self.document_to_pdf_software],
                "spreadsheet_to_pdf": [sw.software_id for sw in self.spreadsheet_to_pdf_software]
            }
        }
        
        logger.debug(f"获取软件优先级设置: {settings}")
        return settings

    def apply_settings(self) -> bool:
        """
        应用当前设置到配置文件
        
        将所有软件优先级设置保存到对应的配置路径。
        
        返回：
            bool: 应用是否成功
        """
        logger.debug("开始应用软件优先级设置到配置文件")

        try:
            settings = self.get_settings()
            success = True
            
            # 更新默认优先级
            default_priority = settings["default_priority"]
            if default_priority:
                result = self.config_manager.update_config_section(
                    "software_priority", "default_priority", default_priority
                )
                if not result:
                    success = False

            # 更新特殊转换配置
            special_conversions = settings["special_conversions"]
            if special_conversions:
                result = self.config_manager.update_config_section(
                    "software_priority", "special_conversions", special_conversions
                )
                if not result:
                    success = False

            if success:
                logger.info("✓ 软件优先级设置已成功应用到配置文件")
            else:
                logger.error("✗ 部分软件优先级设置更新失败")

            return success

        except Exception as e:
            logger.error(f"应用软件优先级设置失败: {str(e)}", exc_info=True)
            return False
