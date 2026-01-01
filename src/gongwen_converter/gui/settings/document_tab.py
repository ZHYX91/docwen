"""
文档设置选项卡模块

实现设置对话框的文档设置选项卡，包含：
- 提取/OCR默认设置
- 优化设置
- DOCX转MD序号设置
- 文档处理软件优先级
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


class DocumentTab(BaseSettingsTab):
    """
    文档设置选项卡类
    
    管理文档文件相关的所有配置选项。
    包含提取/OCR设置、校对设置、词库配置和软件优先级。
    """
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
        """初始化文档设置选项卡"""
        # 软件名称映射
        self.software_names = {
            "wps_writer": "WPS文字",
            "msoffice_word": "Word",
            "libreoffice": "LibreOffice"
        }
        
        # 软件列表存储
        self.word_processors: List[SoftwareInfo] = []
        self.odt_software: List[SoftwareInfo] = []
        self.document_to_pdf_software: List[SoftwareInfo] = []
        
        # 全局选中的软件
        self.selected_software: Optional[SoftwareInfo] = None
        self.selected_category: Optional[str] = None
        
        # 卡片容器引用
        self.word_cards_frame: Optional[tb.Frame] = None
        self.odt_cards_frame: Optional[tb.Frame] = None
        self.document_to_pdf_cards_frame: Optional[tb.Frame] = None
        
        # 联动逻辑标志位和状态记录
        self._updating_doc_options: bool = False
        self._doc_last_image_state: bool = False
        self._doc_last_ocr_state: bool = False
        
        # 加载配置数据
        self._load_settings_data(config_manager)
        
        super().__init__(parent, config_manager, on_change)
        logger.info("文档设置选项卡初始化完成")
    
    def _load_settings_data(self, config_manager):
        """加载配置数据"""
        try:
            # 加载软件优先级
            word_processors = config_manager.get_word_processors_priority()
            self.word_processors = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id, sw_id), i + 1)
                for i, sw_id in enumerate(word_processors)
            ]
            
            odt_priority = config_manager.get_special_conversion_priority("odt")
            self.odt_software = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id, sw_id), i + 1)
                for i, sw_id in enumerate(odt_priority)
            ]
            
            doc_to_pdf_priority = config_manager.get_document_to_pdf_priority()
            self.document_to_pdf_software = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id, sw_id), i + 1)
                for i, sw_id in enumerate(doc_to_pdf_priority)
            ]
            
            # 默认选中第一个
            if self.word_processors:
                self.word_processors[0].is_selected = True
                self.selected_software = self.word_processors[0]
                self.selected_category = 'word'
                
        except Exception as e:
            logger.error(f"加载文档配置失败: {e}")
    
    def _get_scheme_mappings(self):
        """
        动态获取序号方案ID和名称的双向映射
        
        Returns:
            tuple: (id_to_name, name_to_id) 两个字典
        """
        id_to_name = {}
        try:
            schemes = self.config_manager.get_heading_schemes()
            for sid, sconfig in schemes.items():
                id_to_name[sid] = sconfig.get("name", sid)
        except Exception as e:
            logger.warning(f"动态获取序号方案映射失败: {e}")
            # 后备默认值
            id_to_name = {
                "gongwen_standard": "公文标准",
                "hierarchical_standard": "层级数字标准",
                "legal_standard": "法律条文标准"
            }
        name_to_id = {v: k for k, v in id_to_name.items()}
        return id_to_name, name_to_id
    
    def _create_interface(self):
        """创建选项卡界面"""
        logger.debug("开始创建文档设置选项卡界面")
        
        self._create_extraction_section()
        self._create_optimization_section()
        self._create_docx_to_md_numbering_section()
        self._create_software_priority_section()
        
        logger.debug("文档设置选项卡界面创建完成")
    
    def _post_initialize(self):
        """初始化后处理"""
        # 先刷新软件卡片
        self._refresh_all_categories()
        # 再绑定滚轮事件
        super()._post_initialize()
    
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
            doc_keep = self.config_manager.get_docx_to_md_keep_images()
            doc_ocr = self.config_manager.get_docx_to_md_enable_ocr()
        except Exception as e:
            logger.warning(f"读取文档配置失败: {e}")
            doc_keep, doc_ocr = True, False
        
        # 提取图片
        self.doc_keep_var = tk.BooleanVar(value=doc_keep)
        self.create_checkbox_with_info(
            frame,
            '默认勾选"提取图片"',
            self.doc_keep_var,
            '文档转MD时默认勾选"提取图片"',
            self._on_doc_keep_changed
        )
        
        # OCR识别
        self.doc_ocr_var = tk.BooleanVar(value=doc_ocr)
        self.create_checkbox_with_info(
            frame,
            '默认勾选"图片文字识别"',
            self.doc_ocr_var,
            '文档转MD时默认勾选"图片文字识别"',
            self._on_doc_ocr_changed
        )
        
        # 初始化状态记录
        self._doc_last_image_state = doc_keep
        self._doc_last_ocr_state = doc_ocr
    
    def _create_optimization_section(self):
        """创建优化设置区域"""
        logger.debug("创建优化设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "优化设置",
            SectionStyle.INFO
        )
        
        # 获取当前配置
        try:
            enable_opt = self.config_manager.get_docx_to_md_enable_optimization()
            opt_type = self.config_manager.get_docx_to_md_optimization_type()
        except Exception as e:
            logger.warning(f"读取优化配置失败: {e}")
            enable_opt, opt_type = True, "公文"
        
        # 启用优化 + 优化类型（同一行，左右各50%）
        opt_frame = tb.Frame(frame)
        opt_frame.pack(fill="x")
        
        # 左侧 - 复选框（占50%）
        left_frame = tb.Frame(opt_frame)
        left_frame.pack(side="left", expand=True, anchor="w")
        
        self.doc_enable_opt_var = tk.BooleanVar(value=enable_opt)
        opt_checkbox = tb.Checkbutton(
            left_frame,
            text='默认启用"针对文档类型优化"',
            variable=self.doc_enable_opt_var,
            command=lambda: self.on_change("to_md_enable_optimization", self.doc_enable_opt_var.get()),
            bootstyle="round-toggle"
        )
        opt_checkbox.pack(side="left")
        
        # 右侧 - 标签 + 下拉框（占50%）
        right_frame = tb.Frame(opt_frame)
        right_frame.pack(side="left", expand=True, anchor="w")
        
        type_label = tb.Label(
            right_frame,
            text="默认优化类型:",
            font=(self.small_font, self.small_size)
        )
        type_label.pack(side="left", padx=(0, 10))
        
        self.doc_opt_type_var = tk.StringVar(value=opt_type)
        self.doc_opt_type_combo = tb.Combobox(
            right_frame,
            textvariable=self.doc_opt_type_var,
            values=["公文", "合同", "论文"],
            state="readonly",
            width=10
        )
        self.doc_opt_type_combo.pack(side="left")
        self.doc_opt_type_combo.bind(
            '<<ComboboxSelected>>',
            lambda e: self.on_change("to_md_optimization_type", self.doc_opt_type_var.get())
        )
    
    
    def _create_software_priority_section(self):
        """创建软件优先级区域"""
        logger.debug("创建软件优先级区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "软件优先级",
            SectionStyle.SUCCESS
        )
        
        # 文档处理软件
        doc_label = tb.Label(
            frame,
            text="文档处理软件优先级",
            font=(self.small_font, self.small_size, "bold")
        )
        doc_label.pack(anchor="w", pady=(0, 10))
        
        self.word_cards_frame = tb.Frame(frame)
        self.word_cards_frame.pack(fill="x", pady=(0, 15))
        
        # ODT格式转换
        odt_label = tb.Label(
            frame,
            text="ODT格式转换",
            font=(self.small_font, self.small_size, "bold")
        )
        odt_label.pack(anchor="w", pady=(0, 10))
        
        self.odt_cards_frame = tb.Frame(frame)
        self.odt_cards_frame.pack(fill="x", pady=(0, 15))
        
        # 文档转PDF
        pdf_label = tb.Label(
            frame,
            text="文档转PDF",
            font=(self.small_font, self.small_size, "bold")
        )
        pdf_label.pack(anchor="w", pady=(0, 10))
        
        self.document_to_pdf_cards_frame = tb.Frame(frame)
        self.document_to_pdf_cards_frame.pack(fill="x")
    
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
            # 左移按钮
            if index > 0:
                left_btn = tk.Button(
                    card, text="◀", width=1,
                    command=lambda: self._move_software('left')
                )
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
            
            # 名称
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
            
            # 右移按钮
            if index < list_len - 1:
                right_btn = tk.Button(
                    card, text="▶", width=1,
                    command=lambda: self._move_software('right')
                )
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
        # 清除所有选中
        for sw in self.word_processors:
            sw.is_selected = False
        for sw in self.odt_software:
            sw.is_selected = False
        for sw in self.document_to_pdf_software:
            sw.is_selected = False
        
        # 选中当前
        software_info.is_selected = True
        self.selected_software = software_info
        self.selected_category = category
        
        self._refresh_all_categories()
    
    def _refresh_all_categories(self):
        """刷新所有类别"""
        self._refresh_category('word')
        self._refresh_category('odt')
        self._refresh_category('document_to_pdf')
    
    def _refresh_category(self, category: str):
        """刷新指定类别"""
        if category == 'word':
            cards_frame = self.word_cards_frame
            software_list = self.word_processors
        elif category == 'odt':
            cards_frame = self.odt_cards_frame
            software_list = self.odt_software
        elif category == 'document_to_pdf':
            cards_frame = self.document_to_pdf_cards_frame
            software_list = self.document_to_pdf_software
        else:
            return
        
        # 清空并重建
        for widget in cards_frame.winfo_children():
            widget.destroy()
        
        list_len = len(software_list)
        for i, sw in enumerate(software_list):
            self._create_software_card(cards_frame, sw, category, i, list_len)
    
    def _move_software(self, direction: str):
        """移动软件位置"""
        if not self.selected_software or not self.selected_category:
            return
        
        if self.selected_category == 'word':
            software_list = self.word_processors
        elif self.selected_category == 'odt':
            software_list = self.odt_software
        elif self.selected_category == 'document_to_pdf':
            software_list = self.document_to_pdf_software
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
    
    def _create_docx_to_md_numbering_section(self):
        """创建DOCX转MD序号设置区域"""
        logger.debug("创建DOCX转MD序号设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "小标题序号设置（文档转为MarkDown）",
            SectionStyle.PRIMARY
        )
        
        # 获取当前配置
        try:
            remove_numbering = self.config_manager.get_docx_to_md_remove_numbering()
            add_numbering = self.config_manager.get_docx_to_md_add_numbering()
            default_scheme = self.config_manager.get_docx_to_md_default_scheme()
        except Exception as e:
            logger.warning(f"读取DOCX转MD序号配置失败: {e}")
            remove_numbering, add_numbering = True, False
            default_scheme = "gongwen_standard"
        
        # 获取序号方案列表
        try:
            scheme_names = self.config_manager.get_scheme_names()
            if not scheme_names:
                scheme_names = ["公文标准", "层级数字标准", "法律条文标准"]
        except Exception as e:
            logger.warning(f"获取序号方案列表失败: {e}")
            scheme_names = ["公文标准", "层级数字标准", "法律条文标准"]
        
        # 动态获取方案ID到名称的映射
        scheme_id_to_name, _ = self._get_scheme_mappings()
        default_scheme_name = scheme_id_to_name.get(default_scheme, "公文标准")
        
        # 默认清除原有文档小标题序号
        self.docx_to_md_remove_var = tk.BooleanVar(value=remove_numbering)
        self.create_checkbox_with_info(
            frame,
            '默认清除原有文档小标题序号',
            self.docx_to_md_remove_var,
            'DOCX转MD时默认清除文档中已有的标题序号',
            lambda: self.on_change("docx_to_md_remove_numbering", self.docx_to_md_remove_var.get())
        )
        
        # 默认新增小标题序号到Markdown + 默认序号方案（同一行，左右各50%）
        add_scheme_frame = tb.Frame(frame)
        add_scheme_frame.pack(fill="x", pady=(10, 0))
        
        # 左侧 - 复选框（占50%）
        left_frame = tb.Frame(add_scheme_frame)
        left_frame.pack(side="left", expand=True, anchor="w")
        
        self.docx_to_md_add_var = tk.BooleanVar(value=add_numbering)
        add_checkbox = tb.Checkbutton(
            left_frame,
            text='默认新增小标题序号到Markdown',
            variable=self.docx_to_md_add_var,
            command=lambda: self.on_change("docx_to_md_add_numbering", self.docx_to_md_add_var.get()),
            bootstyle="round-toggle"
        )
        add_checkbox.pack(side="left")
        
        # 右侧 - 标签 + 下拉框（占50%）
        right_frame = tb.Frame(add_scheme_frame)
        right_frame.pack(side="left", expand=True, anchor="w")
        
        scheme_label = tb.Label(
            right_frame,
            text="默认序号方案:",
            font=(self.small_font, self.small_size)
        )
        scheme_label.pack(side="left", padx=(0, 10))
        
        self.docx_to_md_scheme_var = tk.StringVar(value=default_scheme_name)
        self.docx_to_md_scheme_combo = tb.Combobox(
            right_frame,
            textvariable=self.docx_to_md_scheme_var,
            values=scheme_names,
            state="readonly",
            width=15
        )
        self.docx_to_md_scheme_combo.pack(side="left")
        self.docx_to_md_scheme_combo.bind(
            '<<ComboboxSelected>>',
            lambda e: self._on_docx_to_md_scheme_changed()
        )
    
    # ========== 事件处理方法 ==========
    
    def _on_docx_to_md_scheme_changed(self):
        """处理DOCX转MD序号方案变更"""
        scheme_name = self.docx_to_md_scheme_var.get()
        
        # 动态获取名称到ID的映射
        _, scheme_name_to_id = self._get_scheme_mappings()
        
        scheme_id = scheme_name_to_id.get(scheme_name, "gongwen_standard")
        logger.info(f"DOCX转MD序号方案变更: {scheme_name} ({scheme_id})")
        self.on_change("docx_to_md_default_scheme", scheme_id)
    
    def _on_doc_keep_changed(self):
        """
        处理提取图片设置变更
        
        实现联动逻辑：
        1. 勾选OCR时，自动勾选"提取图片"
        2. 取消"提取图片"时，自动取消OCR
        """
        if self._updating_doc_options:
            return
        
        try:
            self._updating_doc_options = True
            
            extract_image = self.doc_keep_var.get()
            extract_ocr = self.doc_ocr_var.get()
            
            # 检测状态变化
            image_changed = (extract_image != self._doc_last_image_state)
            
            # 场景：用户取消提取图片（从有到无）
            if not extract_image and self._doc_last_image_state and image_changed:
                if extract_ocr:
                    logger.debug("文档设置：取消提取图片，自动取消OCR")
                    self.doc_ocr_var.set(False)
            
            # 更新状态记录
            self._doc_last_image_state = self.doc_keep_var.get()
            self._doc_last_ocr_state = self.doc_ocr_var.get()
            
            # 通知配置变更
            value = self.doc_keep_var.get()
            logger.info(f"文档提取图片设置变更: {value}")
            self.on_change("to_md_keep_images", value)
        
        finally:
            self._updating_doc_options = False
    
    def _on_doc_ocr_changed(self):
        """
        处理OCR设置变更
        
        实现联动逻辑：
        1. 勾选OCR时，自动勾选"提取图片"
        2. 取消"提取图片"时，自动取消OCR
        """
        if self._updating_doc_options:
            return
        
        try:
            self._updating_doc_options = True
            
            extract_image = self.doc_keep_var.get()
            extract_ocr = self.doc_ocr_var.get()
            
            # 检测状态变化
            ocr_changed = (extract_ocr != self._doc_last_ocr_state)
            
            # 场景：用户勾选OCR（从无到有）
            if extract_ocr and not self._doc_last_ocr_state and ocr_changed:
                if not extract_image:
                    logger.debug("文档设置：勾选OCR，自动勾选提取图片")
                    self.doc_keep_var.set(True)
            
            # 更新状态记录
            self._doc_last_image_state = self.doc_keep_var.get()
            self._doc_last_ocr_state = self.doc_ocr_var.get()
            
            # 通知配置变更
            value = self.doc_ocr_var.get()
            logger.info(f"文档OCR设置变更: {value}")
            self.on_change("to_md_enable_ocr", value)
        
        finally:
            self._updating_doc_options = False
    
    # ========== 配置获取和应用方法 ==========
    
    def get_settings(self) -> Dict[str, Any]:
        """获取当前设置"""
        # 动态获取方案名称到ID的映射
        _, scheme_name_to_id = self._get_scheme_mappings()
        
        settings = {
            # 提取/OCR设置
            "to_md_keep_images": self.doc_keep_var.get(),
            "to_md_enable_ocr": self.doc_ocr_var.get(),
            # 优化设置
            "to_md_enable_optimization": self.doc_enable_opt_var.get(),
            "to_md_optimization_type": self.doc_opt_type_var.get(),
            # DOCX转MD序号设置
            "to_md_remove_numbering": self.docx_to_md_remove_var.get(),
            "to_md_add_numbering": self.docx_to_md_add_var.get(),
            "to_md_default_scheme": scheme_name_to_id.get(
                self.docx_to_md_scheme_var.get(),
                "gongwen_standard"
            ),
            # 软件优先级
            "word_processors": [sw.software_id for sw in self.word_processors],
            "odt": [sw.software_id for sw in self.odt_software],
            "document_to_pdf": [sw.software_id for sw in self.document_to_pdf_software]
        }
        logger.debug(f"获取文档设置: {settings}")
        return settings
    
    def apply_settings(self) -> bool:
        """应用设置到配置文件"""
        logger.debug("开始应用文档设置")
        
        try:
            settings = self.get_settings()
            success = True
            
            # 保存提取/OCR、优化和序号设置到conversion_defaults.toml
            conversion_defaults_keys = [
                "to_md_keep_images", "to_md_enable_ocr",
                "to_md_enable_optimization", "to_md_optimization_type",
                "to_md_remove_numbering", "to_md_add_numbering", "to_md_default_scheme"
            ]
            for key in conversion_defaults_keys:
                if not self.config_manager.update_config_value(
                    "conversion_defaults", "document", key, settings[key]
                ):
                    success = False
            
            # 保存软件优先级到software_priority.toml
            if not self.config_manager.update_config_value(
                "software_priority", "default_priority", "word_processors", settings["word_processors"]
            ):
                success = False
            
            special_conversions = {
                "odt": settings["odt"],
                "ods": self.config_manager.get_special_conversion_priority("ods"),
                "pdf_to_office": self.config_manager.get_special_conversion_priority("pdf_to_office"),
                "document_to_pdf": settings["document_to_pdf"],
                "spreadsheet_to_pdf": self.config_manager.get_spreadsheet_to_pdf_priority()
            }
            
            if not self.config_manager.update_config_section(
                "software_priority", "special_conversions", special_conversions
            ):
                success = False
            
            if success:
                logger.info("✓ 文档设置已成功应用")
            else:
                logger.error("✗ 部分文档设置更新失败")
            
            return success
            
        except Exception as e:
            logger.error(f"应用文档设置失败: {e}", exc_info=True)
            return False
