"""
文档设置选项卡模块

实现设置对话框的文档设置选项卡，包含：
- 提取/OCR默认设置
- 校对选项开关
- 词库配置按钮
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
    
    def _create_interface(self):
        """创建选项卡界面"""
        logger.debug("开始创建文档设置选项卡界面")
        
        self._create_extraction_section()
        self._create_validation_section()
        self._create_dictionary_section()
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
            "提取/OCR设置",
            SectionStyle.PRIMARY
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
    
    def _create_validation_section(self):
        """创建校对设置区域"""
        logger.debug("创建校对设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "校对设置",
            SectionStyle.INFO
        )
        
        # 获取当前配置
        try:
            document_config = self.config_manager.get_document_defaults()
            symbol_pairing = document_config.get("enable_symbol_pairing", True)
            symbol_correction = document_config.get("enable_symbol_correction", True)
            typos_rule = document_config.get("enable_typos_rule", True)
            sensitive_word = document_config.get("enable_sensitive_word", True)
        except Exception as e:
            logger.warning(f"读取校对配置失败: {e}")
            symbol_pairing, symbol_correction = True, True
            typos_rule, sensitive_word = True, True
        
        # 标点配对
        self.symbol_pairing_var = tk.BooleanVar(value=symbol_pairing)
        self.create_checkbox_with_info(
            frame,
            "启用标点配对检查",
            self.symbol_pairing_var,
            "检查括号、引号等是否成对出现",
            lambda: self.on_change("enable_symbol_pairing", self.symbol_pairing_var.get())
        )
        
        # 符号校对
        self.symbol_correction_var = tk.BooleanVar(value=symbol_correction)
        self.create_checkbox_with_info(
            frame,
            "启用符号校对",
            self.symbol_correction_var,
            "自动纠正全角/半角符号",
            lambda: self.on_change("enable_symbol_correction", self.symbol_correction_var.get())
        )
        
        # 错别字校对
        self.typos_rule_var = tk.BooleanVar(value=typos_rule)
        self.create_checkbox_with_info(
            frame,
            "启用错别字校对",
            self.typos_rule_var,
            "根据自定义词典检查并纠正错别字",
            lambda: self.on_change("enable_typos_rule", self.typos_rule_var.get())
        )
        
        # 敏感词匹配
        self.sensitive_word_var = tk.BooleanVar(value=sensitive_word)
        self.create_checkbox_with_info(
            frame,
            "启用敏感词匹配",
            self.sensitive_word_var,
            "检查文本中是否包含需要关注的敏感词",
            lambda: self.on_change("enable_sensitive_word", self.sensitive_word_var.get())
        )
    
    def _create_dictionary_section(self):
        """创建词库配置区域"""
        logger.debug("创建词库配置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "词库配置",
            SectionStyle.WARNING
        )
        
        # 说明文本
        desc_label = tb.Label(
            frame,
            text="配置校对词库：错误符号、错别字、敏感词。",
            bootstyle="secondary",
            wraplength=400
        )
        desc_label.pack(anchor="w", pady=(0, 10))
        
        # 按钮容器
        button_frame = tb.Frame(frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # 三个编辑按钮
        tb.Button(
            button_frame,
            text="📝 错误符号",
            command=self._on_edit_symbol_mapping,
            bootstyle="info",
            width=12
        ).pack(side="left", padx=(0, 5))
        
        tb.Button(
            button_frame,
            text="📝 错别字",
            command=self._on_edit_custom_typos,
            bootstyle="warning",
            width=12
        ).pack(side="left", padx=(0, 5))
        
        tb.Button(
            button_frame,
            text="📝 敏感词",
            command=self._on_edit_sensitive_words,
            bootstyle="danger",
            width=12
        ).pack(side="left")
    
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
    
    # ========== 词库配置方法 ==========
    
    def _on_edit_symbol_mapping(self):
        """编辑符号映射"""
        logger.info("打开符号映射编辑器")
        config_file_path = self._get_config_file_path("symbol_settings.toml")
        from gongwen_converter.config.toml_operations import read_toml_file
        config_data = read_toml_file(config_file_path)
        symbol_map = config_data.get("symbol_map", {})
        self._open_editor("symbol", symbol_map, self._on_symbol_mapping_saved, config_file_path)
    
    def _on_edit_custom_typos(self):
        """编辑错别字"""
        logger.info("打开错别字映射编辑器")
        config_file_path = self._get_config_file_path("typos_settings.toml")
        from gongwen_converter.config.toml_operations import read_toml_file
        config_data = read_toml_file(config_file_path)
        custom_typos = config_data.get("typos", {})
        self._open_editor("typo", custom_typos, self._on_custom_typos_saved, config_file_path)
    
    def _on_edit_sensitive_words(self):
        """编辑敏感词"""
        logger.info("打开敏感词映射编辑器")
        config_file_path = self._get_config_file_path("sensitive_words.toml")
        from gongwen_converter.config.toml_operations import read_toml_file
        config_data = read_toml_file(config_file_path)
        sensitive_words = config_data.get("sensitive_words", {})
        self._open_editor("sensitive", sensitive_words, self._on_sensitive_words_saved, config_file_path)
    
    def _get_config_file_path(self, filename: str) -> str:
        """获取配置文件完整路径"""
        import os
        try:
            config_dir = self.config_manager._config_dir
            return os.path.join(config_dir, filename)
        except Exception as e:
            logger.error(f"获取配置文件路径失败: {e}")
            return None
    
    def _open_editor(self, editor_type, data, save_callback, config_file_path=None):
        """打开映射编辑器"""
        try:
            from .mapping_editor import MappingEditorDialog
            editor = MappingEditorDialog(
                self,
                editor_type,
                data,
                save_callback,
                config_file_path=config_file_path
            )
            self.wait_window(editor)
        except ImportError as e:
            logger.error(f"导入映射编辑器失败: {e}")
    
    def _on_symbol_mapping_saved(self, new_mapping: Dict[str, List[str]]):
        """保存符号映射"""
        logger.info("符号映射已保存")
        config_file_path = self._get_config_file_path("symbol_settings.toml")
        self._save_mapping_with_comments(config_file_path, "symbol_map", new_mapping)
    
    def _on_custom_typos_saved(self, new_mapping: Dict[str, List[str]]):
        """保存错别字"""
        logger.info("错别字映射已保存")
        config_file_path = self._get_config_file_path("typos_settings.toml")
        self._save_mapping_with_comments(config_file_path, "typos", new_mapping)
    
    def _on_sensitive_words_saved(self, new_mapping: Dict[str, List[str]]):
        """保存敏感词"""
        logger.info("敏感词映射已保存")
        config_file_path = self._get_config_file_path("sensitive_words.toml")
        self._save_mapping_with_comments(config_file_path, "sensitive_words", new_mapping)
    
    def _save_mapping_with_comments(self, filepath, section, mapping_data):
        """保存映射数据和备注"""
        try:
            from gongwen_converter.config.toml_operations import (
                save_mapping_with_comments,
                extract_inline_comments
            )
            comments_data = extract_inline_comments(filepath, section)
            success = save_mapping_with_comments(filepath, section, mapping_data, comments_data)
            if success:
                self.config_manager.reload_configs()
        except Exception as e:
            logger.error(f"保存映射失败: {e}")
    
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
    
    # ========== 事件处理方法 ==========
    
    def _on_doc_keep_changed(self):
        """处理提取图片设置变更"""
        value = self.doc_keep_var.get()
        logger.info(f"文档提取图片设置变更: {value}")
        self.on_change("to_md_keep_images", value)
    
    def _on_doc_ocr_changed(self):
        """处理OCR设置变更"""
        value = self.doc_ocr_var.get()
        logger.info(f"文档OCR设置变更: {value}")
        self.on_change("to_md_enable_ocr", value)
    
    # ========== 配置获取和应用方法 ==========
    
    def get_settings(self) -> Dict[str, Any]:
        """获取当前设置"""
        settings = {
            # 提取/OCR设置
            "to_md_keep_images": self.doc_keep_var.get(),
            "to_md_enable_ocr": self.doc_ocr_var.get(),
            # 校对设置
            "enable_symbol_pairing": self.symbol_pairing_var.get(),
            "enable_symbol_correction": self.symbol_correction_var.get(),
            "enable_typos_rule": self.typos_rule_var.get(),
            "enable_sensitive_word": self.sensitive_word_var.get(),
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
            
            # 保存提取/OCR和校对设置到file_defaults.toml
            file_defaults_keys = [
                "to_md_keep_images", "to_md_enable_ocr",
                "enable_symbol_pairing", "enable_symbol_correction",
                "enable_typos_rule", "enable_sensitive_word"
            ]
            for key in file_defaults_keys:
                if not self.config_manager.update_config_value(
                    "file_defaults", "document", key, settings[key]
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
