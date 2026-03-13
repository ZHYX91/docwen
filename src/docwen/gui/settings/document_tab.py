"""
文档设置选项卡模块

实现设置对话框的文档设置选项卡，包含：
- 提取/OCR默认设置
- 优化设置
- DOCX转MD序号设置
- 文档处理软件优先级

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。
"""

from __future__ import annotations

import logging
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.gui.settings.base_tab import BaseSettingsTab
from docwen.gui.settings.config import SectionStyle
from docwen.i18n import t
from docwen.utils.dpi_utils import scale

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

    def __init__(self, parent, config_manager: Any, on_change: Callable[[str, Any], None]):
        """初始化文档设置选项卡"""
        # 软件名称映射（国际化）
        self.software_names = {
            "wps_writer": t("settings.document.software.wps_writer"),
            "msoffice_word": t("settings.document.software.msoffice_word"),
            "libreoffice": t("settings.document.software.libreoffice"),
        }

        # 软件列表存储
        self.word_processors: list[SoftwareInfo] = []
        self.odt_software: list[SoftwareInfo] = []
        self.document_to_pdf_software: list[SoftwareInfo] = []

        # 全局选中的软件
        self.selected_software: SoftwareInfo | None = None
        self.selected_category: str | None = None

        # 卡片容器引用
        self.word_cards_frame: tb.Frame | None = None
        self.odt_cards_frame: tb.Frame | None = None
        self.document_to_pdf_cards_frame: tb.Frame | None = None

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
                SoftwareInfo(sw_id, self.software_names.get(sw_id) or sw_id, i + 1)
                for i, sw_id in enumerate(word_processors)
            ]

            odt_priority = config_manager.get_special_conversion_priority("odt")
            self.odt_software = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id) or sw_id, i + 1)
                for i, sw_id in enumerate(odt_priority)
            ]

            doc_to_pdf_priority = config_manager.get_document_to_pdf_priority()
            self.document_to_pdf_software = [
                SoftwareInfo(sw_id, self.software_names.get(sw_id) or sw_id, i + 1)
                for i, sw_id in enumerate(doc_to_pdf_priority)
            ]

            # 默认选中第一个
            if self.word_processors:
                self.word_processors[0].is_selected = True
                self.selected_software = self.word_processors[0]
                self.selected_category = "word"

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
                "legal_standard": "法律条文标准",
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
            self.scrollable_frame, t("settings.document.extraction_section"), SectionStyle.DANGER
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
            t("settings.document.keep_images"),
            self.doc_keep_var,
            t("settings.document.keep_images_tooltip"),
            self._on_doc_keep_changed,
        )

        # OCR识别
        self.doc_ocr_var = tk.BooleanVar(value=doc_ocr)
        self.create_checkbox_with_info(
            frame,
            t("settings.document.enable_ocr"),
            self.doc_ocr_var,
            t("settings.document.enable_ocr_tooltip"),
            self._on_doc_ocr_changed,
        )

    def _create_optimization_section(self):
        """创建优化设置区域（仅在当前语言有可用优化类型时显示）"""
        logger.debug("创建优化设置区域")

        # 检查当前语言是否有可用的优化类型（从配置文件获取）
        optimization_types_dict = self.config_manager.get_localized_optimization_types(scope="document_to_md")
        self._doc_optimization_types = optimization_types_dict

        # 如果当前语言没有可用的优化类型，不创建此区域
        if not optimization_types_dict:
            logger.info("当前语言没有可用的优化类型，优化设置区域已跳过")
            # 创建默认变量以避免后续代码出错
            self.doc_enable_opt_var = tk.BooleanVar(value=False)
            self.doc_opt_type_var = tk.StringVar(value="")
            return

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.document.optimization_section"), SectionStyle.INFO
        )

        # 获取当前配置
        try:
            enable_opt = self.config_manager.get_docx_to_md_enable_optimization()
            opt_type_id = self.config_manager.get_docx_to_md_optimization_type()
        except Exception as e:
            logger.warning(f"读取优化配置失败: {e}")
            enable_opt = True
            opt_type_id = "gongwen"

        # 获取优化类型名称列表（使用本地化选项功能）
        optimization_types = list(optimization_types_dict.values())

        # 确定默认优化类型名称
        if opt_type_id in optimization_types_dict:
            opt_type = optimization_types_dict[opt_type_id]
        elif optimization_types:
            opt_type = optimization_types[0]
        else:
            opt_type = ""

        # 启用优化 + 优化类型（同一行，左右各50%）
        opt_frame = tb.Frame(frame)
        opt_frame.pack(fill="x")

        # 左侧 - 复选框（占50%）
        left_frame = tb.Frame(opt_frame)
        left_frame.pack(side="left", expand=True, anchor="w")

        self.doc_enable_opt_var = tk.BooleanVar(value=enable_opt)
        opt_checkbox = tb.Checkbutton(
            left_frame,
            text=t("settings.document.enable_optimization"),
            variable=self.doc_enable_opt_var,
            command=lambda: self.on_change("to_md_enable_optimization", self.doc_enable_opt_var.get()),
            bootstyle="round-toggle",
        )
        opt_checkbox.pack(side="left")

        # 右侧 - 标签 + 下拉框（占50%）
        right_frame = tb.Frame(opt_frame)
        right_frame.pack(side="left", expand=True, anchor="w")

        type_label = tb.Label(
            right_frame, text=t("settings.document.optimization_type_label"), font=(self.small_font, self.small_size)
        )
        type_label.pack(side="left", padx=(0, scale(10)))

        self.doc_opt_type_var = tk.StringVar(value=opt_type)
        self.doc_opt_type_combo = tb.Combobox(
            right_frame, textvariable=self.doc_opt_type_var, values=optimization_types, state="readonly", width=10
        )
        self.doc_opt_type_combo.pack(side="left")
        self.doc_opt_type_combo.bind("<<ComboboxSelected>>", lambda e: self._on_optimization_type_changed())

        logger.debug(f"优化设置区域已创建，可用类型: {optimization_types}")

    def _on_optimization_type_changed(self):
        optimization_types = getattr(self, "_doc_optimization_types", {}) or {}
        name_to_id = {name: type_id for type_id, name in optimization_types.items()}
        selected_id = name_to_id.get(self.doc_opt_type_var.get(), "")
        self.on_change("to_md_optimization_type", selected_id)

    def _create_software_priority_section(self):
        """创建软件优先级区域"""
        logger.debug("创建软件优先级区域")

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.document.software_section"), SectionStyle.SUCCESS
        )

        # 文档处理软件
        doc_label = tb.Label(
            frame, text=t("settings.document.word_processors_label"), font=(self.small_font, self.small_size, "bold")
        )
        doc_label.pack(anchor="w", pady=(0, scale(10)))

        self.word_cards_frame = tb.Frame(frame)
        self.word_cards_frame.pack(fill="x", pady=(0, scale(15)))

        # ODT格式转换
        odt_label = tb.Label(
            frame, text=t("settings.document.odt_conversion_label"), font=(self.small_font, self.small_size, "bold")
        )
        odt_label.pack(anchor="w", pady=(0, scale(10)))

        self.odt_cards_frame = tb.Frame(frame)
        self.odt_cards_frame.pack(fill="x", pady=(0, scale(15)))

        # 文档转PDF
        pdf_label = tb.Label(
            frame, text=t("settings.document.document_to_pdf_label"), font=(self.small_font, self.small_size, "bold")
        )
        pdf_label.pack(anchor="w", pady=(0, scale(10)))

        self.document_to_pdf_cards_frame = tb.Frame(frame)
        self.document_to_pdf_cards_frame.pack(fill="x")

    # ========== 软件优先级方法 ==========

    def _create_software_card_for_category(
        self, parent, software_info: SoftwareInfo, category: str, index: int, list_len: int
    ):
        """创建软件卡片（使用基类方法）"""
        self.create_software_card(
            parent=parent,
            software_id=software_info.software_id,
            display_name=software_info.display_name,
            is_selected=software_info.is_selected,
            index=index,
            list_len=list_len,
            on_select=lambda: self._select_software(software_info, category),
            on_move=self._move_software,
        )

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
        self._refresh_category("word")
        self._refresh_category("odt")
        self._refresh_category("document_to_pdf")

    def _refresh_category(self, category: str):
        """刷新指定类别"""
        if category == "word":
            cards_frame = self.word_cards_frame
            software_list = self.word_processors
        elif category == "odt":
            cards_frame = self.odt_cards_frame
            software_list = self.odt_software
        elif category == "document_to_pdf":
            cards_frame = self.document_to_pdf_cards_frame
            software_list = self.document_to_pdf_software
        else:
            return

        if not cards_frame:
            return

        # 清空并重建
        for widget in cards_frame.winfo_children():
            widget.destroy()

        list_len = len(software_list)
        for i, sw in enumerate(software_list):
            self._create_software_card_for_category(cards_frame, sw, category, i, list_len)

    def _move_software(self, direction: str):
        """移动软件位置"""
        if not self.selected_software or not self.selected_category:
            return

        if self.selected_category == "word":
            software_list = self.word_processors
        elif self.selected_category == "odt":
            software_list = self.odt_software
        elif self.selected_category == "document_to_pdf":
            software_list = self.document_to_pdf_software
        else:
            return

        index = software_list.index(self.selected_software)

        if direction == "left" and index > 0:
            software_list[index], software_list[index - 1] = software_list[index - 1], software_list[index]
        elif direction == "right" and index < len(software_list) - 1:
            software_list[index], software_list[index + 1] = software_list[index + 1], software_list[index]

        for i, sw in enumerate(software_list):
            sw.priority = i + 1

        self._refresh_category(self.selected_category)

    def _create_docx_to_md_numbering_section(self):
        """创建DOCX转MD序号设置区域"""
        logger.debug("创建DOCX转MD序号设置区域")

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.document.numbering.section"), SectionStyle.PRIMARY
        )

        # 获取当前配置
        try:
            remove_numbering = self.config_manager.get_docx_to_md_remove_numbering()
            add_numbering = self.config_manager.get_docx_to_md_add_numbering()
            default_scheme = self.config_manager.get_docx_to_md_default_scheme()
        except Exception as e:
            logger.warning(f"读取DOCX转MD序号配置失败: {e}")
            remove_numbering, add_numbering = True, False
            default_scheme = "hierarchical_standard"

        # 获取适用于当前语言的序号方案（从配置文件获取）
        self._docx_numbering_schemes = self.config_manager.get_localized_numbering_schemes()
        scheme_names = list(self._docx_numbering_schemes.values())
        scheme_id_to_name = self._docx_numbering_schemes

        # 确定默认序号方案名称（如果配置的方案不适用于当前语言，使用第一个可用方案）
        if default_scheme in scheme_id_to_name:
            default_scheme_name = scheme_id_to_name[default_scheme]
        elif scheme_names:
            default_scheme_name = scheme_names[0]
            logger.debug(f"配置的序号方案不适用于当前语言，使用: {next(iter(self._docx_numbering_schemes.keys()))}")
        else:
            default_scheme_name = ""
            logger.warning("当前语言没有可用的序号方案")

        # 默认清除原有文档小标题序号
        self.docx_to_md_remove_var = tk.BooleanVar(value=remove_numbering)
        self.create_checkbox_with_info(
            frame,
            t("settings.document.numbering.remove"),
            self.docx_to_md_remove_var,
            t("settings.document.numbering.remove_tooltip"),
            lambda: self.on_change("docx_to_md_remove_numbering", self.docx_to_md_remove_var.get()),
        )

        # 默认新增小标题序号到Markdown + 默认序号方案（同一行，左右各50%）
        add_scheme_frame = tb.Frame(frame)
        add_scheme_frame.pack(fill="x", pady=(scale(10), 0))

        # 左侧 - 复选框（占50%）
        left_frame = tb.Frame(add_scheme_frame)
        left_frame.pack(side="left", expand=True, anchor="w")

        self.docx_to_md_add_var = tk.BooleanVar(value=add_numbering)
        add_checkbox = tb.Checkbutton(
            left_frame,
            text=t("settings.document.numbering.add"),
            variable=self.docx_to_md_add_var,
            command=lambda: self.on_change("docx_to_md_add_numbering", self.docx_to_md_add_var.get()),
            bootstyle="round-toggle",
        )
        add_checkbox.pack(side="left")

        # 右侧 - 标签 + 下拉框（占50%）
        right_frame = tb.Frame(add_scheme_frame)
        right_frame.pack(side="left", expand=True, anchor="w")

        scheme_label = tb.Label(
            right_frame, text=t("settings.document.numbering.scheme_label"), font=(self.small_font, self.small_size)
        )
        scheme_label.pack(side="left", padx=(0, scale(10)))

        self.docx_to_md_scheme_var = tk.StringVar(value=default_scheme_name)
        self.docx_to_md_scheme_combo = tb.Combobox(
            right_frame, textvariable=self.docx_to_md_scheme_var, values=scheme_names, state="readonly", width=15
        )
        self.docx_to_md_scheme_combo.pack(side="left")
        self.docx_to_md_scheme_combo.bind("<<ComboboxSelected>>", lambda e: self._on_docx_to_md_scheme_changed())

    # ========== 事件处理方法 ==========

    def _on_docx_to_md_scheme_changed(self):
        """处理DOCX转MD序号方案变更"""
        scheme_name = self.docx_to_md_scheme_var.get()

        # 使用本地化方案映射（名称到ID的反向映射）
        scheme_name_to_id = {v: k for k, v in self._docx_numbering_schemes.items()}

        scheme_id = scheme_name_to_id.get(scheme_name, "hierarchical_standard")
        logger.info(f"DOCX转MD序号方案变更: {scheme_name} ({scheme_id})")
        self.on_change("docx_to_md_default_scheme", scheme_id)

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

    def get_settings(self) -> dict[str, Any]:
        """获取当前设置"""
        # 使用本地化方案映射（名称到ID的反向映射）
        scheme_name_to_id = {}
        if hasattr(self, "_docx_numbering_schemes") and self._docx_numbering_schemes:
            scheme_name_to_id = {v: k for k, v in self._docx_numbering_schemes.items()}

        optimization_types = getattr(self, "_doc_optimization_types", {}) or {}
        name_to_id = {name: type_id for type_id, name in optimization_types.items()}
        selected_opt_id = name_to_id.get(self.doc_opt_type_var.get(), next(iter(optimization_types.keys()), ""))

        settings = {
            # 提取/OCR设置
            "to_md_keep_images": self.doc_keep_var.get(),
            "to_md_enable_ocr": self.doc_ocr_var.get(),
            # 优化设置
            "to_md_enable_optimization": self.doc_enable_opt_var.get(),
            "to_md_optimization_type": selected_opt_id,
            # DOCX转MD序号设置
            "to_md_remove_numbering": self.docx_to_md_remove_var.get(),
            "to_md_add_numbering": self.docx_to_md_add_var.get(),
            "to_md_default_scheme": scheme_name_to_id.get(self.docx_to_md_scheme_var.get(), "hierarchical_standard"),
            # 软件优先级
            "word_processors": [sw.software_id for sw in self.word_processors],
            "odt": [sw.software_id for sw in self.odt_software],
            "document_to_pdf": [sw.software_id for sw in self.document_to_pdf_software],
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
                "to_md_keep_images",
                "to_md_enable_ocr",
                "to_md_enable_optimization",
                "to_md_optimization_type",
                "to_md_remove_numbering",
                "to_md_add_numbering",
                "to_md_default_scheme",
            ]
            for key in conversion_defaults_keys:
                if not self.config_manager.update_config_value("conversion_defaults", "document", key, settings[key]):
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
                "spreadsheet_to_pdf": self.config_manager.get_spreadsheet_to_pdf_priority(),
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
