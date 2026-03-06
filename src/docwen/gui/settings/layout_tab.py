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


class LayoutTab(BaseSettingsTab):
    """
    版式设置选项卡类

    管理版式文件相关的所有配置选项。
    包含提取/OCR设置、渲染DPI设置和软件优先级。
    """

    def __init__(self, parent, config_manager: Any, on_change: Callable[[str, Any], None]):
        """初始化版式设置选项卡"""
        # 软件名称映射
        self.software_names = {"msoffice_word": "Word", "libreoffice": "LibreOffice"}

        # 软件列表
        self.pdf_to_office_software: list[SoftwareInfo] = []

        # 选中软件
        self.selected_software: SoftwareInfo | None = None
        self.selected_category: str | None = None

        # 卡片容器
        self.pdf_to_office_cards_frame: tb.Frame | None = None

        # 状态记录
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
                SoftwareInfo(sw_id, self.software_names.get(sw_id) or sw_id, i + 1)
                for i, sw_id in enumerate(pdf_to_office_priority)
            ]

            # 默认选中第一个
            if self.pdf_to_office_software:
                self.pdf_to_office_software[0].is_selected = True
                self.selected_software = self.pdf_to_office_software[0]
                self.selected_category = "pdf_to_office"

        except Exception as e:
            logger.error(f"加载版式配置失败: {e}")

    def _create_interface(self):
        """创建选项卡界面"""
        logger.debug("开始创建版式设置选项卡界面")

        self._create_extraction_section()
        self._create_optimization_section()
        self._create_dpi_section()
        self._create_software_priority_section()

        logger.debug("版式设置选项卡界面创建完成")

    def _post_initialize(self):
        """初始化后处理"""
        self._refresh_category("pdf_to_office")
        super()._post_initialize()

    def _create_extraction_section(self):
        """创建提取/OCR设置区域"""
        logger.debug("创建提取/OCR设置区域")

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.layout.extraction_section"), SectionStyle.DANGER
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
            self._on_layout_keep_changed,
        )

        # OCR识别
        self.layout_ocr_var = tk.BooleanVar(value=layout_ocr)
        self.create_checkbox_with_info(
            frame,
            t("settings.layout.enable_ocr"),
            self.layout_ocr_var,
            t("settings.layout.enable_ocr_tooltip"),
            self._on_layout_ocr_changed,
        )

        # 初始化状态记录
        self._layout_last_image_state = layout_keep
        self._layout_last_ocr_state = layout_ocr

    def _create_optimization_section(self):
        logger.debug("创建优化设置区域")

        optimization_types_dict = self.config_manager.get_localized_optimization_types(scope="layout_to_md")
        self._layout_optimization_types = optimization_types_dict

        if not optimization_types_dict:
            self.layout_enable_opt_var = tk.BooleanVar(value=False)
            self.layout_opt_type_var = tk.StringVar(value="")
            return

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.layout.optimization_section"), SectionStyle.INFO
        )

        try:
            enable_opt = self.config_manager.get_layout_to_md_enable_optimization()
            opt_type_id = self.config_manager.get_layout_to_md_optimization_type()
        except Exception as e:
            logger.warning(f"读取版式优化配置失败: {e}")
            enable_opt = False
            opt_type_id = "invoice_cn"

        optimization_types = list(optimization_types_dict.values())

        if opt_type_id in optimization_types_dict:
            opt_type = optimization_types_dict[opt_type_id]
        elif optimization_types:
            opt_type = optimization_types[0]
        else:
            opt_type = ""

        opt_frame = tb.Frame(frame)
        opt_frame.pack(fill="x")

        left_frame = tb.Frame(opt_frame)
        left_frame.pack(side="left", expand=True, anchor="w")

        self.layout_enable_opt_var = tk.BooleanVar(value=enable_opt)
        opt_checkbox = tb.Checkbutton(
            left_frame,
            text=t("settings.layout.enable_optimization"),
            variable=self.layout_enable_opt_var,
            command=lambda: self.on_change("to_md_enable_optimization", self.layout_enable_opt_var.get()),
            bootstyle="round-toggle",
        )
        opt_checkbox.pack(side="left")

        right_frame = tb.Frame(opt_frame)
        right_frame.pack(side="left", expand=True, anchor="w")

        type_label = tb.Label(
            right_frame,
            text=t("settings.layout.optimization_type_label"),
            font=(self.small_font, self.small_size),
        )
        type_label.pack(side="left", padx=(0, scale(10)))

        self.layout_opt_type_var = tk.StringVar(value=opt_type)
        self.layout_opt_type_combo = tb.Combobox(
            right_frame, textvariable=self.layout_opt_type_var, values=optimization_types, state="readonly", width=10
        )
        self.layout_opt_type_combo.pack(side="left")
        self.layout_opt_type_combo.bind("<<ComboboxSelected>>", lambda e: self._on_layout_optimization_type_changed())

    def _on_layout_optimization_type_changed(self):
        optimization_types = getattr(self, "_layout_optimization_types", {}) or {}
        name_to_id = {name: type_id for type_id, name in optimization_types.items()}
        selected_id = name_to_id.get(self.layout_opt_type_var.get(), "")
        self.on_change("to_md_optimization_type", selected_id)

    def _create_dpi_section(self):
        """创建渲染DPI设置区域"""
        logger.debug("创建渲染DPI设置区域")

        frame = self.create_section_frame(self.scrollable_frame, t("settings.layout.dpi_section"), SectionStyle.INFO)

        # 获取当前配置
        try:
            render_dpi = self.config_manager.get_layout_render_dpi()
        except Exception as e:
            logger.warning(f"读取渲染DPI失败: {e}")
            render_dpi = 300

        # 说明文本
        desc_label = tb.Label(frame, text=t("settings.layout.dpi_desc"), bootstyle="secondary")
        desc_label.pack(anchor="w", pady=(0, scale(10)))
        self.bind_label_wraplength(desc_label, frame, min_wraplength=scale(320))

        # 单选按钮变量
        self.render_dpi_var = tk.IntVar(value=render_dpi)

        # 单选按钮容器
        radio_frame = tb.Frame(frame)
        radio_frame.pack(fill="x", pady=(scale(10), 0))

        # 导入信息图标
        from docwen.utils.gui_utils import create_info_icon

        # 150 DPI
        dpi_150_frame = tb.Frame(radio_frame)
        dpi_150_frame.pack(fill="x", pady=(0, scale(8)))

        tb.Radiobutton(
            dpi_150_frame,
            text=t("settings.layout.dpi_min"),
            variable=self.render_dpi_var,
            value=150,
            command=self._on_dpi_changed,
            bootstyle="primary",
        ).pack(side="left")

        create_info_icon(dpi_150_frame, t("settings.layout.dpi_min_tooltip"), "info").pack(
            side="left", padx=(scale(5), 0)
        )

        # 300 DPI
        dpi_300_frame = tb.Frame(radio_frame)
        dpi_300_frame.pack(fill="x", pady=(0, scale(8)))

        tb.Radiobutton(
            dpi_300_frame,
            text=t("settings.layout.dpi_medium"),
            variable=self.render_dpi_var,
            value=300,
            command=self._on_dpi_changed,
            bootstyle="primary",
        ).pack(side="left")

        create_info_icon(dpi_300_frame, t("settings.layout.dpi_medium_tooltip"), "info").pack(
            side="left", padx=(scale(5), 0)
        )

        # 600 DPI
        dpi_600_frame = tb.Frame(radio_frame)
        dpi_600_frame.pack(fill="x")

        tb.Radiobutton(
            dpi_600_frame,
            text=t("settings.layout.dpi_high"),
            variable=self.render_dpi_var,
            value=600,
            command=self._on_dpi_changed,
            bootstyle="primary",
        ).pack(side="left")

        create_info_icon(dpi_600_frame, t("settings.layout.dpi_high_tooltip"), "info").pack(
            side="left", padx=(scale(5), 0)
        )

    def _create_software_priority_section(self):
        """创建软件优先级区域"""
        logger.debug("创建软件优先级区域")

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.layout.software_section"), SectionStyle.SUCCESS
        )

        # PDF转文档
        pdf_label = tb.Label(
            frame, text=t("settings.layout.pdf_to_doc_label"), font=(self.small_font, self.small_size, "bold")
        )
        pdf_label.pack(anchor="w", pady=(0, scale(10)))

        self.pdf_to_office_cards_frame = tb.Frame(frame)
        self.pdf_to_office_cards_frame.pack(fill="x")

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
        for sw in self.pdf_to_office_software:
            sw.is_selected = False

        software_info.is_selected = True
        self.selected_software = software_info
        self.selected_category = category

        self._refresh_category("pdf_to_office")

    def _refresh_category(self, category: str):
        """刷新指定类别"""
        if category == "pdf_to_office":
            cards_frame = self.pdf_to_office_cards_frame
            software_list = self.pdf_to_office_software
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

        software_list = self.pdf_to_office_software
        index = software_list.index(self.selected_software)

        if direction == "left" and index > 0:
            software_list[index], software_list[index - 1] = software_list[index - 1], software_list[index]
        elif direction == "right" and index < len(software_list) - 1:
            software_list[index], software_list[index + 1] = software_list[index + 1], software_list[index]

        for i, sw in enumerate(software_list):
            sw.priority = i + 1

        self._refresh_category("pdf_to_office")

    # ========== 事件处理 ==========

    def _on_layout_keep_changed(self):
        """
        处理提取图片设置变更
        """
        value = self.layout_keep_var.get()
        self._layout_last_image_state = value
        self._layout_last_ocr_state = self.layout_ocr_var.get()
        logger.info(f"版式提取图片设置变更: {value}")
        self.on_change("to_md_keep_images", value)

    def _on_layout_ocr_changed(self):
        """
        处理OCR设置变更
        """
        value = self.layout_ocr_var.get()
        self._layout_last_image_state = self.layout_keep_var.get()
        self._layout_last_ocr_state = value
        logger.info(f"版式OCR设置变更: {value}")
        self.on_change("to_md_enable_ocr", value)

    def _on_dpi_changed(self):
        """处理DPI设置变更"""
        value = self.render_dpi_var.get()
        logger.info(f"渲染DPI设置变更: {value}")
        self.on_change("render_dpi", value)

    # ========== 配置获取和应用 ==========

    def get_settings(self) -> dict[str, Any]:
        """获取当前设置"""
        optimization_types = getattr(self, "_layout_optimization_types", {}) or {}
        name_to_id = {name: type_id for type_id, name in optimization_types.items()}
        selected_opt_id = name_to_id.get(self.layout_opt_type_var.get(), next(iter(optimization_types.keys()), ""))

        settings = {
            "to_md_keep_images": self.layout_keep_var.get(),
            "to_md_enable_ocr": self.layout_ocr_var.get(),
            "to_md_enable_optimization": self.layout_enable_opt_var.get(),
            "to_md_optimization_type": selected_opt_id,
            "render_dpi": self.render_dpi_var.get(),
            "pdf_to_office": [sw.software_id for sw in self.pdf_to_office_software],
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
            conversion_defaults_keys = [
                "to_md_keep_images",
                "to_md_enable_ocr",
                "to_md_enable_optimization",
                "to_md_optimization_type",
                "render_dpi",
            ]
            for key in conversion_defaults_keys:
                if not self.config_manager.update_config_value("conversion_defaults", "layout", key, settings[key]):
                    success = False

            # 保存软件优先级
            special_conversions = {
                "odt": self.config_manager.get_special_conversion_priority("odt"),
                "ods": self.config_manager.get_special_conversion_priority("ods"),
                "pdf_to_office": settings["pdf_to_office"],
                "document_to_pdf": self.config_manager.get_document_to_pdf_priority(),
                "spreadsheet_to_pdf": self.config_manager.get_spreadsheet_to_pdf_priority(),
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
