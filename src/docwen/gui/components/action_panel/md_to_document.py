"""
MD转文档功能模块

提供 Markdown 文件转换为文档格式的按钮和选项：
- 转换按钮：DOCX、DOC、ODT、RTF
- 生成选项：序号配置、校对选项

依赖：
- ActionPanelBase: 提供按钮样式、颜色映射等公共属性
- config_manager: 读取选项默认值

使用方式：
    此模块作为 Mixin 类被 ActionPanel 继承，不应直接实例化。
"""

import logging
import tkinter as tk
from typing import Dict

import ttkbootstrap as tb

from docwen.utils.dpi_utils import scale
from docwen.utils.gui_utils import ToolTip, create_info_icon
from docwen.i18n import t
from docwen.i18n.i18n_manager import I18nManager

logger = logging.getLogger(__name__)


class MdToDocumentMixin:
    """
    MD转文档功能混入类
    
    提供 Markdown 文件转换为文档格式的所有功能：
    - setup_for_md_to_document: 设置MD转文档模式
    - _create_md_to_document_buttons: 创建转换按钮
    - _create_md_to_document_options: 创建生成选项
    - 点击事件处理方法
    
    依赖基类属性：
        button_container: 按钮容器框架
        button_colors: 按钮颜色映射
        button_style_2: 按钮样式配置
        config_manager: 配置管理器
        on_action: 操作回调函数
    """
    
    def setup_for_md_to_document(self, file_path: str):
        """
        设置为MD转文档处理模式
        
        为MD文件显示文档转换按钮（DOCX/DOC/ODT/RTF）和校对选项。
        
        参数：
            file_path: MD文件路径
        """
        logger.debug(f"设置MD转文档处理模式: {file_path}")
        self.file_type = "docx"
        self.file_path = file_path
        self._clear_buttons()
        self._clear_options()
        self._create_md_to_document_buttons()
        self._create_md_to_document_options()
        self.status_var.set(t("action_panel.md_to_document.ready"))
        logger.info("MD转文档操作面板设置完成")
    
    def _create_md_to_document_buttons(self):
        """
        创建MD转文档格式的按钮
        
        显示两行按钮：
        - 第一行：生成DOCX（主色调）、生成DOC（蓝色调）
        - 第二行：生成ODT（绿色调）、生成RTF（橙色调）
        """
        logger.debug("创建MD到文档系列的按钮 - 两行布局")
        
        # 第一行：生成DOCX | 生成DOC
        first_row_frame = tb.Frame(self.button_container, bootstyle="default")
        first_row_frame.grid(row=0, column=0, pady=(0, scale(10)))
        
        self.convert_docx_button = tb.Button(
            first_row_frame,
            text=t("action_panel.md_to_document.generate_docx"),
            command=self._on_convert_docx_clicked,
            bootstyle=self.button_colors['primary'],
            **self.button_style_2
        )
        self.convert_docx_button.grid(row=0, column=0, padx=(0, scale(25)))
        ToolTip(self.convert_docx_button, t("action_panel.md_to_document.generate_docx_tooltip"))
        
        self.convert_doc_button = tb.Button(
            first_row_frame,
            text=t("action_panel.md_to_document.generate_doc"),
            command=self._on_convert_doc_clicked,
            bootstyle=self.button_colors['info'],
            **self.button_style_2
        )
        self.convert_doc_button.grid(row=0, column=1)
        ToolTip(
            self.convert_doc_button,
            t("action_panel.md_to_document.generate_doc_tooltip")
        )
        
        # 第二行：生成ODT | 生成RTF
        second_row_frame = tb.Frame(self.button_container, bootstyle="default")
        second_row_frame.grid(row=1, column=0, pady=(0, scale(10)))
        
        self.convert_odt_button = tb.Button(
            second_row_frame,
            text=t("action_panel.md_to_document.generate_odt"),
            command=self._on_convert_odt_clicked,
            bootstyle=self.button_colors['success'],
            **self.button_style_2
        )
        self.convert_odt_button.grid(row=0, column=0, padx=(0, scale(25)))
        ToolTip(
            self.convert_odt_button,
            t("action_panel.md_to_document.generate_odt_tooltip")
        )
        
        self.convert_rtf_button = tb.Button(
            second_row_frame,
            text=t("action_panel.md_to_document.generate_rtf"),
            command=self._on_convert_rtf_clicked,
            bootstyle=self.button_colors['warning'],
            **self.button_style_2
        )
        self.convert_rtf_button.grid(row=0, column=1)
        ToolTip(
            self.convert_rtf_button,
            t("action_panel.md_to_document.generate_rtf_tooltip")
        )
        
        logger.debug("MD到文档系列按钮创建完成 - 两行布局")
    
    def _create_md_to_document_options(self):
        """
        创建MD转文档的生成选项
        
        显示选项区域：
        - 序号选项：清除原有序号、新增序号（含方案下拉框）
        - 校对选项：标点配对、错别字校对、符号校对、敏感词匹配
        """
        logger.debug("创建MD转文档的生成选项")

        # 创建生成选项边框
        self.md_to_doc_options_frame = tb.Labelframe(
            self.button_container,
            text=t("action_panel.generation_options"),
            bootstyle="info"
        )
        self.md_to_doc_options_frame.grid(
            row=2, column=0, sticky="ew", padx=scale(20), pady=scale(10)
        )
        
        # 配置选项框架网格权重
        self.md_to_doc_options_frame.grid_rowconfigure(0, weight=1)
        self.md_to_doc_options_frame.grid_columnconfigure(0, weight=1)

        # 创建容器
        checkbox_container = tb.Frame(self.md_to_doc_options_frame)
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置行列权重
        for i in range(6):
            checkbox_container.grid_rowconfigure(i, weight=0)
        checkbox_container.grid_columnconfigure(0, weight=1)
        checkbox_container.grid_columnconfigure(1, weight=1)

        # 读取序号默认值
        try:
            default_remove_numbering = self.config_manager.get_md_to_docx_remove_numbering()
            default_add_numbering = self.config_manager.get_md_to_docx_add_numbering()
            default_scheme_id = self.config_manager.get_md_to_docx_default_scheme()
        except Exception as e:
            logger.warning(f"读取MD转DOCX序号配置失败，使用默认值: {e}")
            default_remove_numbering = True
            default_add_numbering = True
            default_scheme_id = "hierarchical_standard"
        
        # 获取适用于当前语言的序号方案（从配置文件获取，与文档转MD保持一致）
        self._md_numbering_schemes = self.config_manager.get_localized_numbering_schemes()
        scheme_names = list(self._md_numbering_schemes.values())
        scheme_id_to_name = self._md_numbering_schemes
        
        # 确定默认序号方案名称（如果配置的方案不适用于当前语言，使用第一个可用方案）
        if default_scheme_id in scheme_id_to_name:
            default_scheme_name = scheme_id_to_name[default_scheme_id]
        elif scheme_names:
            default_scheme_name = scheme_names[0]
            default_scheme_id = list(self._md_numbering_schemes.keys())[0]
            logger.debug(f"配置的序号方案不适用于当前语言，使用: {default_scheme_id}")
        else:
            default_scheme_name = ""
            logger.warning("当前语言没有可用的序号方案")
        
        # ===== 第1行：清除原有 Markdown 小标题序号 =====
        self.md_remove_numbering_var = tk.BooleanVar(value=default_remove_numbering)
        remove_numbering_container = tb.Frame(checkbox_container, bootstyle="default")
        remove_numbering_container.grid(
            row=0, column=0, columnspan=2, sticky="w",
            padx=(scale(10), scale(10)), pady=(0, scale(5))
        )
        
        md_remove_numbering_check = tb.Checkbutton(
            remove_numbering_container,
            text=t("action_panel.md_to_document.remove_existing_numbering"),
            variable=self.md_remove_numbering_var,
            bootstyle="round-toggle"
        )
        md_remove_numbering_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        remove_numbering_info = create_info_icon(
            remove_numbering_container,
            t("action_panel.md_to_document.remove_existing_numbering_tooltip"),
            bootstyle="info"
        )
        remove_numbering_info.pack(side=tk.LEFT)
        
        # ===== 第2行：新增小标题序号 + 下拉框 =====
        self.md_add_numbering_var = tk.BooleanVar(value=default_add_numbering)
        add_numbering_container = tb.Frame(checkbox_container, bootstyle="default")
        add_numbering_container.grid(
            row=1, column=0, columnspan=2, sticky="w",
            padx=(scale(10), scale(10)), pady=(0, scale(5))
        )
        
        md_add_numbering_check = tb.Checkbutton(
            add_numbering_container,
            text=t("action_panel.md_to_document.add_new_numbering"),
            variable=self.md_add_numbering_var,
            command=self._on_md_add_numbering_toggle,
            bootstyle="round-toggle"
        )
        md_add_numbering_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        add_numbering_info = create_info_icon(
            add_numbering_container,
            t("action_panel.md_to_document.add_new_numbering_tooltip"),
            bootstyle="info"
        )
        add_numbering_info.pack(side=tk.LEFT, padx=(0, scale(10)))
        
        # 下拉框
        self.md_numbering_scheme_var = tk.StringVar(value=default_scheme_name)
        self.md_numbering_scheme_combo = tb.Combobox(
            add_numbering_container,
            textvariable=self.md_numbering_scheme_var,
            values=scheme_names,
            state="disabled" if not default_add_numbering else "readonly",
            width=15
        )
        self.md_numbering_scheme_combo.pack(side=tk.LEFT)
        
        # ===== 第3行：分割线 =====
        separator = tb.Separator(checkbox_container, bootstyle="info")
        separator.grid(row=2, column=0, columnspan=2, sticky="ew", padx=0, pady=scale(20))
        
        # ===== 第4-5行：校对选项 =====
        default_options = self._get_default_proofread_options()
        options = [
            (t("action_panel.proofread.symbol_pairing"), "symbol_pairing", t("action_panel.proofread.symbol_pairing_tooltip")),
            (t("action_panel.proofread.typos_rule"), "typos_rule", t("action_panel.proofread.typos_rule_tooltip")),
            (t("action_panel.proofread.symbol_correction"), "symbol_correction", t("action_panel.proofread.symbol_correction_tooltip")),
            (t("action_panel.proofread.sensitive_word"), "sensitive_word", t("action_panel.proofread.sensitive_word_tooltip"))
        ]
        
        for i, (text, key, tooltip_text) in enumerate(options):
            var = tk.BooleanVar(value=default_options.get(key, False))
            self.checkbox_vars[key] = var
            
            option_frame = tb.Frame(checkbox_container, bootstyle="default")
            row, col = divmod(i, 2)
            option_frame.grid(row=row + 3, column=col, sticky="", padx=scale(10), pady=scale(5))
            
            checkbox = tb.Checkbutton(
                option_frame,
                text=text,
                variable=var,
                bootstyle="round-toggle"
            )
            checkbox.pack(side=tk.LEFT, padx=(0, scale(5)))
            
            info_icon = create_info_icon(option_frame, tooltip_text, bootstyle="info")
            info_icon.pack(side=tk.LEFT)
        
        logger.debug("MD转文档生成选项创建完成（含序号选项和校对选项）")
    
    def _get_default_proofread_options(self) -> Dict[str, bool]:
        """
        从配置文件获取默认校对选项
        
        返回：
            Dict[str, bool]: 校对选项字典
        """
        try:
            engine_settings = self.config_manager.get_proofread_engine_config()
            return {
                "symbol_pairing": engine_settings.get("enable_symbol_pairing", True),
                "symbol_correction": engine_settings.get("enable_symbol_correction", True),
                "typos_rule": engine_settings.get("enable_typos_rule", True),
                "sensitive_word": engine_settings.get("enable_sensitive_word", True),
            }
        except Exception as e:
            logger.error(f"获取默认选项失败: {str(e)}")
            return {
                "symbol_pairing": True,
                "symbol_correction": True,
                "typos_rule": True,
                "sensitive_word": False
            }
    
    def refresh_options(self):
        """刷新校对选项状态"""
        logger.debug("刷新选项状态以匹配当前配置")
        default_options = self._get_default_proofread_options()
        for key, var in self.checkbox_vars.items():
            if key in default_options:
                var.set(default_options[key])
                logger.debug(f"更新选项 {key} = {default_options[key]}")
        logger.info("选项状态已刷新")
    
    def _on_md_add_numbering_toggle(self):
        """
        处理MD转文档"添加标题序号"复选框切换事件
        
        控制序号方案下拉框的启用/禁用状态。
        """
        if hasattr(self, 'md_add_numbering_var') and hasattr(self, 'md_numbering_scheme_combo'):
            if self.md_add_numbering_var.get():
                self.md_numbering_scheme_combo.config(state="readonly")
                logger.debug("MD转文档：添加序号已启用，序号方案下拉框可选")
            else:
                self.md_numbering_scheme_combo.config(state="disabled")
                logger.debug("MD转文档：添加序号已禁用，序号方案下拉框灰色")
    
    def _on_convert_docx_clicked(self):
        """处理生成DOCX按钮点击事件"""
        if self.on_action:
            self.on_action("convert_md_to_docx", self.file_path, self.get_selected_options())
    
    def _on_convert_doc_clicked(self):
        """处理生成DOC按钮点击事件"""
        if self.on_action:
            self.on_action("convert_md_to_doc", self.file_path, self.get_selected_options())
    
    def _on_convert_odt_clicked(self):
        """处理生成ODT按钮点击事件"""
        if self.on_action:
            self.on_action("convert_md_to_odt", self.file_path, self.get_selected_options())
    
    def _on_convert_rtf_clicked(self):
        """处理生成RTF按钮点击事件"""
        if self.on_action:
            self.on_action("convert_md_to_rtf", self.file_path, self.get_selected_options())
            self.on_action("convert_md_to_rtf", self.file_path, self.get_selected_options())
            self.on_action("convert_md_to_rtf", self.file_path, self.get_selected_options())
