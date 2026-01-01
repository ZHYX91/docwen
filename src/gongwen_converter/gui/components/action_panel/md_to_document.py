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

from gongwen_converter.utils.dpi_utils import scale
from gongwen_converter.utils.gui_utils import ToolTip, create_info_icon

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
        self.status_var.set("准备处理文本文件 - 转为文档")
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
            text="📝 生成 DOCX",
            command=self._on_convert_docx_clicked,
            bootstyle=self.button_colors['primary'],
            **self.button_style_2
        )
        self.convert_docx_button.grid(row=0, column=0, padx=(0, scale(25)))
        ToolTip(self.convert_docx_button, "转换为Word/WPS文档格式（推荐）")
        
        self.convert_doc_button = tb.Button(
            first_row_frame,
            text="📝 生成 DOC",
            command=self._on_convert_doc_clicked,
            bootstyle=self.button_colors['info'],
            **self.button_style_2
        )
        self.convert_doc_button.grid(row=0, column=1)
        ToolTip(
            self.convert_doc_button,
            "需要通过本地安装的 WPS、Microsoft Office 或 LibreOffice 进行转换，"
            "用户可自行设置使用软件的优先级。"
        )
        
        # 第二行：生成ODT | 生成RTF
        second_row_frame = tb.Frame(self.button_container, bootstyle="default")
        second_row_frame.grid(row=1, column=0, pady=(0, scale(10)))
        
        self.convert_odt_button = tb.Button(
            second_row_frame,
            text="📝 生成 ODT",
            command=self._on_convert_odt_clicked,
            bootstyle=self.button_colors['success'],
            **self.button_style_2
        )
        self.convert_odt_button.grid(row=0, column=0, padx=(0, scale(25)))
        ToolTip(
            self.convert_odt_button,
            "需要通过本地安装的 Microsoft Office 或 LibreOffice 进行转换，"
            "用户可自行设置使用软件的优先级。"
        )
        
        self.convert_rtf_button = tb.Button(
            second_row_frame,
            text="📝 生成 RTF",
            command=self._on_convert_rtf_clicked,
            bootstyle=self.button_colors['warning'],
            **self.button_style_2
        )
        self.convert_rtf_button.grid(row=0, column=1)
        ToolTip(
            self.convert_rtf_button,
            "需要通过本地安装的 WPS、Microsoft Office 或 LibreOffice 进行转换，"
            "用户可自行设置使用软件的优先级。"
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
            text="生成选项",
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
            default_scheme_id = "gongwen_standard"
        
        # 获取序号方案名称列表
        try:
            scheme_names = self.config_manager.get_scheme_names()
            if not scheme_names:
                scheme_names = ["公文标准", "层级数字标准", "法律条文标准"]
        except Exception as e:
            logger.warning(f"获取序号方案名称列表失败: {e}")
            scheme_names = ["公文标准", "层级数字标准", "法律条文标准"]
        
        # 方案ID到名称的映射
        scheme_id_to_name = {
            "gongwen_standard": "公文标准",
            "hierarchical_standard": "层级数字标准",
            "legal_standard": "法律条文标准"
        }
        default_scheme_name = scheme_id_to_name.get(default_scheme_id, "公文标准")
        
        # ===== 第1行：清除原有 Markdown 小标题序号 =====
        self.md_remove_numbering_var = tk.BooleanVar(value=default_remove_numbering)
        remove_numbering_container = tb.Frame(checkbox_container, bootstyle="default")
        remove_numbering_container.grid(
            row=0, column=0, columnspan=2, sticky="w",
            padx=(scale(10), scale(10)), pady=(0, scale(5))
        )
        
        md_remove_numbering_check = tb.Checkbutton(
            remove_numbering_container,
            text="清除原有 Markdown 小标题序号",
            variable=self.md_remove_numbering_var,
            bootstyle="round-toggle"
        )
        md_remove_numbering_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        remove_numbering_info = create_info_icon(
            remove_numbering_container,
            "自动识别并去除Markdown中已有的标题序号",
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
            text="新增小标题序号",
            variable=self.md_add_numbering_var,
            command=self._on_md_add_numbering_toggle,
            bootstyle="round-toggle"
        )
        md_add_numbering_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        add_numbering_info = create_info_icon(
            add_numbering_container,
            "根据选择的序号方案为文档小标题添加序号",
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
            ("标点配对", "symbol_pairing", "检查文档中的标点符号配对情况，如括号、引号等是否成对出现"),
            ("错别字校对", "typos_rule", "检查文档中的错别字，根据用户自定义的错别字词库进行匹配"),
            ("符号校对", "symbol_correction", "检查文档中的标点符号使用规范，如全角和半角符号误用等"),
            ("敏感词匹配", "sensitive_word", "检查文档中的敏感词，根据用户自定义的敏感词库进行匹配")
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
