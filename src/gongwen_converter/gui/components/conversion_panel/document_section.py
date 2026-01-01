"""
文档类格式转换功能模块

提供文档类文件（DOCX/DOC/ODT/RTF/PDF）的格式转换按钮和校对功能。

主要功能：
- 创建文档格式转换按钮（转为其他文档格式）
- 创建另存为PDF/OFD按钮
- 创建校对选项区域（标点配对、错别字、符号、敏感词）

依赖：
- ConversionPanelBase: 提供按钮样式、颜色映射等公共属性
- config_manager: 读取校对选项默认值

使用方式：
    此模块作为 Mixin 类被 ConversionPanel 继承，不应直接实例化。
"""

import logging
import tkinter as tk
from typing import Dict

import ttkbootstrap as tb

from gongwen_converter.utils.dpi_utils import scale
from gongwen_converter.utils.gui_utils import ToolTip, create_info_icon

logger = logging.getLogger(__name__)


class DocumentSectionMixin:
    """
    文档类转换功能混入类
    
    提供文档类文件的格式转换按钮和校对功能。
    需要与 ConversionPanelBase 一起使用。
    """
    
    def _create_document_buttons(self):
        """
        创建文档类格式转换按钮
        
        布局：
        - 格式转换section: DOCX/DOC/ODT/RTF按钮（2行2列）
        - 另存为section: PDF/OFD按钮
        - 扩展section: 校对选项
        """
        logger.debug("创建文档类格式按钮 - 拆分布局")
        
        # === 格式转换 section ===
        hint_label = tb.Label(
            self.conversion_container,
            text="转换为文档类型的其他格式",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置网格布局
        for col in range(2):
            self.conversion_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        for row in range(1, 3):
            self.conversion_container.grid_rowconfigure(row, weight=1)
        
        # 第一行：DOCX、DOC
        formats_row1 = ['DOCX', 'DOC']
        for idx, fmt in enumerate(formats_row1):
            button = tb.Button(
                self.conversion_container,
                text=fmt,
                command=lambda f=fmt: self._on_format_clicked(f),
                bootstyle=self.button_colors[fmt],
                **self.button_style_2col
            )
            button.grid(row=1, column=idx, padx=scale(5), pady=scale(5), sticky="ew")
            self.format_buttons[fmt] = button
            logger.debug(f"  创建按钮: {fmt} at (1, {idx})")
        
        # 第二行：ODT、RTF
        formats_row2 = ['ODT', 'RTF']
        for idx, fmt in enumerate(formats_row2):
            button = tb.Button(
                self.conversion_container,
                text=fmt,
                command=lambda f=fmt: self._on_format_clicked(f),
                bootstyle=self.button_colors[fmt],
                **self.button_style_2col
            )
            button.grid(row=2, column=idx, padx=scale(5), pady=scale(5), sticky="ew")
            self.format_buttons[fmt] = button
            logger.debug(f"  创建按钮: {fmt} at (2, {idx})")
        
        # === 另存为 section ===
        saveas_hint_label = tb.Label(
            self.saveas_container,
            text="转换为版式文件",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        saveas_hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        for col in range(2):
            self.saveas_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # PDF按钮
        pdf_button = tb.Button(
            self.saveas_container,
            text="另存为 PDF",
            command=lambda: self._on_format_clicked("PDF"),
            bootstyle=self.button_colors['PDF'],
            **self.button_style_2col
        )
        pdf_button.grid(row=1, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        self.format_buttons['PDF'] = pdf_button
        ToolTip(pdf_button, "需要本地Office软件转换为PDF版式文件")
        
        # OFD按钮（禁用）
        ofd_button = tb.Button(
            self.saveas_container,
            text="另存为 OFD",
            command=lambda: None,
            bootstyle='secondary',
            **self.button_style_2col,
            state='disabled'
        )
        ofd_button.grid(row=1, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        self.format_buttons['OFD'] = ofd_button
        ToolTip(ofd_button, "暂不支持")
        
        # 更新按钮状态
        self._update_button_states()
        
        # 创建校对section
        self._create_document_validation_section()
        
        logger.info("文档类格式按钮创建完成")
    
    def _create_document_validation_section(self):
        """创建文档校对section"""
        logger.debug("创建文档校对section")
        
        # 显示扩展框架
        self.extra_frame.config(text="校对文档")
        self.extra_frame.grid()
        
        # 清空容器
        for widget in self.extra_container.winfo_children():
            widget.destroy()
        
        # 配置布局
        self.extra_container.grid_rowconfigure(0, weight=0)
        self.extra_container.grid_rowconfigure(1, weight=0)
        self.extra_container.grid_rowconfigure(2, weight=0)
        self.extra_container.grid_columnconfigure(0, weight=1)
        
        # 说明文字
        hint_label = tb.Label(
            self.extra_container,
            text="根据选项对文档进行校对",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, sticky="ew", padx=5, pady=(0, 5))
        
        # 校对按钮
        self.validate_button = tb.Button(
            self.extra_container,
            text="🔍 校对",
            command=self._on_validate_clicked,
            bootstyle='info',
            **self.button_style_1col,
            state="disabled"
        )
        self.validate_button.grid(row=1, column=0, pady=scale(5))
        ToolTip(
            self.validate_button,
            "对文档进行校对，包括标点配对、错别字、符号错误和敏感词检测。需至少选择一个校对选项"
        )
        
        # 校对选项边框
        validation_options_frame = tb.Labelframe(
            self.extra_container,
            text="校对选项",
            bootstyle="info"
        )
        validation_options_frame.grid(row=2, column=0, sticky="ew", padx=scale(5), pady=scale(5))
        validation_options_frame.grid_rowconfigure(0, weight=1)
        validation_options_frame.grid_columnconfigure(0, weight=1)
        
        # 复选框容器
        checkbox_container = tb.Frame(validation_options_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        checkbox_container.grid_columnconfigure(0, weight=1)
        checkbox_container.grid_columnconfigure(1, weight=1)
        
        # 获取默认选项
        default_options = self._get_default_validation_options()
        
        # 校对选项配置
        options = [
            ("标点配对", "symbol_pairing", "检查文档中的标点符号配对情况，如括号、引号等是否成对出现"),
            ("错别字校对", "typos_rule", "检查文档中的错别字，根据用户自定义的错别字词库进行匹配"),
            ("符号校对", "symbol_correction", "检查文档中的标点符号使用规范，如全角和半角符号误用等"),
            ("敏感词匹配", "sensitive_word", "检查文档中的敏感词，根据用户自定义的敏感词库进行匹配")
        ]
        
        self.checkbox_vars = {}
        for i, (text, key, tooltip_text) in enumerate(options):
            var = tk.BooleanVar(value=default_options.get(key, False))
            self.checkbox_vars[key] = var
            
            option_frame = tb.Frame(checkbox_container, bootstyle="default")
            row, col = divmod(i, 2)
            option_frame.grid(row=row, column=col, sticky="", padx=scale(10), pady=scale(5))
            
            checkbox = tb.Checkbutton(
                option_frame,
                text=text,
                variable=var,
                command=self._on_validation_option_changed,
                bootstyle="round-toggle"
            )
            checkbox.pack(side=tk.LEFT, padx=(0, scale(5)))
            
            info_icon = create_info_icon(option_frame, tooltip_text, bootstyle="info")
            info_icon.pack(side=tk.LEFT)
        
        # 初始化按钮状态
        self._on_validation_option_changed()
        
        logger.debug("文档校对section创建完成")
    
    def _get_default_validation_options(self) -> Dict[str, bool]:
        """
        从配置文件获取默认校对选项
        
        返回：
            Dict[str, bool]: 校对选项字典
        """
        if not self.config_manager:
            return {
                "symbol_pairing": True,
                "symbol_correction": True,
                "typos_rule": True,
                "sensitive_word": False
            }
        
        try:
            engine_settings = self.config_manager.get_proofread_engine_config()
            return {
                "symbol_pairing": engine_settings.get("enable_symbol_pairing", True),
                "symbol_correction": engine_settings.get("enable_symbol_correction", True),
                "typos_rule": engine_settings.get("enable_typos_rule", True),
                "sensitive_word": engine_settings.get("enable_sensitive_word", True),
            }
        except Exception as e:
            logger.error(f"获取默认校对选项失败: {str(e)}")
            return {
                "symbol_pairing": True,
                "symbol_correction": True,
                "typos_rule": True,
                "sensitive_word": False
            }
    
    def _on_validation_option_changed(self):
        """处理校对选项变更事件"""
        if hasattr(self, 'validate_button') and self.validate_button:
            any_selected = any(var.get() for var in self.checkbox_vars.values())
            self.validate_button.config(state="normal" if any_selected else "disabled")
            logger.debug(f"校对按钮状态: {'启用' if any_selected else '禁用'}")
    
    def _on_validate_clicked(self):
        """处理校对按钮点击事件"""
        if self.on_action and self.current_file_path:
            options = {key: var.get() for key, var in self.checkbox_vars.items()}
            logger.info(f"执行校对操作，选项: {options}")
            self.on_action("validate", self.current_file_path, options)
