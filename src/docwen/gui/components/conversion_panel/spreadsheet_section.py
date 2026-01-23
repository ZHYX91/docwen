"""
表格类格式转换功能模块

提供表格类文件（XLSX/XLS/ODS/CSV/ET）的格式转换按钮和汇总功能。

主要功能：
- 创建表格格式转换按钮（转为其他表格格式）
- 创建另存为PDF/OFD按钮
- 创建表格汇总选项区域（按行/按列/按单元格汇总）

依赖：
- ConversionPanelBase: 提供按钮样式、颜色映射等公共属性
- config_manager: 读取汇总选项默认值

使用方式：
    此模块作为 Mixin 类被 ConversionPanel 继承，不应直接实例化。
"""

import logging
import tkinter as tk

import ttkbootstrap as tb

from docwen.utils.dpi_utils import scale
from docwen.utils.font_utils import get_small_font
from docwen.utils.gui_utils import ToolTip, create_info_icon
from docwen.i18n import t

logger = logging.getLogger(__name__)


class SpreadsheetSectionMixin:
    """
    表格类转换功能混入类
    
    提供表格类文件的格式转换按钮和汇总功能。
    需要与 ConversionPanelBase 一起使用。
    """
    
    def _create_spreadsheet_buttons(self):
        """
        创建表格类格式转换按钮
        
        布局：
        - 格式转换section: XLSX/XLS/ODS/CSV按钮（2行2列）
        - 另存为section: PDF/OFD按钮
        - 扩展section: 汇总选项
        """
        logger.debug("创建表格类格式按钮 - 拆分布局")
        
        # === 格式转换 section ===
        hint_label = tb.Label(
            self.conversion_container,
            text=t("conversion_panel.spreadsheet.convert_to_other_formats"),
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
        
        # 第一行：XLSX、XLS
        formats_row1 = ['XLSX', 'XLS']
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
        
        # 第二行：ODS、CSV
        formats_row2 = ['ODS', 'CSV']
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
            text=t("conversion_panel.spreadsheet.convert_to_layout"),
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
            text=t("conversion_panel.spreadsheet.save_as_pdf"),
            command=lambda: self._on_format_clicked("PDF"),
            bootstyle=self.button_colors['PDF'],
            **self.button_style_2col
        )
        pdf_button.grid(row=1, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        self.format_buttons['PDF'] = pdf_button
        ToolTip(pdf_button, t("conversion_panel.spreadsheet.save_as_pdf_tooltip"))
        
        # OFD按钮（禁用）
        ofd_button = tb.Button(
            self.saveas_container,
            text=t("conversion_panel.spreadsheet.save_as_ofd"),
            command=lambda: None,
            bootstyle='secondary',
            **self.button_style_2col,
            state='disabled'
        )
        ofd_button.grid(row=1, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        self.format_buttons['OFD'] = ofd_button
        ToolTip(ofd_button, t("conversion_panel.spreadsheet.save_as_ofd_tooltip"))
        
        # 更新按钮状态
        self._update_button_states()
        
        # 创建汇总section
        self._create_spreadsheet_merge_section()
        
        logger.info("表格类格式按钮创建完成")
    
    def _create_spreadsheet_merge_section(self):
        """创建表格汇总section"""
        logger.debug("创建表格汇总section")
        
        # 显示扩展框架
        self.extra_frame.config(text=t("conversion_panel.spreadsheet.merge_tables"))
        self.extra_frame.grid()
        
        # 清空容器
        for widget in self.extra_container.winfo_children():
            widget.destroy()
        
        # 配置布局
        self.extra_container.grid_rowconfigure(0, weight=0)
        self.extra_container.grid_rowconfigure(1, weight=0)
        self.extra_container.grid_rowconfigure(2, weight=0)
        self.extra_container.grid_rowconfigure(3, weight=0)
        self.extra_container.grid_columnconfigure(0, weight=1)
        
        # 说明文字
        hint_label = tb.Label(
            self.extra_container,
            text=t("conversion_panel.spreadsheet.merge_tables_hint"),
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, sticky="ew", padx=5, pady=(0, 5))
        
        # 汇总按钮
        self.merge_tables_button = tb.Button(
            self.extra_container,
            text=t("conversion_panel.spreadsheet.merge_tables_button"),
            command=self._on_merge_tables_clicked,
            bootstyle='info',
            **self.button_style_1col
        )
        self.merge_tables_button.grid(row=1, column=0, pady=scale(5))
        ToolTip(
            self.merge_tables_button,
            t("conversion_panel.spreadsheet.merge_tables_button_tooltip")
        )
        
        # 汇总选项边框
        merge_options_frame = tb.Labelframe(
            self.extra_container,
            text=t("conversion_panel.spreadsheet.merge_options"),
            bootstyle="info"
        )
        merge_options_frame.grid(row=2, column=0, sticky="ew", padx=scale(5), pady=scale(5))
        merge_options_frame.grid_rowconfigure(0, weight=1)
        merge_options_frame.grid_columnconfigure(0, weight=1)
        
        # 获取默认汇总模式
        default_merge_mode = 3
        if self.config_manager:
            try:
                default_merge_mode = self.config_manager.get_spreadsheet_merge_mode()
            except Exception as e:
                logger.warning(f"读取汇总模式默认值失败: {e}")
        
        self.merge_mode_var = tk.IntVar(value=default_merge_mode)
        
        # 单选按钮容器（单列布局）
        merge_radio_frame = tb.Frame(merge_options_frame, bootstyle="default")
        merge_radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        merge_radio_frame.grid_columnconfigure(0, weight=1)
        
        # 按行汇总 + 图标（第0行）
        row0_frame = tb.Frame(merge_radio_frame, bootstyle="default")
        row0_frame.grid(row=0, column=0, sticky="w", padx=scale(10), pady=scale(3))
        
        self.merge_by_row_radio = tb.Radiobutton(
            row0_frame,
            text=t("conversion_panel.spreadsheet.merge_by_row"),
            variable=self.merge_mode_var,
            value=1,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        )
        self.merge_by_row_radio.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        merge_row_info = create_info_icon(
            row0_frame,
            t("conversion_panel.spreadsheet.merge_by_row_tooltip"),
            bootstyle="info"
        )
        merge_row_info.pack(side=tk.LEFT)
        
        # 按列汇总 + 图标（第1行）
        row1_frame = tb.Frame(merge_radio_frame, bootstyle="default")
        row1_frame.grid(row=1, column=0, sticky="w", padx=scale(10), pady=scale(3))
        
        self.merge_by_column_radio = tb.Radiobutton(
            row1_frame,
            text=t("conversion_panel.spreadsheet.merge_by_column"),
            variable=self.merge_mode_var,
            value=2,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        )
        self.merge_by_column_radio.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        merge_col_info = create_info_icon(
            row1_frame,
            t("conversion_panel.spreadsheet.merge_by_column_tooltip"),
            bootstyle="info"
        )
        merge_col_info.pack(side=tk.LEFT)
        
        # 按单元格汇总 + 图标（第2行）
        row2_frame = tb.Frame(merge_radio_frame, bootstyle="default")
        row2_frame.grid(row=2, column=0, sticky="w", padx=scale(10), pady=scale(3))
        
        self.merge_by_cell_radio = tb.Radiobutton(
            row2_frame,
            text=t("conversion_panel.spreadsheet.merge_by_cell"),
            variable=self.merge_mode_var,
            value=3,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        )
        self.merge_by_cell_radio.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        merge_cell_info = create_info_icon(
            row2_frame,
            t("conversion_panel.spreadsheet.merge_by_cell_tooltip"),
            bootstyle="info"
        )
        merge_cell_info.pack(side=tk.LEFT)
        
        # 基准表格显示
        small_font, small_size = get_small_font()
        
        reference_table_frame = tb.Frame(self.extra_container, bootstyle="default")
        reference_table_frame.grid(row=3, column=0, sticky="ew", padx=scale(5), pady=(scale(10), scale(5)))
        reference_table_frame.grid_columnconfigure(0, weight=1)
        
        reference_label = tb.Label(
            reference_table_frame,
            text=t("conversion_panel.spreadsheet.selected_reference_table"),
            font=(small_font, small_size),
            bootstyle="warning",
            anchor=tk.CENTER
        )
        reference_label.grid(row=0, column=0, sticky="ew")
        
        self.reference_table_var = tk.StringVar(value=t("conversion_panel.spreadsheet.no_table_selected"))
        self.reference_table_label = tb.Label(
            reference_table_frame,
            textvariable=self.reference_table_var,
            font=(small_font, small_size),
            bootstyle="primary",
            anchor=tk.CENTER,
            wraplength=250
        )
        self.reference_table_label.grid(row=1, column=0, sticky="ew")
        
        logger.debug("表格汇总section创建完成")
    
    def _on_merge_mode_changed(self):
        """处理汇总模式变更事件"""
        if hasattr(self, 'merge_tables_button') and self.merge_tables_button:
            mode = self.merge_mode_var.get() if hasattr(self, 'merge_mode_var') else 0
            should_enable = (mode > 0)
            self.merge_tables_button.config(state="normal" if should_enable else "disabled")
            logger.debug(f"汇总按钮状态: {'启用' if should_enable else '禁用'} (模式={mode})")
    
    def _on_merge_tables_clicked(self):
        """处理汇总表格按钮点击事件"""
        if self.on_action and self.current_file_path:
            mode = self.merge_mode_var.get() if hasattr(self, 'merge_mode_var') else 0
            options = {"mode": mode}
            logger.info(f"执行汇总表格操作，模式: {mode}")
            self.on_action("merge_tables", self.current_file_path, options)
    
    def set_reference_table(self, file_name: str):
        """
        设置当前选中的基准表格文件名
        
        参数：
            file_name: 文件名（显示用）
        """
        if hasattr(self, 'reference_table_var') and self.reference_table_var:
            if file_name:
                self.reference_table_var.set(file_name)
                logger.debug(f"设置基准表格: {file_name}")
            else:
                self.reference_table_var.set(t("conversion_panel.spreadsheet.no_table_selected"))
                logger.debug("清空基准表格显示")
