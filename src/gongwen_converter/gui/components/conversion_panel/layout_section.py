"""
版式类格式转换功能模块

提供版式类文件（PDF/OFD/XPS/CEB）的格式转换按钮和合并拆分功能。

主要功能：
- 创建版式格式转换按钮（转为文档/图片格式）
- 创建合并拆分选项区域（按页码拆分/合并多个PDF）
- 创建渲染为图片选项（DPI选择）
- 页码范围解析和验证

依赖：
- ConversionPanelBase: 提供按钮样式、颜色映射等公共属性
- config_manager: 读取转换选项默认值

使用方式：
    此模块作为 Mixin 类被 ConversionPanel 继承，不应直接实例化。
"""

import logging
import re
import tkinter as tk
from typing import List

import ttkbootstrap as tb

from gongwen_converter.utils.dpi_utils import scale
from gongwen_converter.utils.font_utils import get_small_font
from gongwen_converter.utils.gui_utils import ToolTip, create_info_icon

logger = logging.getLogger(__name__)


class LayoutSectionMixin:
    """
    版式类转换功能混入类
    
    提供版式类文件的格式转换按钮和合并拆分功能。
    需要与 ConversionPanelBase 一起使用。
    """
    
    def _create_layout_buttons(self):
        """
        创建版式类格式转换按钮
        
        布局：
        - 格式转换section: PDF/OFD按钮
        - 另存为section: DOCX/DOC按钮 + TIF/JPG按钮 + DPI选项
        - 扩展section: 合并拆分选项
        """
        logger.debug("创建版式类格式按钮 - 拆分布局")
        
        # === 格式转换 section ===
        hint_label = tb.Label(
            self.conversion_container,
            text="转换为版式类型的其他格式",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置网格布局
        for col in range(2):
            self.conversion_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        self.conversion_container.grid_rowconfigure(1, weight=1)
        
        # 只显示两个按钮：转为PDF、转为OFD
        pdf_button = tb.Button(
            self.conversion_container,
            text="PDF",
            command=lambda: self._on_format_clicked("PDF"),
            bootstyle=self.button_colors['PDF'],
            **self.button_style_2col
        )
        pdf_button.grid(row=1, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        self.format_buttons['PDF'] = pdf_button
        logger.debug("  创建按钮: 转为PDF at (1, 0) - 颜色: danger")
        
        ofd_button = tb.Button(
            self.conversion_container,
            text="OFD",
            command=lambda: None,
            bootstyle='secondary',
            **self.button_style_2col,
            state='disabled'
        )
        ofd_button.grid(row=1, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        self.format_buttons['OFD'] = ofd_button
        ToolTip(ofd_button, "暂不支持")
        logger.debug("  创建按钮: 转为OFD at (1, 1) - 颜色: secondary (禁用)")
        
        # === 另存为 section ===
        self._create_layout_saveas_section()
        
        # 更新按钮状态
        self._update_button_states()
        
        # === 扩展section: 合并拆分 ===
        self._create_layout_merge_split_section()
        
        logger.info("版式类格式按钮创建完成")
    
    def _create_layout_saveas_section(self):
        """
        创建版式类另存为区域
        
        包含：
        - DOCX/DOC 文档转换按钮
        - TIF/JPG 图片渲染按钮
        - DPI 选项
        """
        # 添加说明文本
        saveas_hint_label = tb.Label(
            self.saveas_container,
            text="转换为文档或图片",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        saveas_hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置另存为容器的网格布局为2列
        for col in range(2):
            self.saveas_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 第一行按钮：另存为DOCX、另存为DOC
        formats_row1 = ['DOCX', 'DOC']
        for idx, fmt in enumerate(formats_row1):
            button = tb.Button(
                self.saveas_container,
                text=f"另存为 {fmt}",
                command=lambda f=fmt: self._on_format_clicked(f),
                bootstyle=self.button_colors[fmt],
                **self.button_style_2col
            )
            button.grid(row=1, column=idx, padx=scale(5), pady=scale(5), sticky="ew")
            self.format_buttons[fmt] = button
            # 添加tooltip
            if fmt == 'DOCX':
                ToolTip(button, "优先使用Word/LibreOffice转换，备选使用内置工具")
            elif fmt == 'DOC':
                ToolTip(button, "先转为DOCX，再通过本地Office软件转为DOC格式")
            logger.debug(f"  创建按钮: 另存为{fmt} at (1, {idx}) - 颜色: {self.button_colors[fmt]}")
        
        # 分割线
        separator = tb.Separator(self.saveas_container, bootstyle="danger")
        separator.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=scale(10))
        
        # 第二行按钮：渲染为TIF、渲染为JPG
        self.convert_to_tif_button = tb.Button(
            self.saveas_container,
            text="渲染为 TIF",
            command=lambda: self._on_convert_to_image_clicked("TIF"),
            bootstyle='warning',
            **self.button_style_2col
        )
        self.convert_to_tif_button.grid(row=3, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(self.convert_to_tif_button, "将PDF页面渲染为TIF图片，支持多页TIFF，支持选择DPI质量")
        logger.debug("  创建按钮: 渲染为TIF at (3, 0)")
        
        self.convert_to_jpg_button = tb.Button(
            self.saveas_container,
            text="渲染为 JPG",
            command=lambda: self._on_convert_to_image_clicked("JPG"),
            bootstyle='success',
            **self.button_style_2col
        )
        self.convert_to_jpg_button.grid(row=3, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(self.convert_to_jpg_button, "将PDF渲染为JPG图片（每页一个文件）")
        logger.debug("  创建按钮: 渲染为JPG at (3, 1)")
        
        # DPI选项框
        dpi_options_frame = tb.Labelframe(
            self.saveas_container,
            text="图片质量（DPI）",
            bootstyle="info"
        )
        dpi_options_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=5, pady=(5, 5))
        
        # 配置选项框架网格权重
        dpi_options_frame.grid_rowconfigure(0, weight=1)
        dpi_options_frame.grid_columnconfigure(0, weight=1)
        
        # 获取默认DPI（从配置读取）
        default_layout_dpi = 300
        if self.config_manager:
            try:
                default_layout_dpi = self.config_manager.get_layout_render_dpi()
            except Exception as e:
                logger.warning(f"读取版式DPI默认值失败: {e}")
        self.layout_image_dpi_var = tk.StringVar(value=str(default_layout_dpi))
        
        # 三个互斥单选按钮
        dpi_radio_frame = tb.Frame(dpi_options_frame, bootstyle="default")
        dpi_radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置单选按钮容器的列权重
        dpi_radio_frame.grid_columnconfigure(0, weight=1)
        dpi_radio_frame.grid_columnconfigure(1, weight=1)
        dpi_radio_frame.grid_columnconfigure(2, weight=1)
        
        self.dpi_150_radio = tb.Radiobutton(
            dpi_radio_frame, text="最小(150)", variable=self.layout_image_dpi_var,
            value='150', bootstyle="primary"
        )
        self.dpi_150_radio.grid(row=0, column=0, sticky="", padx=scale(5))
        
        self.dpi_300_radio = tb.Radiobutton(
            dpi_radio_frame, text="适中(300)", variable=self.layout_image_dpi_var,
            value='300', bootstyle="primary"
        )
        self.dpi_300_radio.grid(row=0, column=1, sticky="", padx=scale(5))
        
        self.dpi_600_radio = tb.Radiobutton(
            dpi_radio_frame, text="高清(600)", variable=self.layout_image_dpi_var,
            value='600', bootstyle="primary"
        )
        self.dpi_600_radio.grid(row=0, column=2, sticky="", padx=scale(5))
        
        logger.debug("版式类另存为section创建完成")
    
    def _create_layout_merge_split_section(self):
        """
        创建版式文件合并拆分section
        
        布局：
        - 说明文字
        - 合并按钮行：合并为PDF、合并为OFD
        - 分隔线
        - 拆分按钮行：拆分为PDF、拆分为OFD
        - 页码输入区（带placeholder）
        - 帮助文本
        - PDF文件信息显示
        - 警告标签
        """
        logger.debug("创建版式文件合并拆分section")
        
        # 显示extra_frame并设置标题
        self.extra_frame.config(text="合并拆分")
        self.extra_frame.grid()
        
        # 清空extra_container
        for widget in self.extra_container.winfo_children():
            widget.destroy()
        
        # 配置容器的grid布局为2列（与格式转换/另存为section保持一致）
        for col in range(2):
            self.extra_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 配置行权重
        for row in range(8):  # 8行：说明文字、合并按钮、分割线、拆分按钮、拆分输入、文件信息、警告
            self.extra_container.grid_rowconfigure(row, weight=0)
        
        # === Row 0: 说明文字 ===
        hint_label = tb.Label(
            self.extra_container,
            text="合并或拆分版式文件",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # === Row 1: 合并按钮（2个） ===
        # 创建"合并为PDF"按钮
        self.merge_pdfs_button = tb.Button(
            self.extra_container,
            text="合并为 PDF",
            command=self._on_merge_pdfs_clicked,
            bootstyle='info',
            **self.button_style_2col
        )
        self.merge_pdfs_button.grid(row=1, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(
            self.merge_pdfs_button,
            "将批量列表中的文件按从上到下的顺序合并为一个PDF。选中文件后可使用方向键↑↓调整合并顺序。"
        )
        
        # 创建"合并为OFD"按钮（禁用）
        merge_ofd_button = tb.Button(
            self.extra_container,
            text="合并为 OFD",
            command=lambda: None,
            bootstyle='secondary',
            **self.button_style_2col,
            state='disabled'
        )
        merge_ofd_button.grid(row=1, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(merge_ofd_button, "暂不支持")
        
        # === Row 2: 分割线 ===
        separator = tb.Separator(self.extra_container, bootstyle="warning")
        separator.grid(row=2, column=0, columnspan=2, sticky="ew", padx=scale(5), pady=scale(10))
        
        # === Row 3: 拆分按钮（2个） ===
        # 创建"拆分为PDF"按钮
        self.split_pdf_button = tb.Button(
            self.extra_container,
            text="拆分为 PDF",
            command=self._on_split_pdf_clicked,
            bootstyle='danger',
            **self.button_style_2col,
            state="disabled"
        )
        self.split_pdf_button.grid(row=3, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(
            self.split_pdf_button,
            "根据输入的页码范围将当前文件拆分为两个PDF。第1个文件包含输入的页码，第2个文件包含剩余页码（如果有）。"
        )
        
        # 创建"拆分为OFD"按钮（禁用）
        split_ofd_button = tb.Button(
            self.extra_container,
            text="拆分为 OFD",
            command=lambda: None,
            bootstyle='secondary',
            **self.button_style_2col,
            state='disabled'
        )
        split_ofd_button.grid(row=3, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(split_ofd_button, "暂不支持")
        
        # === Row 4: 拆分页码输入区 ===
        input_frame = tb.Frame(self.extra_container, bootstyle="default")
        input_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, scale(5)))
        
        # 配置输入框架
        input_frame.grid_columnconfigure(0, weight=0)  # 标签列
        input_frame.grid_columnconfigure(1, weight=1)  # 输入框列
        input_frame.grid_columnconfigure(2, weight=0)  # 信息图标列
        
        # 获取小字体
        small_font, small_size = get_small_font()
        
        # 标签
        input_label = tb.Label(
            input_frame,
            text="拆分页码范围：",
            font=(small_font, small_size),
            bootstyle="secondary",
            anchor=tk.W
        )
        input_label.grid(row=0, column=0, sticky="w", padx=(0, scale(5)))
        
        # 创建页码输入变量和输入框
        self.page_input_var = tk.StringVar()
        self.page_input_entry = tb.Entry(
            input_frame,
            textvariable=self.page_input_var,
            bootstyle="default",
            font=(small_font, small_size)
        )
        self.page_input_entry.grid(row=0, column=1, sticky="ew", padx=(0, scale(5)))
        
        # 添加信息图标
        page_range_info = create_info_icon(
            input_frame,
            "输入页码范围，例如：\n"
            "• 1-5：第1到第5页\n"
            "• 1,3,5：第1、3、5页\n"
            "• 1-3,7-9：第1-3页和第7-9页\n"
            "• 支持中文分隔符：1~3;5;7至10",
            bootstyle="info"
        )
        page_range_info.grid(row=0, column=2, sticky="w")
        
        # 设置placeholder文本（通过绑定事件模拟）
        placeholder_text = "例如：1-3,5,7-10"
        
        def on_focus_in(event):
            if self.page_input_var.get() == placeholder_text:
                self.page_input_var.set("")
                self.page_input_entry.configure(foreground="black")
        
        def on_focus_out(event):
            if not self.page_input_var.get():
                self.page_input_var.set(placeholder_text)
                self.page_input_entry.configure(foreground="gray")
        
        # 初始化placeholder
        self.page_input_var.set(placeholder_text)
        self.page_input_entry.configure(foreground="gray")
        self.page_input_entry.bind('<FocusIn>', on_focus_in)
        self.page_input_entry.bind('<FocusOut>', on_focus_out)
        
        # 绑定实时验证
        self.page_input_var.trace_add('write', lambda *args: self._on_page_input_changed())
        
        # === Row 5: 文件信息和警告 ===
        info_frame = tb.Frame(self.extra_container, bootstyle="default")
        info_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(scale(10), scale(5)))
        
        info_frame.grid_columnconfigure(0, weight=1)
        
        # PDF文件信息标题（第1行）
        self.pdf_info_title_label = tb.Label(
            info_frame,
            text="当前选中的拆分文件（共0页）：",
            font=(small_font, small_size),
            bootstyle="warning",
            anchor=tk.CENTER
        )
        self.pdf_info_title_label.grid(row=0, column=0, sticky="ew")
        
        # PDF文件名显示（第2行，支持自动换行）
        self.pdf_info_var = tk.StringVar(value="未选择")
        self.pdf_info_label = tb.Label(
            info_frame,
            textvariable=self.pdf_info_var,
            font=(small_font, small_size),
            bootstyle="primary",
            anchor=tk.CENTER,
            wraplength=250  # 设置自动换行宽度
        )
        self.pdf_info_label.grid(row=1, column=0, sticky="ew")
        
        # 页码警告标签（动态显示）
        self.page_warning_label = tb.Label(
            info_frame,
            text="",
            font=(small_font, small_size),
            bootstyle="danger",
            anchor=tk.W
        )
        self.page_warning_label.grid(row=2, column=0, sticky="w", pady=(scale(2), 0))
        
        logger.debug("版式文件合并拆分section创建完成")
    
    def _on_convert_to_image_clicked(self, image_format: str):
        """
        处理版式文件转图片按钮点击事件（TIF/JPG）
        
        参数:
            image_format: 目标图片格式 ('TIF' 或 'JPG')
        """
        if self.on_action and self.current_file_path:
            # 获取DPI值
            dpi = self.layout_image_dpi_var.get() if hasattr(self, 'layout_image_dpi_var') else '150'
            
            # 获取源文件的实际格式
            from gongwen_converter.utils.file_type_utils import detect_actual_file_format
            source_format = detect_actual_file_format(self.current_file_path)
            
            # 构建策略名称
            action_type = f"convert_layout_to_{image_format.lower()}"
            
            options = {
                'dpi': int(dpi),
                'actual_format': source_format
            }
            
            logger.info(f"版式文件转{image_format}，DPI: {dpi}, 源格式: {source_format}")
            self.on_action(action_type, self.current_file_path, options)
    
    def _on_split_pdf_clicked(self):
        """
        处理拆分PDF按钮点击事件
        
        获取页码输入，解析后调用回调函数执行拆分操作。
        """
        if self.on_action and self.current_file_path:
            # 获取并解析页码
            input_text = self.page_input_var.get() if hasattr(self, 'page_input_var') and self.page_input_var else ""
            
            # 检查是否为placeholder
            placeholder_text = "例如：1-3,5,7-10"
            if not input_text or input_text == placeholder_text:
                logger.warning("页码输入为空")
                return
            
            try:
                pages = self._parse_page_ranges(input_text)
                # 自动截断超界页码
                if self.pdf_total_pages > 0:
                    pages = [p for p in pages if p <= self.pdf_total_pages]
                
                if not pages:
                    logger.warning("没有有效的页码")
                    return
                
                options = {
                    "pages": pages,
                    "total_pages": self.pdf_total_pages
                }
                
                logger.info(f"执行拆分PDF操作，页码: {pages}")
                self.on_action("split_pdf", self.current_file_path, options)
            except ValueError as e:
                logger.error(f"页码解析失败: {e}")
    
    def _on_merge_pdfs_clicked(self):
        """
        处理合并PDF按钮点击事件
        
        调用回调函数执行合并操作。
        """
        if self.on_action and self.current_file_path:
            logger.info("执行合并PDF操作")
            self.on_action("merge_pdfs", self.current_file_path, {})
    
    def _on_page_input_changed(self, event=None):
        """
        处理页码输入变更事件
        
        实时验证输入格式，更新输入框样式和拆分按钮状态。
        """
        if not hasattr(self, 'page_input_var') or not self.page_input_var:
            return
        
        input_text = self.page_input_var.get()
        placeholder_text = "例如：1-3,5,7-10"
        
        # 空输入或placeholder：恢复默认样式
        if not input_text.strip() or input_text == placeholder_text:
            if hasattr(self, 'page_input_entry') and self.page_input_entry:
                self.page_input_entry.configure(bootstyle="default")
            if hasattr(self, 'split_pdf_button') and self.split_pdf_button:
                self.split_pdf_button.config(state="disabled")
            return
        
        # 验证输入
        is_valid = self._validate_page_input(input_text)
        
        if hasattr(self, 'page_input_entry') and self.page_input_entry:
            if is_valid:
                # 合法：绿色边框
                self.page_input_entry.configure(bootstyle="success")
            else:
                # 非法：红色边框
                self.page_input_entry.configure(bootstyle="danger")
        
        if hasattr(self, 'split_pdf_button') and self.split_pdf_button:
            # 启用/禁用拆分按钮
            self.split_pdf_button.config(state="normal" if is_valid else "disabled")
    
    def _validate_page_input(self, input_text: str) -> bool:
        """
        验证页码输入格式
        
        参数:
            input_text: 页码输入字符串
            
        返回:
            bool: 输入是否有效
        """
        # 1. 空输入检查
        if not input_text.strip():
            self._clear_page_warning()
            return False
        
        # 2. 格式检查（正则匹配：只允许数字、分隔符、范围符、空格）
        pattern = r'^[\d,，;；、\-~－至\s]+$'
        if not re.match(pattern, input_text):
            self._clear_page_warning()
            return False
        
        # 3. 解析页码范围
        try:
            pages = self._parse_page_ranges(input_text)
        except ValueError as e:
            logger.debug(f"页码解析失败: {e}")
            self._clear_page_warning()
            return False
        
        # 4. 页码有效性检查
        if not pages:
            self._clear_page_warning()
            return False
        
        # 5. 页码超界检查
        if self.pdf_total_pages > 0:
            max_page = max(pages)
            if max_page > self.pdf_total_pages:
                # 自动截断
                valid_pages = [p for p in pages if p <= self.pdf_total_pages]
                if valid_pages:
                    # 检查截断后是否涵盖全部页码
                    if set(valid_pages) == set(range(1, self.pdf_total_pages + 1)):
                        self._show_page_warning(f"❌ 不能拆分全部页码（请留下部分页码）")
                        return False
                    # 显示警告
                    self._show_page_warning(f"⚠️ 部分页码超出范围，已自动调整（共{self.pdf_total_pages}页）")
                    return True
                else:
                    self._show_page_warning(f"❌ 所有页码均超出范围（共{self.pdf_total_pages}页）")
                    return False
            else:
                # 6. 检查是否涵盖全部页码
                if set(pages) == set(range(1, self.pdf_total_pages + 1)):
                    self._show_page_warning(f"❌ 不能拆分全部页码（请留下部分页码）")
                    return False
                self._clear_page_warning()
        
        return True
    
    def _parse_page_ranges(self, input_text: str) -> List[int]:
        """
        解析页码范围字符串为页码列表
        
        支持多种分隔符：
        - 逗号：, ，
        - 分号：; ；
        - 顿号：、
        
        支持多种范围符：
        - 连字符：- ー
        - 波浪号：~
        - 中文"至"字
        
        参数:
            input_text: 页码输入字符串 (如 "1-3,5,7-10")
            
        返回:
            List[int]: 排序去重后的页码列表
            
        异常:
            ValueError: 输入格式无效时抛出
        """
        pages = set()  # 使用set自动去重
        
        # 1. 移除所有空格
        input_text = re.sub(r'\s+', '', input_text)
        
        # 2. 统一分隔符为英文逗号
        input_text = re.sub(r'[，;；、]', ',', input_text)
        
        # 3. 统一范围符号为连字符（支持波浪号、全角连字符、"至"字）
        input_text = re.sub(r'[~－至]', '-', input_text)
        
        # 4. 按逗号分割
        parts = input_text.split(',')
        
        for part in parts:
            part = part.strip()
            if not part:
                continue
            
            # 5. 处理范围 (如 "1-3")
            if '-' in part:
                range_parts = part.split('-')
                if len(range_parts) == 2:
                    try:
                        start = int(range_parts[0])
                        end = int(range_parts[1])
                        
                        # 处理倒序范围 (如 "10-5" -> "5-10")
                        if start > end:
                            start, end = end, start
                        
                        # 添加范围内所有页码
                        pages.update(range(start, end + 1))
                    except ValueError:
                        raise ValueError(f"无效的页码范围: {part}")
            else:
                # 6. 处理单个页码
                try:
                    page = int(part)
                    if page > 0:
                        pages.add(page)
                except ValueError:
                    raise ValueError(f"无效的页码: {part}")
        
        # 7. 返回排序后的列表
        return sorted(list(pages))
    
    def _show_page_warning(self, message: str):
        """
        显示页码警告信息
        
        参数:
            message: 警告消息文本
        """
        if hasattr(self, 'page_warning_label') and self.page_warning_label:
            self.page_warning_label.config(text=message)
            logger.debug(f"显示页码警告: {message}")
    
    def _clear_page_warning(self):
        """清除页码警告信息"""
        if hasattr(self, 'page_warning_label') and self.page_warning_label:
            self.page_warning_label.config(text="")
    
    def set_pdf_info(self, total_pages: int, file_name: str):
        """
        设置PDF文件信息（两行显示 - 标题+文件名）
        
        参数:
            total_pages: PDF总页数
            file_name: 文件名（显示用）
        """
        self.pdf_total_pages = total_pages
        
        # 更新标题（包含页数）
        if hasattr(self, 'pdf_info_title_label') and self.pdf_info_title_label:
            self.pdf_info_title_label.config(text=f"当前选中的拆分文件（共{total_pages}页）：")
        
        # 更新文件名
        if hasattr(self, 'pdf_info_var') and self.pdf_info_var:
            self.pdf_info_var.set(file_name)
            logger.debug(f"设置PDF信息: {total_pages}页, 文件名={file_name}")
    
    def clear_split_input(self):
        """
        清空拆分PDF的页码输入框
        
        在批量模式下切换不同PDF文件时调用，避免误用之前的页码范围
        """
        if hasattr(self, 'page_input_var') and self.page_input_var:
            # 恢复placeholder文本
            placeholder_text = "例如：1-3,5,7-10"
            self.page_input_var.set(placeholder_text)
            logger.debug("已清空拆分输入框并恢复placeholder")
        
        # 恢复placeholder样式
        if hasattr(self, 'page_input_entry') and self.page_input_entry:
            self.page_input_entry.configure(foreground="gray", bootstyle="default")
        
        # 清除警告信息
        if hasattr(self, 'page_warning_label') and self.page_warning_label:
            self.page_warning_label.config(text="")
        
        # 禁用拆分按钮
        if hasattr(self, 'split_pdf_button') and self.split_pdf_button:
            self.split_pdf_button.config(state="disabled")
            logger.debug("已禁用拆分按钮")
