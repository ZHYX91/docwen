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

import ttkbootstrap as tb

from docwen.gui.core.mixins_protocols import ConversionPanelHost
from docwen.i18n import t
from docwen.utils.dpi_utils import scale
from docwen.utils.font_utils import get_small_font
from docwen.utils.gui_utils import ToolTip, bind_label_wraplength_to_container, create_info_icon

logger = logging.getLogger(__name__)


class LayoutSectionMixin:
    """
    版式类转换功能混入类

    提供版式类文件的格式转换按钮和合并拆分功能。
    需要与 ConversionPanelBase 一起使用。
    """

    def _create_layout_buttons(self: ConversionPanelHost):
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
            text=t("conversion_panel.layout.convert_to_other_formats"),
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=scale(5), pady=(0, scale(5)))

        # 配置网格布局
        for col in range(2):
            self.conversion_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        self.conversion_container.grid_rowconfigure(1, weight=1)

        # 只显示两个按钮：转为PDF、转为OFD
        pdf_button = tb.Button(
            self.conversion_container,
            text="PDF",
            command=lambda: self.on_format_clicked("PDF"),
            bootstyle=self.button_colors["PDF"],
            **self.button_style_2col,
        )
        pdf_button.grid(row=1, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        self.format_buttons["PDF"] = pdf_button
        logger.debug("  创建按钮: 转为PDF at (1, 0) - 颜色: danger")

        ofd_button = tb.Button(
            self.conversion_container,
            text="OFD",
            command=lambda: None,
            bootstyle="secondary",
            **self.button_style_2col,
            state="disabled",
        )
        ofd_button.grid(row=1, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        self.format_buttons["OFD"] = ofd_button
        ToolTip(ofd_button, t("conversion_panel.layout.save_as_ofd_tooltip"))
        logger.debug("  创建按钮: 转为OFD at (1, 1) - 颜色: secondary (禁用)")

        # === 另存为 section ===
        self._create_layout_saveas_section()

        # 更新按钮状态
        self.update_button_states()

        # === 扩展section: 合并拆分 ===
        self._create_layout_merge_split_section()

        logger.info("版式类格式按钮创建完成")

    def _create_layout_saveas_section(self: ConversionPanelHost):
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
            text=t("conversion_panel.layout.convert_to_document_or_image"),
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
        )
        saveas_hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=scale(5), pady=(0, scale(5)))

        # 配置另存为容器的网格布局为2列
        for col in range(2):
            self.saveas_container.grid_columnconfigure(col, weight=1, uniform="format_col")

        # 第一行按钮：另存为DOCX、另存为DOC
        formats_row1 = ["DOCX", "DOC"]
        for idx, fmt in enumerate(formats_row1):
            button = tb.Button(
                self.saveas_container,
                text=t("conversion_panel.layout.save_as_format", format=fmt),
                command=lambda f=fmt: self.on_format_clicked(f),
                bootstyle=self.button_colors[fmt],
                **self.button_style_2col,
            )
            button.grid(row=1, column=idx, padx=scale(5), pady=scale(5), sticky="ew")
            self.format_buttons[fmt] = button
            # 添加tooltip
            if fmt == "DOCX":
                ToolTip(button, t("conversion_panel.layout.save_as_docx_tooltip"))
            elif fmt == "DOC":
                ToolTip(button, t("conversion_panel.layout.save_as_doc_tooltip"))
            logger.debug(f"  创建按钮: 另存为{fmt} at (1, {idx}) - 颜色: {self.button_colors[fmt]}")

        # 分割线
        separator = tb.Separator(self.saveas_container, bootstyle="danger")
        separator.grid(row=2, column=0, columnspan=2, sticky="ew", padx=scale(5), pady=scale(10))

        # 第二行按钮：渲染为TIF、渲染为JPG
        self.convert_to_tif_button = tb.Button(
            self.saveas_container,
            text=t("conversion_panel.layout.render_as_tif"),
            command=lambda: self._on_convert_to_image_clicked("TIF"),
            bootstyle="warning",
            **self.button_style_2col,
        )
        self.convert_to_tif_button.grid(row=3, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(self.convert_to_tif_button, t("conversion_panel.layout.render_as_tif_tooltip"))
        logger.debug("  创建按钮: 渲染为TIF at (3, 0)")

        self.convert_to_jpg_button = tb.Button(
            self.saveas_container,
            text=t("conversion_panel.layout.render_as_jpg"),
            command=lambda: self._on_convert_to_image_clicked("JPG"),
            bootstyle="success",
            **self.button_style_2col,
        )
        self.convert_to_jpg_button.grid(row=3, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(self.convert_to_jpg_button, t("conversion_panel.layout.render_as_jpg_tooltip"))
        logger.debug("  创建按钮: 渲染为JPG at (3, 1)")

        # DPI选项框
        dpi_options_frame = tb.Labelframe(
            self.saveas_container, text=t("conversion_panel.layout.image_quality_dpi"), bootstyle="info"
        )
        dpi_options_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=scale(5), pady=(scale(5), scale(5)))

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

        # 单列布局
        dpi_radio_frame = tb.Frame(dpi_options_frame, bootstyle="default")
        dpi_radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))

        dpi_radio_frame.grid_columnconfigure(0, weight=1)

        self.dpi_150_radio = tb.Radiobutton(
            dpi_radio_frame,
            text=t("conversion_panel.layout.minimum_150"),
            variable=self.layout_image_dpi_var,
            value="150",
            bootstyle="primary",
        )
        self.dpi_150_radio.grid(row=0, column=0, sticky="w", padx=scale(10), pady=scale(3))

        self.dpi_300_radio = tb.Radiobutton(
            dpi_radio_frame,
            text=t("conversion_panel.layout.moderate_300"),
            variable=self.layout_image_dpi_var,
            value="300",
            bootstyle="primary",
        )
        self.dpi_300_radio.grid(row=1, column=0, sticky="w", padx=scale(10), pady=scale(3))

        self.dpi_600_radio = tb.Radiobutton(
            dpi_radio_frame,
            text=t("conversion_panel.layout.high_definition_600"),
            variable=self.layout_image_dpi_var,
            value="600",
            bootstyle="primary",
        )
        self.dpi_600_radio.grid(row=2, column=0, sticky="w", padx=scale(10), pady=scale(3))

        logger.debug("版式类另存为section创建完成")

    def _create_layout_merge_split_section(self: ConversionPanelHost):
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
        self.extra_frame.config(text=t("conversion_panel.layout.merge_split"))
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
            text=t("conversion_panel.layout.merge_split_hint"),
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=scale(5), pady=(0, scale(5)))

        # === Row 1: 合并按钮（2个） ===
        # 创建"合并为PDF"按钮
        self.merge_pdfs_button = tb.Button(
            self.extra_container,
            text=t("conversion_panel.layout.merge_to_pdf"),
            command=self._on_merge_pdfs_clicked,
            bootstyle="info",
            **self.button_style_2col,
        )
        self.merge_pdfs_button.grid(row=1, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(self.merge_pdfs_button, t("conversion_panel.layout.merge_to_pdf_tooltip"))

        # 创建"合并为OFD"按钮（禁用）
        merge_ofd_button = tb.Button(
            self.extra_container,
            text=t("conversion_panel.layout.merge_to_ofd"),
            command=lambda: None,
            bootstyle="secondary",
            **self.button_style_2col,
            state="disabled",
        )
        merge_ofd_button.grid(row=1, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(merge_ofd_button, t("conversion_panel.layout.merge_to_ofd_tooltip"))

        # === Row 2: 分割线 ===
        separator = tb.Separator(self.extra_container, bootstyle="warning")
        separator.grid(row=2, column=0, columnspan=2, sticky="ew", padx=scale(5), pady=scale(10))

        # === Row 3: 拆分按钮（2个） ===
        # 创建"拆分为PDF"按钮
        self.split_pdf_button = tb.Button(
            self.extra_container,
            text=t("conversion_panel.layout.split_to_pdf"),
            command=self._on_split_pdf_clicked,
            bootstyle="danger",
            **self.button_style_2col,
            state="disabled",
        )
        self.split_pdf_button.grid(row=3, column=0, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(self.split_pdf_button, t("conversion_panel.layout.split_to_pdf_tooltip"))

        # 创建"拆分为OFD"按钮（禁用）
        split_ofd_button = tb.Button(
            self.extra_container,
            text=t("conversion_panel.layout.split_to_ofd"),
            command=lambda: None,
            bootstyle="secondary",
            **self.button_style_2col,
            state="disabled",
        )
        split_ofd_button.grid(row=3, column=1, padx=scale(5), pady=scale(5), sticky="ew")
        ToolTip(split_ofd_button, t("conversion_panel.layout.split_to_ofd_tooltip"))

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
            text=t("conversion_panel.layout.split_page_range"),
            font=(small_font, small_size),
            bootstyle="secondary",
            anchor=tk.W,
        )
        input_label.grid(row=0, column=0, sticky="w", padx=(0, scale(5)))

        # 创建页码输入变量和输入框
        self.page_input_var = tk.StringVar()
        self.page_input_entry = tb.Entry(
            input_frame, textvariable=self.page_input_var, bootstyle="default", font=(small_font, small_size)
        )
        self.page_input_entry.grid(row=0, column=1, sticky="ew", padx=(0, scale(5)))

        # 添加信息图标
        page_range_info = create_info_icon(input_frame, t("conversion_panel.layout.page_range_info"), bootstyle="info")
        page_range_info.grid(row=0, column=2, sticky="w")

        # 设置placeholder文本（通过绑定事件模拟）
        placeholder_text = t("conversion_panel.layout.page_range_placeholder")

        def on_focus_in(event):
            if self.page_input_var.get() == placeholder_text:
                self.page_input_var.set("")
                self.page_input_entry.configure(bootstyle="default")

        def on_focus_out(event):
            if not self.page_input_var.get():
                self.page_input_var.set(placeholder_text)
                self.page_input_entry.configure(bootstyle="secondary")

        # 初始化placeholder
        self.page_input_var.set(placeholder_text)
        self.page_input_entry.configure(bootstyle="secondary")
        self.page_input_entry.bind("<FocusIn>", on_focus_in)
        self.page_input_entry.bind("<FocusOut>", on_focus_out)

        # 绑定实时验证
        self.page_input_var.trace_add("write", lambda *args: self._on_page_input_changed())

        # === Row 5: 文件信息和警告 ===
        info_frame = tb.Frame(self.extra_container, bootstyle="default")
        info_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(scale(10), scale(5)))

        info_frame.grid_columnconfigure(0, weight=1)

        # PDF文件信息标题（第1行）
        self.pdf_info_title_label = tb.Label(
            info_frame,
            text=t("conversion_panel.layout.selected_split_file", pages=0),
            font=(small_font, small_size),
            bootstyle="warning",
            anchor=tk.CENTER,
        )
        self.pdf_info_title_label.grid(row=0, column=0, sticky="ew")

        # PDF文件名显示（第2行，支持自动换行）
        self.pdf_info_var = tk.StringVar(value=t("conversion_panel.layout.no_file_selected"))
        self.pdf_info_label = tb.Label(
            info_frame,
            textvariable=self.pdf_info_var,
            font=(small_font, small_size),
            bootstyle="primary",
            anchor=tk.CENTER,
        )
        self.pdf_info_label.grid(row=1, column=0, sticky="ew")
        bind_label_wraplength_to_container(
            self.pdf_info_label, info_frame, min_wraplength=scale(200), padding=scale(20)
        )

        # 页码警告标签（动态显示）
        self.page_warning_label = tb.Label(
            info_frame, text="", font=(small_font, small_size), bootstyle="danger", anchor=tk.W
        )
        self.page_warning_label.grid(row=2, column=0, sticky="w", pady=(scale(2), 0))

        logger.debug("版式文件合并拆分section创建完成")

    def _on_convert_to_image_clicked(self: ConversionPanelHost, image_format: str):
        """
        处理版式文件转图片按钮点击事件（TIF/JPG）

        参数:
            image_format: 目标图片格式 ('TIF' 或 'JPG')
        """
        if self.on_action and self.current_file_path:
            # 获取DPI值
            dpi = self.layout_image_dpi_var.get() if hasattr(self, "layout_image_dpi_var") else "150"

            # 获取源文件的实际格式
            from docwen.utils.file_type_utils import detect_actual_file_format

            source_format = detect_actual_file_format(self.current_file_path)

            # 构建策略名称
            action_type = f"convert_layout_to_{image_format.lower()}"

            options = {"dpi": int(dpi), "actual_format": source_format}

            logger.info(f"版式文件转{image_format}，DPI: {dpi}, 源格式: {source_format}")
            self.on_action(action_type, self.current_file_path, options)

    def _on_split_pdf_clicked(self: ConversionPanelHost):
        """
        处理拆分PDF按钮点击事件

        获取页码输入，解析后调用回调函数执行拆分操作。
        """
        if self.on_action and self.current_file_path:
            # 获取并解析页码
            input_text = self.page_input_var.get() if hasattr(self, "page_input_var") and self.page_input_var else ""

            # 检查是否为placeholder
            placeholder_text = t("conversion_panel.layout.page_range_placeholder")
            if not input_text or input_text == placeholder_text:
                logger.warning("页码输入为空")
                return

            try:
                split_mode, pages = self._parse_split_input(input_text)

                if split_mode != "custom" and self.pdf_total_pages <= 1:
                    self._show_page_warning(t("conversion_panel.layout.split_mode_single_page_warning"))
                    return

                options: dict[str, object] = {"split_mode": split_mode, "total_pages": self.pdf_total_pages}
                if split_mode == "custom":
                    parsed_pages = pages or []
                    if self.pdf_total_pages > 0:
                        parsed_pages = [p for p in parsed_pages if p <= self.pdf_total_pages]
                    if not parsed_pages:
                        logger.warning("没有有效的页码")
                        return
                    options["pages"] = parsed_pages

                logger.info(f"执行拆分PDF操作，模式: {split_mode}, 页码: {pages}")
                self.on_action("split_pdf", self.current_file_path, options)
            except ValueError as e:
                logger.error(f"页码解析失败: {e}")

    def _on_merge_pdfs_clicked(self: ConversionPanelHost):
        """
        处理合并PDF按钮点击事件

        调用回调函数执行合并操作。
        """
        if self.on_action and self.current_file_path:
            logger.info("执行合并PDF操作")
            self.on_action("merge_pdfs", self.current_file_path, {})

    def _on_page_input_changed(self: ConversionPanelHost, event=None):
        """
        处理页码输入变更事件

        实时验证输入格式，更新输入框样式和拆分按钮状态。
        """
        if not hasattr(self, "page_input_var") or not self.page_input_var:
            return

        input_text = self.page_input_var.get()
        placeholder_text = t("conversion_panel.layout.page_range_placeholder")

        # 空输入或placeholder：恢复默认样式
        if not input_text.strip() or input_text == placeholder_text:
            if hasattr(self, "page_input_entry") and self.page_input_entry:
                self.page_input_entry.configure(bootstyle="secondary" if input_text == placeholder_text else "default")
            if hasattr(self, "split_pdf_button") and self.split_pdf_button:
                self.split_pdf_button.config(state="disabled")
            return

        # 验证输入
        is_valid = self._validate_page_input(input_text)

        if hasattr(self, "page_input_entry") and self.page_input_entry:
            if is_valid:
                # 合法：绿色边框
                self.page_input_entry.configure(bootstyle="success")
            else:
                # 非法：红色边框
                self.page_input_entry.configure(bootstyle="danger")

        if hasattr(self, "split_pdf_button") and self.split_pdf_button:
            # 启用/禁用拆分按钮
            self.split_pdf_button.config(state="normal" if is_valid else "disabled")

    def _validate_page_input(self: ConversionPanelHost, input_text: str) -> bool:
        """
        验证页码输入格式

        参数:
            input_text: 页码输入字符串

        返回:
            bool: 输入是否有效
        """
        # 1. 空输入检查
        text = input_text.strip()
        if not text:
            self._clear_page_warning()
            return False

        # 2. 特殊模式输入
        if text in ("*", "#"):
            if self.pdf_total_pages <= 1 and self.pdf_total_pages > 0:
                self._show_page_warning(t("conversion_panel.layout.split_mode_single_page_warning"))
            else:
                self._clear_page_warning()
            return True

        # 2. 格式检查（正则匹配：只允许数字、分隔符、范围符、空格）
        pattern = r"^[\d,，;；、\-~－至\s]+$"
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
                        self._show_page_warning(t("conversion_panel.layout.cannot_split_all_pages"))
                        return False
                    # 显示警告
                    self._show_page_warning(
                        t("conversion_panel.layout.pages_out_of_range_adjusted", pages=self.pdf_total_pages)
                    )
                    return True
                else:
                    self._show_page_warning(
                        t("conversion_panel.layout.all_pages_out_of_range", pages=self.pdf_total_pages)
                    )
                    return False
            else:
                # 6. 检查是否涵盖全部页码
                if set(pages) == set(range(1, self.pdf_total_pages + 1)):
                    self._show_page_warning(t("conversion_panel.layout.cannot_split_all_pages"))
                    return False
                self._clear_page_warning()

        return True

    def _parse_split_input(self: ConversionPanelHost, input_text: str) -> tuple[str, list[int] | None]:
        text = input_text.strip()
        if text == "*":
            return ("every_page", None)
        if text == "#":
            return ("odd_even", None)
        pages = self._parse_page_ranges(input_text)
        return ("custom", pages)

    def _parse_page_ranges(self: ConversionPanelHost, input_text: str) -> list[int]:
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
        input_text = re.sub(r"\s+", "", input_text)

        # 2. 统一分隔符为英文逗号
        input_text = re.sub(r"[，;；、]", ",", input_text)

        # 3. 统一范围符号为连字符（支持波浪号、全角连字符、"至"字）
        input_text = re.sub(r"[~－至]", "-", input_text)

        # 4. 按逗号分割
        parts = input_text.split(",")

        for part in parts:
            part = part.strip()
            if not part:
                continue

            # 5. 处理范围 (如 "1-3")
            if "-" in part:
                range_parts = part.split("-")
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
                        raise ValueError(f"无效的页码范围: {part}") from None
            else:
                # 6. 处理单个页码
                try:
                    page = int(part)
                    if page > 0:
                        pages.add(page)
                except ValueError:
                    raise ValueError(f"无效的页码: {part}") from None

        # 7. 返回排序后的列表
        return sorted(pages)

    def _show_page_warning(self: ConversionPanelHost, message: str):
        """
        显示页码警告信息

        参数:
            message: 警告消息文本
        """
        if hasattr(self, "page_warning_label") and self.page_warning_label:
            self.page_warning_label.config(text=message)
            logger.debug(f"显示页码警告: {message}")

    def _clear_page_warning(self: ConversionPanelHost):
        """清除页码警告信息"""
        if hasattr(self, "page_warning_label") and self.page_warning_label:
            self.page_warning_label.config(text="")

    def set_pdf_info(self: ConversionPanelHost, total_pages: int, file_name: str):
        """
        设置PDF文件信息（两行显示 - 标题+文件名）

        参数:
            total_pages: PDF总页数
            file_name: 文件名（显示用）
        """
        self.pdf_total_pages = total_pages

        # 更新标题（包含页数）
        if hasattr(self, "pdf_info_title_label") and self.pdf_info_title_label:
            self.pdf_info_title_label.config(text=t("conversion_panel.layout.selected_split_file", pages=total_pages))

        # 更新文件名
        if hasattr(self, "pdf_info_var") and self.pdf_info_var:
            self.pdf_info_var.set(file_name)
            logger.debug(f"设置PDF信息: {total_pages}页, 文件名={file_name}")

    def clear_split_input(self: ConversionPanelHost):
        """
        清空拆分PDF的页码输入框

        在批量模式下切换不同PDF文件时调用，避免误用之前的页码范围
        """
        if hasattr(self, "page_input_var") and self.page_input_var:
            # 恢复placeholder文本
            placeholder_text = t("conversion_panel.layout.page_range_placeholder")
            self.page_input_var.set(placeholder_text)
            logger.debug("已清空拆分输入框并恢复placeholder")

        # 恢复placeholder样式
        if hasattr(self, "page_input_entry") and self.page_input_entry:
            self.page_input_entry.configure(bootstyle="secondary")

        # 清除警告信息
        if hasattr(self, "page_warning_label") and self.page_warning_label:
            self.page_warning_label.config(text="")

        # 禁用拆分按钮
        if hasattr(self, "split_pdf_button") and self.split_pdf_button:
            self.split_pdf_button.config(state="disabled")
            logger.debug("已禁用拆分按钮")
