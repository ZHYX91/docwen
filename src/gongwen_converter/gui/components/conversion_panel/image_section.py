"""
图片类格式转换功能模块

提供图片类文件（PNG/JPG/BMP/GIF/TIF/WebP）的格式转换按钮和合并功能。

主要功能：
- 创建图片格式转换按钮（转为其他图片格式）
- 创建压缩选项区域（最高质量/限制文件大小）
- 创建另存为PDF按钮（含尺寸选项）
- 创建合并为TIFF按钮（含透明选项）

依赖：
- ConversionPanelBase: 提供按钮样式、颜色映射等公共属性
- config_manager: 读取压缩/转换选项默认值

使用方式：
    此模块作为 Mixin 类被 ConversionPanel 继承，不应直接实例化。
"""

import logging
import tkinter as tk

import ttkbootstrap as tb

from gongwen_converter.utils.dpi_utils import scale
from gongwen_converter.utils.font_utils import get_small_font
from gongwen_converter.utils.gui_utils import ToolTip, create_info_icon

logger = logging.getLogger(__name__)


class ImageSectionMixin:
    """
    图片类转换功能混入类
    
    提供图片类文件的格式转换按钮和合并功能。
    需要与 ConversionPanelBase 一起使用。
    """
    
    def _create_image_buttons(self):
        """
        创建图片类格式转换按钮
        
        布局：
        - 格式转换section: PNG/BMP/GIF/TIF/WebP/JPG按钮（3行2列）+ 压缩选项
        - 另存为section: PDF/OFD按钮 + 尺寸选项
        - 扩展section: 合并为TIFF选项
        """
        logger.debug("创建图片类格式按钮 - 拆分布局")
        
        # === 格式转换 section ===
        hint_label = tb.Label(
            self.conversion_container,
            text="转换为图片类型的其他格式",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置网格布局
        for col in range(2):
            self.conversion_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        for row in range(1, 4):
            self.conversion_container.grid_rowconfigure(row, weight=1)
        
        # 第一行：PNG、BMP
        formats_row1 = ['PNG', 'BMP']
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
        
        # 第二行：GIF、TIF
        formats_row2 = ['GIF', 'TIF']
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
        
        # 第三行：WebP、JPG
        formats_row3 = ['WebP', 'JPG']
        for idx, fmt in enumerate(formats_row3):
            button = tb.Button(
                self.conversion_container,
                text=fmt,
                command=lambda f=fmt: self._on_format_clicked(f),
                bootstyle=self.button_colors[fmt],
                **self.button_style_2col
            )
            button.grid(row=3, column=idx, padx=scale(5), pady=scale(5), sticky="ew")
            self.format_buttons[fmt] = button
        
        # === 压缩选项 ===
        self._create_compress_options()
        
        # === 另存为 section ===
        self._create_saveas_pdf_section()
        
        # 更新按钮状态
        self._update_button_states()
        
        # === 扩展section: 合并图片 ===
        self._create_image_merge_section()
        
        logger.info("图片类格式按钮创建完成")
    
    def _create_compress_options(self):
        """创建压缩选项区域"""
        compress_options_frame = tb.Labelframe(
            self.conversion_container,
            text="压缩选项",
            bootstyle="info"
        )
        compress_options_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=5, pady=(10, 5))
        compress_options_frame.grid_rowconfigure(0, weight=1)
        compress_options_frame.grid_rowconfigure(1, weight=1)
        compress_options_frame.grid_columnconfigure(0, weight=1)
        
        # 单选按钮行
        radio_frame = tb.Frame(compress_options_frame, bootstyle="default")
        radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=(scale(10), scale(5)))
        
        # 获取默认压缩模式
        default_compress_mode = 'lossless'
        if self.config_manager:
            try:
                default_compress_mode = self.config_manager.get_image_compress_mode()
            except Exception as e:
                logger.warning(f"读取压缩模式默认值失败: {e}")
        
        self.compress_mode_var = tk.StringVar(value=default_compress_mode)
        
        # 最高质量单选按钮 + 信息图标
        lossless_frame = tb.Frame(radio_frame, bootstyle="default")
        lossless_frame.grid(row=0, column=0, sticky="w", padx=(0, scale(20)))
        
        self.compress_lossless_radio = tb.Radiobutton(
            lossless_frame,
            text="最高质量",
            variable=self.compress_mode_var,
            value='lossless',
            command=self._on_compress_mode_changed,
            bootstyle="primary"
        )
        self.compress_lossless_radio.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        quality_info = create_info_icon(
            lossless_frame,
            "无损：PNG↔PNG、PNG↔TIFF、PNG↔BMP\n"
            "有损但高质量：PNG→JPEG/WebP（quality=95）\n"
            "二次有损：JPEG→任何格式",
            bootstyle="info"
        )
        quality_info.pack(side=tk.LEFT)
        
        # 限制文件大小单选按钮
        self.compress_limit_radio = tb.Radiobutton(
            radio_frame,
            text="限制文件大小",
            variable=self.compress_mode_var,
            value='limit_size',
            command=self._on_compress_mode_changed,
            bootstyle="primary"
        )
        self.compress_limit_radio.grid(row=0, column=1, sticky="w")
        
        # 输入行
        small_font, small_size = get_small_font()
        input_frame = tb.Frame(compress_options_frame, bootstyle="default")
        input_frame.grid(row=1, column=0, sticky="ew", padx=scale(10), pady=(0, scale(10)))
        input_frame.grid_columnconfigure(1, weight=1)
        
        size_label = tb.Label(
            input_frame,
            text="文件大小上限：",
            font=(small_font, small_size),
            bootstyle="secondary"
        )
        size_label.grid(row=0, column=0, sticky="w", padx=(0, scale(5)))
        
        # 获取默认值
        default_size_limit = 200
        default_size_unit = 'KB'
        if self.config_manager:
            try:
                default_size_limit = self.config_manager.get_image_size_limit()
                default_size_unit = self.config_manager.get_image_size_unit()
            except Exception as e:
                logger.warning(f"读取大小限制默认值失败: {e}")
        
        self.size_limit_var = tk.StringVar(value=str(default_size_limit))
        self.size_limit_entry = tb.Entry(
            input_frame,
            textvariable=self.size_limit_var,
            bootstyle="default",
            font=(small_font, small_size),
            state='disabled'
        )
        self.size_limit_entry.grid(row=0, column=1, sticky="ew", padx=(0, scale(5)))
        
        self.size_unit_var = tk.StringVar(value=default_size_unit)
        self.size_unit_combo = tb.Combobox(
            input_frame,
            textvariable=self.size_unit_var,
            values=['KB', 'MB'],
            bootstyle="default",
            font=(small_font, small_size),
            state='disabled',
            width=6
        )
        self.size_unit_combo.grid(row=0, column=2, sticky="ew")
        
        # 绑定输入验证
        self.size_limit_var.trace_add('write', lambda *args: self._on_size_input_changed())
        
        # 警告标签
        warning_frame = tb.Frame(compress_options_frame, bootstyle="default")
        warning_frame.grid(row=2, column=0, sticky="ew", padx=scale(10), pady=(0, scale(5)))
        warning_frame.grid_columnconfigure(0, weight=1)
        
        self.size_warning_label = tb.Label(
            warning_frame,
            text="",
            font=(small_font, small_size),
            bootstyle="danger",
            anchor=tk.W
        )
        self.size_warning_label.grid(row=0, column=0, sticky="w")
    
    def _create_saveas_pdf_section(self):
        """创建另存为PDF区域"""
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
        self.convert_image_to_pdf_button = tb.Button(
            self.saveas_container,
            text="另存为 PDF",
            command=self._on_convert_to_pdf_clicked,
            bootstyle=self.button_colors['PDF'],
            **self.button_style_2col
        )
        self.convert_image_to_pdf_button.grid(row=1, column=0, padx=scale(5), pady=(scale(5), scale(10)), sticky="ew")
        ToolTip(self.convert_image_to_pdf_button, "将图片嵌入PDF文件，支持选择尺寸选项")
        
        # OFD按钮（禁用）
        ofd_button = tb.Button(
            self.saveas_container,
            text="另存为 OFD",
            command=lambda: None,
            bootstyle='secondary',
            **self.button_style_2col,
            state='disabled'
        )
        ofd_button.grid(row=1, column=1, padx=scale(5), pady=(scale(5), scale(10)), sticky="ew")
        self.format_buttons['OFD'] = ofd_button
        ToolTip(ofd_button, "暂不支持")
        
        # PDF尺寸选项
        size_options_frame = tb.Labelframe(
            self.saveas_container,
            text="尺寸选项",
            bootstyle="info"
        )
        size_options_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        size_options_frame.grid_rowconfigure(0, weight=1)
        size_options_frame.grid_columnconfigure(0, weight=1)
        
        # 获取默认PDF质量
        default_pdf_quality = 'original'
        if self.config_manager:
            try:
                default_pdf_quality = self.config_manager.get_image_pdf_quality()
            except Exception as e:
                logger.warning(f"读取PDF质量默认值失败: {e}")
        
        self.pdf_quality_var = tk.StringVar(value=default_pdf_quality)
        
        quality_radio_frame = tb.Frame(size_options_frame, bootstyle="default")
        quality_radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        for col in range(3):
            quality_radio_frame.grid_columnconfigure(col, weight=1)
        
        self.quality_original_radio = tb.Radiobutton(
            quality_radio_frame, text="原图嵌入", variable=self.pdf_quality_var,
            value='original', bootstyle="primary", command=self._on_pdf_quality_changed
        )
        self.quality_original_radio.grid(row=0, column=0, sticky="", padx=scale(5))
        
        self.quality_a4_radio = tb.Radiobutton(
            quality_radio_frame, text="适合A4", variable=self.pdf_quality_var,
            value='a4', bootstyle="primary", command=self._on_pdf_quality_changed
        )
        self.quality_a4_radio.grid(row=0, column=1, sticky="", padx=scale(5))
        
        self.quality_a3_radio = tb.Radiobutton(
            quality_radio_frame, text="适合A3", variable=self.pdf_quality_var,
            value='a3', bootstyle="primary", command=self._on_pdf_quality_changed
        )
        self.quality_a3_radio.grid(row=0, column=2, sticky="", padx=scale(5))
    
    def _create_image_merge_section(self):
        """创建图片合并section"""
        logger.debug("创建图片合并section")
        
        self.extra_frame.config(text="合并图片")
        self.extra_frame.grid()
        
        for widget in self.extra_container.winfo_children():
            widget.destroy()
        
        self.extra_container.grid_rowconfigure(0, weight=0)
        self.extra_container.grid_rowconfigure(1, weight=0)
        self.extra_container.grid_rowconfigure(2, weight=0)
        self.extra_container.grid_columnconfigure(0, weight=1)
        
        hint_label = tb.Label(
            self.extra_container,
            text="将图片合并为多页TIFF文件",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, sticky="ew", padx=5, pady=(0, 5))
        
        self.merge_tiff_button = tb.Button(
            self.extra_container,
            text="📋 合并为 TIF",
            command=self._on_merge_tiff_clicked,
            bootstyle='info',
            **self.button_style_1col
        )
        self.merge_tiff_button.grid(row=1, column=0, pady=scale(5))
        
        # 转换选项边框
        tiff_options_frame = tb.Labelframe(
            self.extra_container,
            text="转换选项",
            bootstyle="info"
        )
        tiff_options_frame.grid(row=2, column=0, sticky="ew", padx=scale(5), pady=scale(5))
        tiff_options_frame.grid_rowconfigure(0, weight=1)
        tiff_options_frame.grid_columnconfigure(0, weight=1)
        
        # 获取默认TIFF模式
        default_tiff_mode = "smart"
        if self.config_manager:
            try:
                default_tiff_mode = self.config_manager.get_image_tiff_mode()
            except Exception as e:
                logger.warning(f"读取TIFF模式默认值失败: {e}")
        
        self.tiff_mode_var = tk.StringVar(value=default_tiff_mode)
        
        tiff_radio_frame = tb.Frame(tiff_options_frame, bootstyle="default")
        tiff_radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 保留透明选项
        option1_frame = tb.Frame(tiff_radio_frame, bootstyle="default")
        option1_frame.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(20)))
        
        self.tiff_mode_smart_radio = tb.Radiobutton(
            option1_frame,
            text="保留透明",
            variable=self.tiff_mode_var,
            value="smart",
            bootstyle="primary"
        )
        self.tiff_mode_smart_radio.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        tiff_smart_info = create_info_icon(
            option1_frame,
            "如果源图片有透明背景将被保留，否则转为不透明。文件体积稍大（约增加25%）",
            bootstyle="info"
        )
        tiff_smart_info.pack(side=tk.LEFT)
        
        # 不保留透明选项
        option2_frame = tb.Frame(tiff_radio_frame, bootstyle="default")
        option2_frame.grid(row=0, column=1, sticky="w", padx=(0, scale(10)))
        
        self.tiff_mode_rgb_radio = tb.Radiobutton(
            option2_frame,
            text="不保留透明",
            variable=self.tiff_mode_var,
            value="RGB",
            bootstyle="primary"
        )
        self.tiff_mode_rgb_radio.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        tiff_rgb_info = create_info_icon(
            option2_frame,
            "所有图片统一转为RGB模式，透明背景变为白色。文件较小，适合打印和归档",
            bootstyle="info"
        )
        tiff_rgb_info.pack(side=tk.LEFT)
        
        logger.debug("图片合并section创建完成")
    
    def _on_compress_mode_changed(self):
        """处理压缩模式切换事件"""
        if not hasattr(self, 'compress_mode_var') or not self.compress_mode_var:
            return
        
        mode = self.compress_mode_var.get()
        is_limit = (mode == 'limit_size')
        
        logger.debug(f"压缩模式变更: {mode}")
        
        if hasattr(self, 'size_limit_entry') and hasattr(self, 'size_unit_combo'):
            state = 'normal' if is_limit else 'disabled'
            self.size_limit_entry.config(state=state)
            self.size_unit_combo.config(state=state)
            
            if is_limit:
                if not self.size_limit_var.get():
                    self.size_limit_var.set("200")
                self.size_limit_entry.focus()
                self._on_size_input_changed()
            else:
                self._clear_size_warning()
        
        self._update_button_states()
    
    def _on_size_input_changed(self):
        """处理文件大小输入变更事件"""
        if not hasattr(self, 'size_limit_var') or not self.size_limit_var:
            return
        
        input_text = self.size_limit_var.get()
        
        if not input_text.strip():
            self._show_size_warning("⚠️ 请输入文件大小上限")
            if hasattr(self, 'size_limit_entry'):
                self.size_limit_entry.configure(bootstyle="danger")
            self._disable_format_buttons_for_compress()
            return
        
        unit = self.size_unit_var.get() if hasattr(self, 'size_unit_var') and self.size_unit_var else 'KB'
        is_valid = self._validate_size_input(input_text, unit)
        
        if is_valid:
            if hasattr(self, 'size_limit_entry'):
                self.size_limit_entry.configure(bootstyle="success")
            self._clear_size_warning()
            self._update_button_states()
        else:
            if hasattr(self, 'size_limit_entry'):
                self.size_limit_entry.configure(bootstyle="danger")
            try:
                int(input_text)
                if unit == 'KB':
                    self._show_size_warning("❌ KB范围：1-10240")
                else:
                    self._show_size_warning("❌ MB范围：1-100")
            except ValueError:
                self._show_size_warning("❌ 请输入有效的数字")
            self._disable_format_buttons_for_compress()
    
    def _show_size_warning(self, message: str):
        """显示文件大小警告信息"""
        if hasattr(self, 'size_warning_label') and self.size_warning_label:
            self.size_warning_label.config(text=message)
    
    def _clear_size_warning(self):
        """清除文件大小警告信息"""
        if hasattr(self, 'size_warning_label') and self.size_warning_label:
            self.size_warning_label.config(text="")
    
    def _disable_format_buttons_for_compress(self):
        """在压缩模式下输入无效时禁用所有格式按钮"""
        for fmt, button in self.format_buttons.items():
            if fmt.upper() == 'OFD':
                continue
            button.configure(state='disabled', bootstyle='secondary')
    
    def _on_convert_to_pdf_clicked(self):
        """处理图片转PDF按钮点击事件"""
        if self.on_action and self.current_file_path:
            quality_mode = self.pdf_quality_var.get() if hasattr(self, 'pdf_quality_var') else ''
            if not quality_mode:
                logger.warning("未选择PDF尺寸选项")
                return
            
            options = {'quality_mode': quality_mode}
            logger.info(f"图片转PDF，质量模式: {quality_mode}")
            self.on_action("convert_image_to_pdf", self.current_file_path, options)
    
    def _on_pdf_quality_changed(self):
        """处理PDF尺寸选项变更事件"""
        if hasattr(self, 'pdf_quality_var') and hasattr(self, 'convert_image_to_pdf_button'):
            quality = self.pdf_quality_var.get()
            should_enable = bool(quality)
            self.convert_image_to_pdf_button.config(state="normal" if should_enable else "disabled")
    
    def _on_merge_tiff_clicked(self):
        """处理合并为TIFF按钮点击事件"""
        if self.on_action and self.current_file_path:
            mode = self.tiff_mode_var.get() if hasattr(self, 'tiff_mode_var') else "smart"
            options = {"mode": mode}
            logger.info(f"执行合并为TIFF操作，模式: {mode}")
            self.on_action("merge_images_to_tiff", self.current_file_path, options)
