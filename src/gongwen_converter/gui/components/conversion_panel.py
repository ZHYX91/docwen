"""
格式转换面板组件
用于非文本类文件的同类格式转换
支持文档、表格、图片、版式类文件的格式转换
"""

import logging
import re
import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable, Dict, List
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from gongwen_converter.utils.dpi_utils import scale
from gongwen_converter.utils.font_utils import get_default_font
from gongwen_converter.utils.gui_utils import ToolTip

logger = logging.getLogger(__name__)


class ConversionPanel(tb.Frame):
    """
    格式转换面板组件
    根据文件类别显示对应的转换格式按钮
    """
    
    # 各类别支持的格式定义
    FORMATS_BY_CATEGORY = {
        'document': ['DOCX', 'DOC', 'RTF', 'ODT', 'PDF'],
        'spreadsheet': ['XLSX', 'XLS', 'ET', 'CSV', 'ODS'],
        'image': ['PNG', 'JPG', 'BMP', 'GIF', 'TIF', 'WebP'],
        'layout': ['PDF', 'DOCX']  # PDF可以提取为DOCX
    }
    
    # 支持有损压缩的图片格式（可通过quality参数有效减小文件）
    COMPRESSIBLE_FORMATS = ['JPG', 'JPEG', 'WEBP']
    
    def __init__(self, master, config_manager: any = None, on_format_selected: Optional[Callable] = None, on_action: Optional[Callable] = None, height: Optional[int] = None, **kwargs):
        """
        初始化格式转换面板
        
        参数:
            master: 父组件
            config_manager: 配置管理器实例，用于读取校对选项默认值
            on_format_selected: 格式选择回调函数 fn(target_format: str) (保留用于兼容性，但优先使用on_action)
            on_action: 操作回调函数 fn(action_type: str, file_path: str, options: dict)
            height: 格式转换section的目标高度（用于匹配文件输入区域高度）
        """
        logger.debug("初始化格式转换面板组件")
        super().__init__(master, **kwargs)
        
        self.config_manager = config_manager
        self.on_format_selected = on_format_selected
        self.on_action = on_action
        self.current_category = None
        self.current_format = None
        self.current_file_path = None
        self.format_buttons: Dict[str, tb.Button] = {}
        self.conversion_section_height = height  # 保存格式转换section的目标高度
        
        # 存储校对选项和汇总选项的变量
        self.checkbox_vars: Dict[str, tk.BooleanVar] = {}
        self.merge_mode_var: Optional[tk.IntVar] = None
        self.reference_table_var: Optional[tk.StringVar] = None
        
        # 合并拆分相关状态
        self.pdf_total_pages = 0  # PDF总页数
        self.page_input_var: Optional[tk.StringVar] = None  # 页码输入变量
        self.page_input_entry: Optional[tb.Entry] = None  # 页码输入框引用
        self.split_pdf_button: Optional[tb.Button] = None  # 拆分按钮引用
        self.merge_pdfs_button: Optional[tb.Button] = None  # 合并按钮引用
        self.pdf_info_var: Optional[tk.StringVar] = None  # 文件信息变量
        self.pdf_info_label: Optional[tb.Label] = None  # 文件信息标签引用
        self.page_warning_label: Optional[tb.Label] = None  # 页码警告标签引用
        
        # 获取字体配置
        self.default_font, self.default_size = get_default_font()
        
        # 按钮颜色映射（与操作面板保持一致）
        self.button_colors = {
            # 文档类
            'DOCX': 'primary',    # 主题色：主要格式
            'DOC': 'info',        # 蓝色：旧版格式
            'ODT': 'success',     # 绿色：开放格式
            'RTF': 'warning',     # 橙色：富文本格式
            # 表格类
            'XLSX': 'primary',    # 主题色：主要格式
            'XLS': 'info',        # 蓝色：旧版格式
            'ET': 'info',         # 蓝色：WPS格式
            'ODS': 'success',     # 绿色：开放格式
            'CSV': 'warning',     # 橙色：通用格式
            # 图片类
            'PNG': 'primary',     # 主题色：常用格式
            'JPG': 'primary',     # 主题色：常用格式
            'BMP': 'info',        # 蓝色：基础格式
            'GIF': 'success',     # 绿色：动画格式
            'TIF': 'warning',     # 橙色：专业格式
            'WebP': 'danger',     # 红色：现代格式
            # 版式类
            'PDF': 'danger',      # 红色：PDF格式
            'XPS': 'info',        # 蓝色：版式格式
            'OFD': 'success',     # 绿色：国产格式
            'CEB': 'warning'      # 橙色：电子书格式
        }
        
        # 按钮样式定义（略小于操作面板，支持DPI缩放）
        self.button_style_3col = {
            'width': 8,
            'padding': (scale(8), scale(4))
        }
        self.button_style_2col = {
            'width': 10,
            'padding': (scale(8), scale(4))
        }
        self.button_style_1col = {
            'width': 16,
            'padding': (scale(8), scale(4))
        }
        
        self._create_widgets()
        
        logger.info("格式转换面板组件初始化完成")
    
    def _create_widgets(self):
        """创建界面元素"""
        logger.debug("创建格式转换面板界面元素")
        
        # 配置grid布局 - 三行布局：格式转换 + 另存为 + 校对/汇总
        self.grid_rowconfigure(0, weight=0)  # 格式转换section
        self.grid_rowconfigure(1, weight=0)  # 另存为section
        self.grid_rowconfigure(2, weight=0)  # 校对/汇总section
        self.grid_columnconfigure(0, weight=1)
        
        # 创建"格式转换"section
        self.conversion_frame = tb.Labelframe(
            self,
            text="格式转换",
            bootstyle="success"
        )
        self.conversion_frame.grid(row=0, column=0, sticky="nsew", pady=(0, scale(20)))
        
        # 配置格式转换框架网格权重
        self.conversion_frame.grid_rowconfigure(0, weight=1)
        self.conversion_frame.grid_columnconfigure(0, weight=1)
        
        # 创建格式转换按钮容器
        self.conversion_container = tb.Frame(self.conversion_frame, bootstyle="default")
        self.conversion_container.grid(row=0, column=0, sticky="nsew", padx=scale(10), pady=scale(10))
        
        # 创建"另存为"section
        self.saveas_frame = tb.Labelframe(
            self,
            text="另存为",
            bootstyle="danger"
        )
        self.saveas_frame.grid(row=1, column=0, sticky="nsew", pady=(0, scale(20)))
        
        # 配置另存为框架网格权重
        self.saveas_frame.grid_rowconfigure(0, weight=1)
        self.saveas_frame.grid_columnconfigure(0, weight=1)
        
        # 创建另存为按钮容器
        self.saveas_container = tb.Frame(self.saveas_frame, bootstyle="default")
        self.saveas_container.grid(row=0, column=0, sticky="nsew", padx=scale(10), pady=scale(10))
        
        # 创建"校对文档/汇总表格"section（动态显示）
        self.extra_frame = tb.Labelframe(
            self,
            text="",  # 标题将根据文件类型动态设置
            bootstyle="warning"
        )
        self.extra_frame.grid(row=2, column=0, sticky="nsew")
        self.extra_frame.grid_remove()  # 默认隐藏
        
        # 配置额外框架网格权重
        self.extra_frame.grid_rowconfigure(0, weight=1)
        self.extra_frame.grid_columnconfigure(0, weight=1)
        
        # 创建额外section的容器
        self.extra_container = tb.Frame(self.extra_frame, bootstyle="default")
        self.extra_container.grid(row=0, column=0, sticky="nsew", padx=scale(10), pady=scale(10))
        
        # 默认显示提示信息（放在格式转换section）
        self.hint_label = tb.Label(
            self.conversion_container,
            text="请先选择文件",
            font=("Microsoft YaHei UI", 9),
            bootstyle="secondary"
        )
        self.hint_label.pack(pady=20)
        
        logger.debug("格式转换面板界面元素创建完成")
    
    def set_file_info(self, category: str, current_format: str, file_path: str = None, file_list: list = None, ui_mode: str = 'single'):
        """
        设置文件信息并更新按钮显示
        
        参数:
            category: 文件类别 ('document', 'spreadsheet', 'image', 'layout')
            current_format: 当前文件格式 (如 'docx', 'png' 等)
            file_path: 文件路径（用于执行转换操作）
            file_list: 当前选项卡下的所有文件列表（批量模式使用）
            ui_mode: UI模式 ('single' 或 'batch')
        """
        logger.info(f"设置文件信息: 类别={category}, 当前格式={current_format}, 文件路径={file_path}, UI模式={ui_mode}")
        
        self.current_category = category
        self.current_format = current_format.lower() if current_format else None
        self.current_file_path = file_path
        self.file_list = file_list or []  # 保存文件列表
        self.ui_mode = ui_mode  # 保存UI模式
        
        # 清空现有按钮
        self._clear_buttons()
        
        # 隐藏提示标签
        self.hint_label.pack_forget()
        
        # 根据类别创建对应的格式按钮
        if category == 'document':
            self._create_document_buttons()
        elif category == 'spreadsheet':
            self._create_spreadsheet_buttons()
        elif category == 'image':
            self._create_image_buttons()
        elif category == 'layout':
            self._create_layout_buttons()
        else:
            logger.warning(f"未知的文件类别: {category}")
            self.hint_label.pack(pady=20)
        
        # 更新按钮状态
        self._update_button_states()
        
        logger.info(f"格式转换面板已更新，显示 {len(self.format_buttons)} 个格式按钮")
    
    def _clear_buttons(self):
        """清空所有格式按钮和选项"""
        logger.debug("清空格式按钮和选项")
        
        # 清空格式转换容器的所有子组件
        for widget in self.conversion_container.winfo_children():
            widget.destroy()
        
        # 清空另存为容器的所有子组件
        for widget in self.saveas_container.winfo_children():
            widget.destroy()
        
        # 清空额外section容器的所有子组件（修复：防止校对/汇总section残留）
        for widget in self.extra_container.winfo_children():
            widget.destroy()
        
        # 隐藏额外section框架（修复：防止框架残留显示）
        self.extra_frame.grid_remove()
        
        # 重置容器的grid列配置（避免列宽度问题）
        # 移除所有列的配置
        for col in range(10):  # 假设最多10列
            self.conversion_container.grid_columnconfigure(col, weight=0, uniform="")
            self.saveas_container.grid_columnconfigure(col, weight=0, uniform="")
            self.extra_container.grid_columnconfigure(col, weight=0, uniform="")
        
        # 清空按钮字典
        self.format_buttons.clear()
        
        # 清空校对选项和汇总选项变量（修复：防止状态残留）
        self.checkbox_vars.clear()
        self.merge_mode_var = None
        self.reference_table_var = None
    
    def _create_document_buttons(self):
        """创建文档类格式转换按钮 - 拆分为格式转换和另存为两个section"""
        logger.debug("创建文档类格式按钮 - 拆分布局")
        
        # 清空现有按钮
        self._clear_buttons()
        
        # 隐藏提示标签
        self.hint_label.pack_forget()
        
        # === 格式转换 section ===
        # 添加说明文本
        hint_label = tb.Label(
            self.conversion_container,
            text="转换为文档类型的其他格式",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置格式转换容器的网格布局为2列
        for col in range(2):
            self.conversion_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 配置行权重使按钮在垂直方向均匀分布
        for row in range(1, 3):  # 从row=1开始，因为row=0被说明文本占用
            self.conversion_container.grid_rowconfigure(row, weight=1)
        
        # 第一行按钮（实际在row=1）：转为DOCX、转为DOC
        formats_row1 = ['DOCX', 'DOC']
        for idx, fmt in enumerate(formats_row1):
            button = tb.Button(
                self.conversion_container,
                text=f"{fmt}",
                command=lambda f=fmt: self._on_format_clicked(f),
                bootstyle=self.button_colors[fmt],
                **self.button_style_2col
            )
            button.grid(row=1, column=idx, padx=scale(5), pady=scale(5), sticky="ew")
            self.format_buttons[fmt] = button
            logger.debug(f"  创建按钮: 转为{fmt} at (1, {idx}) - 颜色: {self.button_colors[fmt]}")
        
        # 第二行按钮（实际在row=2）：转为ODT、转为RTF
        formats_row2 = ['ODT', 'RTF']
        for idx, fmt in enumerate(formats_row2):
            button = tb.Button(
                self.conversion_container,
                text=f"{fmt}",
                command=lambda f=fmt: self._on_format_clicked(f),
                bootstyle=self.button_colors[fmt],
                **self.button_style_2col
            )
            button.grid(row=2, column=idx, padx=scale(5), pady=scale(5), sticky="ew")
            self.format_buttons[fmt] = button
            logger.debug(f"  创建按钮: 转为{fmt} at (2, {idx}) - 颜色: {self.button_colors[fmt]}")
        
        # === 另存为 section ===
        # 添加说明文本
        saveas_hint_label = tb.Label(
            self.saveas_container,
            text="转换为版式文件",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        saveas_hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置另存为容器的网格布局为2列
        for col in range(2):
            self.saveas_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 添加另存PDF按钮
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
        logger.debug("  创建按钮: 另存PDF (另存为section) - 颜色: danger")
        
        # 添加另存OFD按钮（禁用）
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
        logger.debug("  创建按钮: 另存OFD (另存为section) - 颜色: secondary (禁用)")
        
        # 更新按钮状态
        self._update_button_states()
        
        # 创建校对section
        self._create_document_validation_section()
    
    def _create_spreadsheet_buttons(self):
        """创建表格类格式转换按钮 - 拆分为格式转换和另存为两个section"""
        logger.debug("创建表格类格式按钮 - 拆分布局")
        
        # 清空现有按钮
        self._clear_buttons()
        
        # 隐藏提示标签
        self.hint_label.pack_forget()
        
        # === 格式转换 section ===
        # 添加说明文本
        hint_label = tb.Label(
            self.conversion_container,
            text="转换为表格类型的其他格式",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置格式转换容器的网格布局为2列
        for col in range(2):
            self.conversion_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 配置行权重使按钮在垂直方向均匀分布
        for row in range(1, 3):
            self.conversion_container.grid_rowconfigure(row, weight=1)
        
        # 第一行按钮（实际在row=1）：转为XLSX、转为XLS
        formats_row1 = ['XLSX', 'XLS']
        for idx, fmt in enumerate(formats_row1):
            button = tb.Button(
                self.conversion_container,
                text=f"{fmt}",
                command=lambda f=fmt: self._on_format_clicked(f),
                bootstyle=self.button_colors[fmt],
                **self.button_style_2col
            )
            button.grid(row=1, column=idx, padx=scale(5), pady=scale(5), sticky="ew")
            self.format_buttons[fmt] = button
            logger.debug(f"  创建按钮: 转为{fmt} at (1, {idx}) - 颜色: {self.button_colors[fmt]}")
        
        # 第二行按钮（实际在row=2）：转为ODS、转为CSV
        formats_row2 = ['ODS', 'CSV']
        for idx, fmt in enumerate(formats_row2):
            button = tb.Button(
                self.conversion_container,
                text=f"{fmt}",
                command=lambda f=fmt: self._on_format_clicked(f),
                bootstyle=self.button_colors[fmt],
                **self.button_style_2col
            )
            button.grid(row=2, column=idx, padx=scale(5), pady=scale(5), sticky="ew")
            self.format_buttons[fmt] = button
            logger.debug(f"  创建按钮: 转为{fmt} at (2, {idx}) - 颜色: {self.button_colors[fmt]}")
        
        # === 另存为 section ===
        # 添加说明文本
        saveas_hint_label = tb.Label(
            self.saveas_container,
            text="转换为版式文件",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        saveas_hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置另存为容器的网格布局为2列
        for col in range(2):
            self.saveas_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 添加另存PDF按钮
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
        logger.debug("  创建按钮: 另存PDF (另存为section) - 颜色: danger")
        
        # 添加另存OFD按钮（禁用）
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
        logger.debug("  创建按钮: 另存OFD (另存为section) - 颜色: secondary (禁用)")
        
        # 更新按钮状态
        self._update_button_states()
        
        # 创建汇总section
        self._create_spreadsheet_merge_section()
    
    def _create_image_buttons(self):
        """创建图片类格式转换按钮 - 拆分为格式转换和另存为两个section"""
        logger.debug("创建图片类格式按钮 - 拆分布局")
        
        # 清空现有按钮
        self._clear_buttons()
        
        # 隐藏提示标签
        self.hint_label.pack_forget()
        
        # === 格式转换 section ===
        # 添加说明文本
        hint_label = tb.Label(
            self.conversion_container,
            text="转换为图片类型的其他格式",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置格式转换容器的网格布局为2列
        for col in range(2):
            self.conversion_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 配置行权重使按钮在垂直方向均匀分布
        for row in range(1, 4):
            self.conversion_container.grid_rowconfigure(row, weight=1)
        
        # 第一行按钮（实际在row=1）：PNG、BMP
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
            logger.debug(f"  创建按钮: {fmt} at (1, {idx}) - 颜色: {self.button_colors[fmt]}")
        
        # 第二行按钮（实际在row=2）：GIF、TIF
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
            logger.debug(f"  创建按钮: {fmt} at (2, {idx}) - 颜色: {self.button_colors[fmt]}")
        
        # 第三行按钮（实际在row=3）：WebP、JPG
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
            logger.debug(f"  创建按钮: {fmt} at (3, {idx}) - 颜色: {self.button_colors[fmt]}")
        
        # === Row 4: 压缩选项 ===
        compress_options_frame = tb.Labelframe(
            self.conversion_container,
            text="压缩选项",
            bootstyle="info"
        )
        compress_options_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=5, pady=(10, 5))
        
        # 配置压缩选项框架网格权重
        compress_options_frame.grid_rowconfigure(0, weight=1)
        compress_options_frame.grid_rowconfigure(1, weight=1)
        compress_options_frame.grid_columnconfigure(0, weight=1)
        
        # === 单选按钮行 ===
        radio_frame = tb.Frame(compress_options_frame, bootstyle="default")
        radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=(scale(10), scale(5)))
        
        # 配置单选按钮框架
        radio_frame.grid_columnconfigure(0, weight=0)
        radio_frame.grid_columnconfigure(1, weight=0)
        radio_frame.grid_columnconfigure(2, weight=0)
        
        # 创建压缩模式变量（从配置读取默认值）
        default_compress_mode = 'lossless'
        if self.config_manager:
            try:
                default_compress_mode = self.config_manager.get_image_compress_mode()
            except Exception as e:
                logger.warning(f"读取压缩模式默认值失败: {e}")
        self.compress_mode_var = tk.StringVar(value=default_compress_mode)
        
        # 导入信息图标创建函数
        from gongwen_converter.utils.gui_utils import create_info_icon
        
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
        
        # 添加说明图标
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
        
        # === 输入行 ===
        input_frame = tb.Frame(compress_options_frame, bootstyle="default")
        input_frame.grid(row=1, column=0, sticky="ew", padx=scale(10), pady=(0, scale(10)))
        
        # 配置输入框架
        input_frame.grid_columnconfigure(0, weight=0)  # 标签
        input_frame.grid_columnconfigure(1, weight=1)  # 输入框
        input_frame.grid_columnconfigure(2, weight=0)  # 单位下拉
        
        # 获取小字体
        from gongwen_converter.utils.font_utils import get_small_font
        small_font, small_size = get_small_font()
        
        # 标签
        size_label = tb.Label(
            input_frame,
            text="文件大小上限：",
            font=(small_font, small_size),
            bootstyle="secondary"
        )
        size_label.grid(row=0, column=0, sticky="w", padx=(0, scale(5)))
        
        # 创建输入变量（从配置读取默认值）
        default_size_limit = 200
        default_size_unit = 'KB'
        if self.config_manager:
            try:
                default_size_limit = self.config_manager.get_image_size_limit()
                default_size_unit = self.config_manager.get_image_size_unit()
            except Exception as e:
                logger.warning(f"读取大小限制默认值失败: {e}")
        self.size_limit_var = tk.StringVar(value=str(default_size_limit))
        
        # 输入框
        self.size_limit_entry = tb.Entry(
            input_frame,
            textvariable=self.size_limit_var,
            bootstyle="default",
            font=(small_font, small_size),
            state='disabled'  # 初始禁用
        )
        self.size_limit_entry.grid(row=0, column=1, sticky="ew", padx=(0, scale(5)))
        
        # 创建单位下拉变量（从配置读取默认值）
        self.size_unit_var = tk.StringVar(value=default_size_unit)
        
        # 单位下拉菜单
        self.size_unit_combo = tb.Combobox(
            input_frame,
            textvariable=self.size_unit_var,
            values=['KB', 'MB'],
            bootstyle="default",
            font=(small_font, small_size),
            state='disabled',  # 初始禁用
            width=6
        )
        self.size_unit_combo.grid(row=0, column=2, sticky="ew")
        
        # 绑定输入验证
        self.size_limit_var.trace_add('write', lambda *args: self._on_size_input_changed())
        
        # === Row 2: 警告标签（动态显示）===
        warning_frame = tb.Frame(compress_options_frame, bootstyle="default")
        warning_frame.grid(row=2, column=0, sticky="ew", padx=scale(10), pady=(0, scale(5)))
        
        warning_frame.grid_columnconfigure(0, weight=1)
        
        # 文件大小警告标签
        self.size_warning_label = tb.Label(
            warning_frame,
            text="",
            font=(small_font, small_size),
            bootstyle="danger",
            anchor=tk.W
        )
        self.size_warning_label.grid(row=0, column=0, sticky="w")
        
        logger.debug("压缩选项UI创建完成")
        
        # === 另存为 section ===
        # 添加说明文本
        saveas_hint_label = tb.Label(
            self.saveas_container,
            text="转换为版式文件",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        saveas_hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置另存为容器的网格布局为2列
        for col in range(2):
            self.saveas_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 添加另存PDF按钮（默认启用，尺寸选项默认为原图嵌入）
        self.convert_image_to_pdf_button = tb.Button(
            self.saveas_container,
            text="另存为 PDF",
            command=self._on_convert_to_pdf_clicked,
            bootstyle=self.button_colors['PDF'],
            **self.button_style_2col
        )
        self.convert_image_to_pdf_button.grid(row=1, column=0, padx=scale(5), pady=(scale(5), scale(10)), sticky="ew")
        ToolTip(self.convert_image_to_pdf_button, "将图片嵌入PDF文件，支持选择尺寸选项")
        logger.debug("  创建按钮: 另存PDF (另存为section) - 颜色: danger")
        
        # 添加另存OFD按钮（禁用）
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
        logger.debug("  创建按钮: 另存OFD (另存为section) - 颜色: secondary (禁用)")
        
        # PDF尺寸选项
        size_options_frame = tb.Labelframe(
            self.saveas_container, 
            text="尺寸选项",
            bootstyle="info"
        )
        size_options_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置选项框架网格权重 (用于居中内部容器)
        size_options_frame.grid_rowconfigure(0, weight=1)
        size_options_frame.grid_columnconfigure(0, weight=1)
        
        # 创建单选按钮变量（从配置读取默认值）
        default_pdf_quality = 'original'
        if self.config_manager:
            try:
                default_pdf_quality = self.config_manager.get_image_pdf_quality()
            except Exception as e:
                logger.warning(f"读取PDF质量默认值失败: {e}")
        self.pdf_quality_var = tk.StringVar(value=default_pdf_quality)
        
        # 三个互斥单选按钮，放在边框内 - 优化对齐
        quality_radio_frame = tb.Frame(size_options_frame, bootstyle="default")
        quality_radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置单选按钮容器的列权重，确保平均分布
        quality_radio_frame.grid_columnconfigure(0, weight=1)
        quality_radio_frame.grid_columnconfigure(1, weight=1)
        quality_radio_frame.grid_columnconfigure(2, weight=1)
        
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
        
        # 更新按钮状态
        self._update_button_states()
        
        # 创建合并图片section
        self._create_image_merge_section()
        
        logger.debug("图片类格式按钮创建完成 - 拆分布局")
    
    def _create_layout_buttons(self):
        """创建版式类格式转换按钮 - 拆分为格式转换和另存为两个section"""
        logger.debug("创建版式类格式按钮 - 拆分布局")
        
        # 清空现有按钮
        self._clear_buttons()
        
        # 隐藏提示标签
        self.hint_label.pack_forget()
        
        # === 格式转换 section ===
        # 添加说明文本
        hint_label = tb.Label(
            self.conversion_container,
            text="转换为版式类型的其他格式",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 5))
        
        # 配置格式转换容器的网格布局为2列
        for col in range(2):
            self.conversion_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 配置行权重使按钮在垂直方向均匀分布
        self.conversion_container.grid_rowconfigure(1, weight=1)
        
        # 只显示两个按钮：转为PDF、转为OFD
        # PDF按钮（可用）
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
        
        # OFD按钮（禁用）
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
        
        # 第一行按钮（实际在row=1）：另存为DOCX、另存为DOC
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
            # 为DOCX/DOC按钮添加tooltip
            if fmt == 'DOCX':
                ToolTip(button, "优先使用Word/LibreOffice转换，备选使用内置工具")
            elif fmt == 'DOC':
                ToolTip(button, "先转为DOCX，再通过本地Office软件转为DOC格式")
            logger.debug(f"  创建按钮: 另存为{fmt} at (1, {idx}) - 颜色: {self.button_colors[fmt]}")
        
        # === Row 2: 分割线 ===
        separator = tb.Separator(self.saveas_container, bootstyle="danger")
        separator.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=scale(10))
        
        # === Row 3: 转为TIF、转为JPG 按钮 ===
        # 创建按钮引用属性
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
        ToolTip(self.convert_to_jpg_button, "将PDF页面渲染为JPG图片，支持选择DPI质量")
        logger.debug("  创建按钮: 渲染为JPG at (3, 1)")
        
        # === Row 4: DPI选项框 ===
        dpi_options_frame = tb.Labelframe(
            self.saveas_container,
            text="图片质量（DPI）",
            bootstyle="info"
        )
        dpi_options_frame.grid(row=4, column=0, columnspan=2, sticky="ew", padx=5, pady=(5, 5))
        
        # 配置选项框架网格权重
        dpi_options_frame.grid_rowconfigure(0, weight=1)
        dpi_options_frame.grid_columnconfigure(0, weight=1)
        
        # 创建单选按钮变量（从配置读取默认值）
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
        
        # 更新按钮状态
        self._update_button_states()
        
        # 创建合并拆分section
        self._create_layout_merge_split_section()
        
        logger.debug("版式类格式按钮创建完成 - 拆分布局")
    
    def _create_format_buttons(self, formats: List[str], columns: int = 3):
        """
        创建格式转换按钮
        
        参数:
            formats: 格式列表
            columns: 每行显示的列数
        """
        logger.debug(f"创建 {len(formats)} 个格式按钮，每行 {columns} 列")
        
        # 根据列数选择按钮样式
        if columns == 3:
            button_style = self.button_style_3col
        elif columns == 2:
            button_style = self.button_style_2col
        else:
            button_style = self.button_style_1col
        
        # 配置按钮容器的网格布局
        for col in range(columns):
            self.button_container.grid_columnconfigure(col, weight=1, uniform="format_col")
        
        # 创建按钮
        for idx, fmt in enumerate(formats):
            row = idx // columns
            col = idx % columns
            
            # 获取按钮颜色，如果没有定义则使用primary
            bootstyle = self.button_colors.get(fmt, 'primary')
            
            button = tb.Button(
                self.button_container,
                text=fmt,
                command=lambda f=fmt: self._on_format_clicked(f),
                bootstyle=bootstyle,
                **button_style
            )
            button.grid(row=row, column=col, padx=5, pady=5, sticky="ew")
            
            self.format_buttons[fmt] = button
            logger.debug(f"  创建按钮: {fmt} at ({row}, {col}) - 颜色: {bootstyle}")
    
    def _update_button_states(self):
        """根据当前格式更新按钮状态并更新tooltip"""
        logger.debug(f"更新按钮状态，当前格式: {self.current_format}, UI模式: {getattr(self, 'ui_mode', 'single')}")
        
        if not self.current_format:
            return
        
        # 检查是否选择了压缩模式
        is_compress_mode = (
            hasattr(self, 'compress_mode_var') and 
            self.compress_mode_var.get() == 'limit_size'
        )
        
        # 检查压缩模式下输入是否有效
        is_compress_input_valid = False
        if is_compress_mode and hasattr(self, 'size_limit_var'):
            input_text = self.size_limit_var.get()
            unit = self.size_unit_var.get() if hasattr(self, 'size_unit_var') else 'KB'
            is_compress_input_valid = self._validate_size_input(input_text, unit)
        
        # 判断是否为批量模式（根据UI模式，而不是文件数量）
        is_batch_mode = (getattr(self, 'ui_mode', 'single') == 'batch')
        
        # 获取所有相关格式
        all_formats = set()
        if is_batch_mode:
            # 批量模式：获取所有文件的格式
            from gongwen_converter.utils.file_type_utils import detect_actual_file_format
            for file_path in self.file_list:
                fmt = detect_actual_file_format(file_path)
                all_formats.add(self._normalize_format(fmt))
            logger.debug(f"批量模式（UI）：文件列表中的格式={all_formats}")
        else:
            # 单文件模式：只使用当前文件格式
            all_formats.add(self._normalize_format(self.current_format))
            logger.debug(f"单文件模式（UI）：当前文件格式={self.current_format}")
        
        # 判断是否所有文件格式一致
        all_same_format = (len(all_formats) == 1)
        logger.debug(f"所有文件格式一致: {all_same_format}, 批量模式: {is_batch_mode}")
        
        # 标准化当前格式（处理 jpg/jpeg、tif/tiff 等等价格式）
        normalized_current = self._normalize_format(self.current_format)
        
        for fmt, button in self.format_buttons.items():
            normalized_fmt = self._normalize_format(fmt.lower())
            
            # 修改判断逻辑：只有当所有文件格式一致且等于目标格式时才算同格式
            is_same_format = (all_same_format and normalized_fmt == normalized_current)
            
            # OFD格式始终保持禁用（暂不支持）
            if fmt.upper() == 'OFD':
                logger.debug(f"  跳过OFD按钮（保持禁用）")
                continue
            
            # 压缩模式且输入有效：只启用可压缩格式
            if is_compress_mode and is_compress_input_valid:
                if fmt.upper() in self.COMPRESSIBLE_FORMATS:
                    # JPEG/WebP：启用
                    original_color = self.button_colors.get(fmt, 'primary')
                    button.configure(state='normal', bootstyle=original_color)
                    # 添加压缩模式tooltip
                    tooltip_text = self._get_button_tooltip(fmt, is_same_format, 'normal')
                    ToolTip(button, tooltip_text)
                    logger.debug(f"  启用按钮: {fmt} (可压缩格式) - 颜色: {original_color}")
                else:
                    # 其他格式：禁用并添加tooltip
                    button.configure(state='disabled', bootstyle='secondary')
                    # 添加压缩模式tooltip（不可压缩）
                    tooltip_text = self._get_button_tooltip(fmt, is_same_format, 'disabled')
                    ToolTip(button, tooltip_text)
                    logger.debug(f"  禁用按钮: {fmt} (不支持有损压缩)")
            else:
                # 非压缩模式或输入无效：使用原有逻辑
                if is_same_format:
                    # 如果是压缩模式，同格式按钮启用（允许压缩）
                    if is_compress_mode:
                        original_color = self.button_colors.get(fmt, 'primary')
                        button.configure(state='normal', bootstyle=original_color)
                        # 添加tooltip
                        tooltip_text = self._get_button_tooltip(fmt, is_same_format, 'normal')
                        ToolTip(button, tooltip_text)
                        logger.debug(f"  启用按钮: {fmt} (压缩模式，允许同格式压缩) - 颜色: {original_color}")
                    else:
                        # 无损模式，同格式按钮禁用
                        button.configure(state='disabled', bootstyle='secondary')
                        # 添加同格式tooltip
                        tooltip_text = self._get_button_tooltip(fmt, is_same_format, 'disabled')
                        ToolTip(button, tooltip_text)
                        logger.debug(f"  禁用按钮: {fmt} (当前格式)")
                else:
                    # 其他格式按钮启用，恢复原有颜色
                    original_color = self.button_colors.get(fmt, 'primary')
                    button.configure(state='normal', bootstyle=original_color)
                    # 添加最高质量模式tooltip
                    tooltip_text = self._get_button_tooltip(fmt, is_same_format, 'normal')
                    ToolTip(button, tooltip_text)
                    logger.debug(f"  启用按钮: {fmt} - 颜色: {original_color}")
    
    def _normalize_format(self, fmt: str) -> str:
        """
        标准化格式名称（处理等价格式）
        
        参数:
            fmt: 格式名称
            
        返回:
            str: 标准化后的格式名称
        """
        fmt = fmt.lower()
        
        # 等价格式映射
        equivalents = {
            'jpg': 'jpeg',
            'jpeg': 'jpeg',
            'tif': 'tiff',
            'tiff': 'tiff',
            'heif': 'heic',
            'heic': 'heic'
        }
        
        return equivalents.get(fmt, fmt)
    
    def _on_format_clicked(self, target_format: str):
        """
        格式按钮点击事件处理
        
        参数:
            target_format: 目标格式
        """
        logger.info(f"格式按钮被点击: {target_format}")
        
        # 优先使用 on_action 执行实际转换
        if self.on_action and self.current_file_path:
            # 获取源文件的实际格式
            from gongwen_converter.utils.file_type_utils import detect_actual_file_format
            source_format = detect_actual_file_format(self.current_file_path)
            
            # 构建 action_type
            target_format_lower = target_format.lower()
            
            # 特殊处理1：PDF转换使用类别名而不是具体格式名
            if target_format_lower == 'pdf':
                # 使用类别名构建策略名称（document/spreadsheet/image/layout）
                action_type = f"convert_{self.current_category}_to_{target_format_lower}"
            # 特殊处理2：layout类别转换到文档格式（DOCX/DOC/ODT/RTF）也使用类别名
            elif self.current_category == 'layout' and target_format_lower in ['docx', 'doc', 'odt', 'rtf']:
                # 版式文件转文档格式：convert_layout_to_{docx|doc|odt|rtf}
                action_type = f"convert_{self.current_category}_to_{target_format_lower}"
            else:
                # 其他格式转换使用具体格式名
                action_type = f"convert_{source_format}_to_{target_format_lower}"
            
            logger.info(f"执行格式转换: {source_format.upper()} → {target_format_lower.upper()}, 策略: {action_type}")
            
            # 构建转换选项，包含目标格式
            options = {
                'target_format': target_format_lower
            }
            
            # 如果是图片类别，添加压缩选项
            if self.current_category == 'image' and hasattr(self, 'compress_mode_var'):
                compress_mode = self.compress_mode_var.get()
                options['compress_mode'] = compress_mode
                logger.debug(f"添加压缩模式: {compress_mode}")
                
                if compress_mode == 'limit_size':
                    # 验证输入有效性
                    if hasattr(self, 'size_limit_var') and hasattr(self, 'size_unit_var'):
                        try:
                            size_limit = int(self.size_limit_var.get())
                            size_unit = self.size_unit_var.get()
                            options['size_limit'] = size_limit
                            options['size_unit'] = size_unit
                            logger.info(f"压缩选项: 限制大小 {size_limit}{size_unit}")
                        except (ValueError, AttributeError) as e:
                            logger.warning(f"无法获取压缩选项: {e}")
                else:
                    logger.info("压缩选项: 最高质量模式")
            
            # 调用 on_action 执行转换
            self.on_action(action_type, self.current_file_path, options)
        
        # 回退到旧的 on_format_selected 接口（向后兼容）
        elif self.on_format_selected:
            logger.debug("使用旧的 on_format_selected 接口")
            self.on_format_selected(target_format)
        
        else:
            logger.warning("没有可用的回调函数来处理格式转换")
    
    def _on_convert_to_pdf_clicked(self):
        """处理图片转PDF按钮点击事件（包含质量模式选项）"""
        if self.on_action and self.current_file_path:
            # 获取选中的质量模式
            quality_mode = self.pdf_quality_var.get() if hasattr(self, 'pdf_quality_var') else ''
            if not quality_mode:
                logger.warning("未选择PDF尺寸选项")
                return
            
            options = {'quality_mode': quality_mode}
            logger.info(f"图片转PDF，质量模式: {quality_mode}")
            self.on_action("convert_image_to_pdf", self.current_file_path, options)
    
    def _on_convert_to_image_clicked(self, image_format: str):
        """处理版式文件转图片按钮点击事件（PNG/JPG）"""
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
    
    def _on_pdf_quality_changed(self):
        """处理PDF尺寸选项变更事件"""
        if hasattr(self, 'pdf_quality_var') and hasattr(self, 'convert_image_to_pdf_button'):
            quality = self.pdf_quality_var.get()
            logger.debug(f"PDF尺寸选项已变更: {quality}")
            
            # 更新PDF按钮状态：只有当选择了尺寸选项时才启用按钮
            should_enable = bool(quality)  # 非空即启用
            self.convert_image_to_pdf_button.config(state="normal" if should_enable else "disabled")
            
            if should_enable:
                logger.debug(f"PDF按钮已启用：已选择尺寸选项 ({quality})")
            else:
                logger.debug("PDF按钮保持禁用：未选择尺寸选项")
    
    def _get_button_tooltip(self, fmt: str, is_same_format: bool, button_state: str) -> str:
        """
        根据当前状态获取按钮的tooltip文本
        
        参数:
            fmt: 格式名称（大写）
            is_same_format: 是否与当前格式相同
            button_state: 按钮状态 ('normal' 或 'disabled')
            
        返回:
            str: tooltip文本
        """
        # 格式标准化
        fmt_upper = fmt.upper()
        
        # 如果是同格式且禁用
        if is_same_format and button_state == 'disabled':
            return f"当前已是{fmt_upper}格式，无需转换"
        
        # 根据文件类别返回不同的tooltip
        if self.current_category == 'image':
            return self._get_image_button_tooltip(fmt_upper, is_same_format, button_state)
        elif self.current_category == 'document':
            return self._get_document_button_tooltip(fmt_upper)
        elif self.current_category == 'spreadsheet':
            return self._get_spreadsheet_button_tooltip(fmt_upper)
        elif self.current_category == 'layout':
            return self._get_layout_button_tooltip(fmt_upper)
        else:
            return ""
    
    def _get_image_button_tooltip(self, fmt_upper: str, is_same_format: bool, button_state: str) -> str:
        """
        获取图片格式按钮的tooltip
        
        参数:
            fmt_upper: 格式名称（大写）
            is_same_format: 是否与当前格式相同
            button_state: 按钮状态
            
        返回:
            str: tooltip文本
        """
        # 检查是否为压缩模式
        is_compress_mode = (
            hasattr(self, 'compress_mode_var') and 
            self.compress_mode_var.get() == 'limit_size'
        )
        
        # 压缩模式
        if is_compress_mode:
            # 可压缩格式（JPEG/WebP）
            if fmt_upper in self.COMPRESSIBLE_FORMATS:
                return "支持有损压缩至目标大小\n通过调整quality参数自动压缩"
            # 不可压缩格式（禁用状态）
            else:
                if fmt_upper == 'PNG':
                    return "PNG只支持无损压缩，无法有效减小文件\n建议选择JPEG或WebP格式"
                elif fmt_upper == 'BMP':
                    return "BMP不支持压缩，无法减小文件\n建议选择JPEG或WebP格式"
                elif fmt_upper == 'GIF':
                    return "GIF只支持有限压缩，无法有效减小文件\n建议选择JPEG或WebP格式"
                elif fmt_upper in ['TIF', 'TIFF']:
                    return "TIF只支持有限压缩，无法有效减小文件\n建议选择JPEG或WebP格式"
                else:
                    return f"{fmt_upper}格式无法有效压缩至目标大小\n建议选择JPEG或WebP格式"
        
        # 最高质量模式
        else:
            # JPEG/WebP
            if fmt_upper in ['JPG', 'JPEG', 'WEBP']:
                return "使用最高质量设置转换（quality=95）"
            # 无损格式
            else:
                return "无损转换，保持完美质量"
    
    def _get_document_button_tooltip(self, fmt_upper: str) -> str:
        """
        获取文档格式按钮的tooltip
        
        参数:
            fmt_upper: 格式名称（大写）
            
        返回:
            str: tooltip文本
        """
        return "可能需要通过本地Office软件（WPS/Word/LibreOffice）进行转换"
    
    def _get_spreadsheet_button_tooltip(self, fmt_upper: str) -> str:
        """
        获取表格格式按钮的tooltip
        
        参数:
            fmt_upper: 格式名称（大写）
            
        返回:
            str: tooltip文本
        """
        if fmt_upper == 'CSV':
            return "可能需要通过本地Office软件进行预处理，再转换为CSV文本格式"
        else:
            return "可能需要通过本地Office软件（WPS/Excel/LibreOffice）进行转换"
    
    def _get_layout_button_tooltip(self, fmt_upper: str) -> str:
        """
        获取版式格式按钮的tooltip
        
        参数:
            fmt_upper: 格式名称（大写）
            
        返回:
            str: tooltip文本
        """
        if fmt_upper == 'PDF':
            # 根据实际文件格式判断
            if self.current_format and self.current_format.lower() in ['ofd', 'xps']:
                return "使用内置库将OFD/XPS转换为PDF格式"
            else:
                return "当前已是PDF格式，无需转换"
        elif fmt_upper == 'DOCX':
            return "优先使用Word/LibreOffice转换，备选使用内置内置工具"
        elif fmt_upper == 'DOC':
            return "先转为DOCX，再通过本地Office软件转为DOC格式"
        elif fmt_upper == 'TIF':
            return "将PDF页面渲染为TIF图片，支持多页TIFF，支持选择DPI质量"
        elif fmt_upper == 'JPG':
            return "将PDF页面渲染为JPG图片，支持选择DPI质量"
        else:
            return ""
    
    def _create_document_validation_section(self):
        """创建文档校对section"""
        logger.debug("创建文档校对section")
        
        # 显示extra_frame并设置标题
        self.extra_frame.config(text="校对文档")
        self.extra_frame.grid()
        
        # 清空extra_container
        for widget in self.extra_container.winfo_children():
            widget.destroy()
        
        # 配置容器的grid布局
        self.extra_container.grid_rowconfigure(0, weight=0)  # 说明文字
        self.extra_container.grid_rowconfigure(1, weight=0)  # 校对按钮
        self.extra_container.grid_rowconfigure(2, weight=0)  # 校对选项
        self.extra_container.grid_columnconfigure(0, weight=1)
        
        # 添加说明文字
        hint_label = tb.Label(
            self.extra_container,
            text="根据选项对文档进行校对",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, sticky="ew", padx=5, pady=(0, 5))
        
        # 创建校对按钮
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
        
        # 创建校对选项边框
        validation_options_frame = tb.Labelframe(
            self.extra_container,
            text="校对选项",
            bootstyle="info"
        )
        validation_options_frame.grid(row=2, column=0, sticky="ew", padx=scale(5), pady=scale(5))
        
        # 配置选项框架网格权重
        validation_options_frame.grid_rowconfigure(0, weight=1)
        validation_options_frame.grid_columnconfigure(0, weight=1)
        
        # 创建复选框容器
        checkbox_container = tb.Frame(validation_options_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置复选框容器的列权重
        checkbox_container.grid_columnconfigure(0, weight=1)
        checkbox_container.grid_columnconfigure(1, weight=1)
        
        # 获取默认选项
        default_options = self._get_default_validation_options()
        
        # 创建四个校对选项（包含说明文本）
        options = [
            ("标点配对", "symbol_pairing", "检查文档中的标点符号配对情况，如括号、引号等是否成对出现"),
            ("错别字校对", "typos_rule", "检查文档中的错别字，根据用户自定义的错别字词库进行匹配"),
            ("符号校对", "symbol_correction", "检查文档中的标点符号使用规范，如全角和半角符号误用等"),
            ("敏感词匹配", "sensitive_word", "检查文档中的敏感词，根据用户自定义的敏感词库进行匹配")
        ]
        
        # 导入信息图标创建函数
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        self.checkbox_vars = {}
        for i, (text, key, tooltip_text) in enumerate(options):
            var = tk.BooleanVar(value=default_options.get(key, False))
            self.checkbox_vars[key] = var
            
            # 创建容器用于放置复选框和信息图标
            option_frame = tb.Frame(checkbox_container, bootstyle="default")
            row, col = divmod(i, 2)
            option_frame.grid(row=row, column=col, sticky="", padx=scale(10), pady=scale(5))
            
            # 创建复选框
            checkbox = tb.Checkbutton(
                option_frame,
                text=text,
                variable=var,
                command=self._on_validation_option_changed,
                bootstyle="round-toggle"
            )
            checkbox.pack(side=tk.LEFT, padx=(0, scale(5)))
            
            # 添加信息图标
            info_icon = create_info_icon(option_frame, tooltip_text, bootstyle="info")
            info_icon.pack(side=tk.LEFT)
        
        # 初始化按钮状态
        self._on_validation_option_changed()
        
        logger.debug("文档校对section创建完成")
    
    def _create_spreadsheet_merge_section(self):
        """创建表格汇总section"""
        logger.debug("创建表格汇总section")
        
        # 显示extra_frame并设置标题
        self.extra_frame.config(text="汇总表格")
        self.extra_frame.grid()
        
        # 清空extra_container
        for widget in self.extra_container.winfo_children():
            widget.destroy()
        
        # 配置容器的grid布局
        self.extra_container.grid_rowconfigure(0, weight=0)  # 说明文字
        self.extra_container.grid_rowconfigure(1, weight=0)  # 汇总按钮
        self.extra_container.grid_rowconfigure(2, weight=0)  # 汇总选项
        self.extra_container.grid_rowconfigure(3, weight=0)  # 基准表格显示
        self.extra_container.grid_columnconfigure(0, weight=1)
        
        # 添加说明文字
        hint_label = tb.Label(
            self.extra_container,
            text="将多个表格按选择的模式汇总",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, sticky="ew", padx=5, pady=(0, 5))
        
        # 创建汇总按钮（默认启用）
        self.merge_tables_button = tb.Button(
            self.extra_container,
            text="📋 汇总表格",
            command=self._on_merge_tables_clicked,
            bootstyle='info',
            **self.button_style_1col
        )
        self.merge_tables_button.grid(row=1, column=0, pady=scale(5))
        ToolTip(
            self.merge_tables_button,
            "将批量列表中的多个表格按选择的模式汇总到基准表格。需先选择汇总模式（按行/按列/按单元格）"
        )
        
        # 创建汇总选项边框
        merge_options_frame = tb.Labelframe(
            self.extra_container,
            text="汇总选项",
            bootstyle="info"
        )
        merge_options_frame.grid(row=2, column=0, sticky="ew", padx=scale(5), pady=scale(5))
        
        # 配置选项框架网格权重
        merge_options_frame.grid_rowconfigure(0, weight=1)
        merge_options_frame.grid_columnconfigure(0, weight=1)
        
        # 创建互斥的单选按钮变量（从配置读取默认值）
        default_merge_mode = 3
        if self.config_manager:
            try:
                default_merge_mode = self.config_manager.get_spreadsheet_merge_mode()
            except Exception as e:
                logger.warning(f"读取汇总模式默认值失败: {e}")
        self.merge_mode_var = tk.IntVar(value=default_merge_mode)  # 0=未选择, 1=按行, 2=按列, 3=按单元格
        
        # 单选按钮容器
        merge_radio_frame = tb.Frame(merge_options_frame, bootstyle="default")
        merge_radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置单选按钮容器的列权重
        merge_radio_frame.grid_columnconfigure(0, weight=0)
        merge_radio_frame.grid_columnconfigure(1, weight=0)
        merge_radio_frame.grid_columnconfigure(2, weight=0)
        merge_radio_frame.grid_columnconfigure(3, weight=0)
        
        # 导入信息图标创建函数
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 左侧第1行：按行汇总 + 图标
        self.merge_by_row_radio = tb.Radiobutton(
            merge_radio_frame,
            text="按行汇总",
            variable=self.merge_mode_var,
            value=1,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        )
        self.merge_by_row_radio.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(5)))
        
        merge_row_info = create_info_icon(
            merge_radio_frame,
            "将一个或多个表格，逐行合并至基准表格。",
            bootstyle="info"
        )
        merge_row_info.grid(row=0, column=1, sticky="w", padx=(0, scale(20)))
        
        # 左侧第2行：按列汇总 + 图标
        self.merge_by_column_radio = tb.Radiobutton(
            merge_radio_frame,
            text="按列汇总",
            variable=self.merge_mode_var,
            value=2,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        )
        self.merge_by_column_radio.grid(row=1, column=0, sticky="w", padx=(scale(10), scale(5)), pady=(scale(5), 0))
        
        merge_col_info = create_info_icon(
            merge_radio_frame,
            "将一个或多个表格，逐列合并至基准表格。",
            bootstyle="info"
        )
        merge_col_info.grid(row=1, column=1, sticky="w", padx=(0, scale(20)), pady=(scale(5), 0))
        
        # 右侧：按单元格汇总 + 图标（跨2行，垂直居中）
        self.merge_by_cell_radio = tb.Radiobutton(
            merge_radio_frame,
            text="按单元格汇总",
            variable=self.merge_mode_var,
            value=3,
            command=self._on_merge_mode_changed,
            bootstyle="primary"
        )
        self.merge_by_cell_radio.grid(row=0, column=2, rowspan=2, sticky="ns", padx=(0, scale(5)))
        
        merge_cell_info = create_info_icon(
            merge_radio_frame,
            "对应位置的单元格，数据相加。",
            bootstyle="info"
        )
        merge_cell_info.grid(row=0, column=3, rowspan=2, sticky="ns", padx=(0, scale(10)))
        
        # 创建基准表格显示（移到汇总选项之后）
        reference_table_frame = tb.Frame(self.extra_container, bootstyle="default")
        reference_table_frame.grid(row=3, column=0, sticky="ew", padx=scale(5), pady=(scale(10), scale(5)))
        
        # 配置框架网格
        reference_table_frame.grid_columnconfigure(0, weight=1)
        
        # 获取小字体
        from gongwen_converter.utils.font_utils import get_small_font
        small_font, small_size = get_small_font()
        
        # 标签："当前选中的基准表格："
        reference_label = tb.Label(
            reference_table_frame,
            text="当前选中的基准表格：",
            font=(small_font, small_size),
            bootstyle="warning",
            anchor=tk.CENTER
        )
        reference_label.grid(row=0, column=0, sticky="ew")
        
        # 基准表格文件名显示（单独一行，支持自动换行）
        self.reference_table_var = tk.StringVar(value="未选择")
        self.reference_table_label = tb.Label(
            reference_table_frame,
            textvariable=self.reference_table_var,
            font=(small_font, small_size),
            bootstyle="primary",
            anchor=tk.CENTER,
            wraplength=250  # 设置自动换行宽度
        )
        self.reference_table_label.grid(row=1, column=0, sticky="ew")
        
        logger.debug("表格汇总section创建完成")
    
    def _create_layout_merge_split_section(self):
        """创建版式文件合并拆分section"""
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
        for row in range(8):  # 8行：说明文字、合并按钮、拆分按钮、分割线、拆分输入、帮助文本、文件信息、警告
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
        
        # 获取小字体
        from gongwen_converter.utils.font_utils import get_small_font
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
        self.page_input_entry.grid(row=0, column=1, sticky="ew")
        
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
        
        # === Row 5: 帮助文本 ===
        help_label = tb.Label(
            self.extra_container,
            text="支持格式：1-3,5-10 或 1~3;5;7至10",
            font=(small_font, small_size),
            bootstyle="secondary",
            anchor=tk.W
        )
        help_label.grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, scale(5)))
        
        # === Row 6: 文件信息和警告 ===
        info_frame = tb.Frame(self.extra_container, bootstyle="default")
        info_frame.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(scale(10), scale(5)))
        
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
    
    def _get_default_validation_options(self) -> Dict[str, bool]:
        """从配置文件获取默认校对选项"""
        if not self.config_manager:
            return {
                "symbol_pairing": True,
                "symbol_correction": True,
                "typos_rule": True,
                "sensitive_word": False
            }
        
        try:
            symbol_settings = self.config_manager.get_symbol_engine_settings()
            typos_settings = self.config_manager.get_typos_engine_settings()
            sensitive_settings = self.config_manager.get_sensitive_words_engine_settings()
            
            return {
                "symbol_pairing": symbol_settings.get("enable_symbol_pairing", True),
                "symbol_correction": symbol_settings.get("enable_symbol_correction", True),
                "typos_rule": typos_settings.get("enable_typos_rule", True),
                "sensitive_word": sensitive_settings.get("enable_sensitive_word", True),
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
    
    def _on_merge_mode_changed(self):
        """处理汇总模式变更事件"""
        if hasattr(self, 'merge_tables_button') and self.merge_tables_button:
            mode = self.merge_mode_var.get() if hasattr(self, 'merge_mode_var') else 0
            # 暂时只根据模式启用/禁用，实际使用时还需考虑批量模式
            should_enable = (mode > 0)
            self.merge_tables_button.config(state="normal" if should_enable else "disabled")
            logger.debug(f"汇总按钮状态: {'启用' if should_enable else '禁用'} (模式={mode})")
    
    def _on_validate_clicked(self):
        """处理校对按钮点击事件"""
        if self.on_action and self.current_file_path:
            options = {key: var.get() for key, var in self.checkbox_vars.items()}
            logger.info(f"执行校对操作，选项: {options}")
            self.on_action("validate", self.current_file_path, options)
    
    def _on_merge_tables_clicked(self):
        """处理汇总表格按钮点击事件"""
        if self.on_action and self.current_file_path:
            mode = self.merge_mode_var.get() if hasattr(self, 'merge_mode_var') else 0
            options = {"mode": mode}
            logger.info(f"执行汇总表格操作，模式: {mode}")
            self.on_action("merge_tables", self.current_file_path, options)
    
    def _create_image_merge_section(self):
        """创建图片合并section"""
        logger.debug("创建图片合并section")
        
        # 显示extra_frame并设置标题
        self.extra_frame.config(text="合并图片")
        self.extra_frame.grid()
        
        # 清空extra_container
        for widget in self.extra_container.winfo_children():
            widget.destroy()
        
        # 配置容器的grid布局
        self.extra_container.grid_rowconfigure(0, weight=0)  # 说明文字
        self.extra_container.grid_rowconfigure(1, weight=0)  # 合并按钮
        self.extra_container.grid_rowconfigure(2, weight=0)  # 转换选项
        self.extra_container.grid_columnconfigure(0, weight=1)
        
        # 添加说明文字
        hint_label = tb.Label(
            self.extra_container,
            text="将图片合并为多页TIFF文件",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        hint_label.grid(row=0, column=0, sticky="ew", padx=5, pady=(0, 5))
        
        # 创建合并按钮
        self.merge_tiff_button = tb.Button(
            self.extra_container,
            text="📋 合并为 TIF",
            command=self._on_merge_tiff_clicked,
            bootstyle='info',
            **self.button_style_1col
        )
        self.merge_tiff_button.grid(row=1, column=0, pady=scale(5))
        
        # 创建转换选项边框
        tiff_options_frame = tb.Labelframe(
            self.extra_container,
            text="转换选项",
            bootstyle="info"
        )
        tiff_options_frame.grid(row=2, column=0, sticky="ew", padx=scale(5), pady=scale(5))
        
        # 配置选项框架网格权重
        tiff_options_frame.grid_rowconfigure(0, weight=1)
        tiff_options_frame.grid_columnconfigure(0, weight=1)
        
        # 创建单选按钮变量（从配置读取默认值）
        default_tiff_mode = "smart"
        if self.config_manager:
            try:
                default_tiff_mode = self.config_manager.get_image_tiff_mode()
            except Exception as e:
                logger.warning(f"读取TIFF模式默认值失败: {e}")
        self.tiff_mode_var = tk.StringVar(value=default_tiff_mode)
        
        # 单选按钮容器
        tiff_radio_frame = tb.Frame(tiff_options_frame, bootstyle="default")
        tiff_radio_frame.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置单选按钮容器的列权重
        tiff_radio_frame.grid_columnconfigure(0, weight=0)
        tiff_radio_frame.grid_columnconfigure(1, weight=0)
        
        # 导入信息图标创建函数
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 选项1：保留透明（推荐，默认）+ 信息图标
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
        
        # 选项2：不保留透明 + 信息图标
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
    
    def _on_merge_tiff_clicked(self):
        """处理合并为TIFF按钮点击事件"""
        if self.on_action and self.current_file_path:
            mode = self.tiff_mode_var.get() if hasattr(self, 'tiff_mode_var') else "smart"
            options = {"mode": mode}
            logger.info(f"执行合并为TIFF操作，模式: {mode}")
            self.on_action("merge_images_to_tiff", self.current_file_path, options)
    
    def set_reference_table(self, file_name: str):
        """设置当前选中的基准表格文件名"""
        if hasattr(self, 'reference_table_var') and self.reference_table_var:
            if file_name:
                self.reference_table_var.set(file_name)
                logger.debug(f"设置基准表格: {file_name}")
            else:
                self.reference_table_var.set("未选择")
                logger.debug("清空基准表格显示")
    
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
    
    def _parse_page_ranges(self, input_text: str) -> List[int]:
        """
        解析页码范围字符串为页码列表
        
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
    
    def _show_page_warning(self, message: str):
        """显示页码警告信息"""
        if hasattr(self, 'page_warning_label') and self.page_warning_label:
            self.page_warning_label.config(text=message)
            logger.debug(f"显示页码警告: {message}")
    
    def _clear_page_warning(self):
        """清除页码警告信息"""
        if hasattr(self, 'page_warning_label') and self.page_warning_label:
            self.page_warning_label.config(text="")
    
    def _on_page_input_changed(self, event=None):
        """页码输入变更事件处理"""
        if not hasattr(self, 'page_input_var') or not self.page_input_var:
            return
        
        input_text = self.page_input_var.get()
        
        # 空输入：恢复默认样式
        if not input_text.strip():
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
    
    def _on_merge_pdfs_clicked(self):
        """合并PDF按钮点击事件"""
        if self.on_action and self.current_file_path:
            logger.info("执行合并PDF操作")
            self.on_action("merge_pdfs", self.current_file_path, {})
    
    def _on_split_pdf_clicked(self):
        """拆分PDF按钮点击事件"""
        if self.on_action and self.current_file_path:
            # 获取并解析页码
            input_text = self.page_input_var.get() if hasattr(self, 'page_input_var') and self.page_input_var else ""
            if not input_text:
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
    
    def _on_compress_mode_changed(self):
        """
        处理压缩模式切换事件
        
        当用户切换无损/压缩模式时：
        1. 启用/禁用输入框和下拉菜单
        2. 重新更新按钮状态（支持同格式压缩）
        3. 验证输入并更新警告
        """
        if not hasattr(self, 'compress_mode_var'):
            return
        
        mode = self.compress_mode_var.get()
        is_limit = (mode == 'limit_size')
        
        logger.debug(f"压缩模式变更: {mode}")
        
        # 1. 启用/禁用输入框和下拉菜单
        if hasattr(self, 'size_limit_entry') and hasattr(self, 'size_unit_combo'):
            state = 'normal' if is_limit else 'disabled'
            self.size_limit_entry.config(state=state)
            self.size_unit_combo.config(state=state)
            
            if is_limit:
                # 启用时：恢复默认值200（如果当前为空）
                if not self.size_limit_var.get():
                    self.size_limit_var.set("200")
                self.size_limit_entry.focus()
                # 立即验证输入
                self._on_size_input_changed()
            else:
                # 禁用时：清空警告和输入
                self._clear_size_warning()
                # 注意：不清空输入值，保留用户输入
        
        # 2. 重新更新按钮状态（关键！支持同格式压缩）
        self._update_button_states()
        
        logger.debug(f"压缩选项状态已更新: {'启用' if is_limit else '禁用'}")
    
    def _on_size_input_changed(self):
        """
        处理文件大小输入变更事件
        
        验证输入是否合法：
        - 只允许正整数
        - KB: 1-10240 (1KB-10MB)
        - MB: 1-100 (1MB-100MB)
        """
        if not hasattr(self, 'size_limit_var') or not hasattr(self, 'size_limit_entry'):
            return
        
        input_text = self.size_limit_var.get()
        
        # 空输入：显示警告并禁用按钮
        if not input_text.strip():
            self._show_size_warning("⚠️ 请输入文件大小上限")
            self.size_limit_entry.configure(bootstyle="danger")
            self._disable_format_buttons_for_compress()
            return
        
        # 获取单位
        unit = self.size_unit_var.get() if hasattr(self, 'size_unit_var') else 'KB'
        
        # 验证输入
        is_valid = self._validate_size_input(input_text, unit)
        
        # 更新输入框样式和警告
        if is_valid:
            self.size_limit_entry.configure(bootstyle="success")
            self._clear_size_warning()
            self._update_button_states()  # 恢复按钮状态
        else:
            self.size_limit_entry.configure(bootstyle="danger")
            # 显示具体错误信息
            try:
                size = int(input_text)
                if unit == 'KB':
                    self._show_size_warning("❌ KB范围：1-10240")
                else:
                    self._show_size_warning("❌ MB范围：1-100")
            except ValueError:
                self._show_size_warning("❌ 请输入有效的数字")
            self._disable_format_buttons_for_compress()
    
    def _validate_size_input(self, value: str, unit: str) -> bool:
        """
        验证文件大小输入
        
        参数:
            value: 输入值
            unit: 单位 (KB或MB)
            
        返回:
            bool: 是否有效
        """
        try:
            size = int(value)
            if unit == 'KB':
                return 1 <= size <= 10240  # 1KB-10MB
            else:  # MB
                return 1 <= size <= 100     # 1MB-100MB
        except ValueError:
            return False
    
    def _show_size_warning(self, message: str):
        """显示文件大小警告信息"""
        if hasattr(self, 'size_warning_label') and self.size_warning_label:
            self.size_warning_label.config(text=message)
            logger.debug(f"显示大小警告: {message}")
    
    def _clear_size_warning(self):
        """清除文件大小警告信息"""
        if hasattr(self, 'size_warning_label') and self.size_warning_label:
            self.size_warning_label.config(text="")
    
    def _disable_format_buttons_for_compress(self):
        """在压缩模式下输入无效时禁用所有格式按钮"""
        # 禁用所有格式按钮（除了OFD，它本来就是禁用的）
        for fmt, button in self.format_buttons.items():
            if fmt.upper() == 'OFD':
                continue  # 保持OFD禁用状态
            
            # 禁用所有按钮
            button.configure(state='disabled', bootstyle='secondary')
            logger.debug(f"  禁用按钮: {fmt} (压缩模式输入无效)")
    
    def show(self):
        """显示格式转换面板"""
        self.grid(row=0, column=0, sticky="nsew")
    
    def hide(self):
        """隐藏格式转换面板"""
        self.grid_remove()
    
    def reset(self):
        """重置面板状态"""
        logger.debug("重置格式转换面板")
        
        self.current_category = None
        self.current_format = None
        
        self._clear_buttons()
        
        # 显示提示标签
        self.hint_label.pack(pady=20)
        
        logger.info("格式转换面板已重置")


# 测试代码
if __name__ == "__main__":
    import tkinter as tk
    
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    root = tb.Window(title="格式转换面板测试", themename="morph")
    root.geometry("300x400")
    
    def on_format_selected(target_format):
        logger.info(f"选择目标格式: {target_format}")
        print(f"选择目标格式: {target_format}")
    
    panel = ConversionPanel(root, on_format_selected=on_format_selected)
    panel.pack(fill="both", expand=True, padx=10, pady=10)
    
    # 测试按钮
    test_frame = tb.Frame(root)
    test_frame.pack(fill="x", padx=10, pady=10)
    
    tb.Button(
        test_frame,
        text="测试文档类(DOCX)",
        command=lambda: panel.set_file_info('document', 'docx'),
        bootstyle="info"
    ).pack(side="left", padx=5)
    
    tb.Button(
        test_frame,
        text="测试表格类(XLSX)",
        command=lambda: panel.set_file_info('spreadsheet', 'xlsx'),
        bootstyle="info"
    ).pack(side="left", padx=5)
    
    tb.Button(
        test_frame,
        text="测试图片类(PNG)",
        command=lambda: panel.set_file_info('image', 'png'),
        bootstyle="info"
    ).pack(side="left", padx=5)
    
    tb.Button(
        test_frame,
        text="重置",
        command=panel.reset,
        bootstyle="warning"
    ).pack(side="left", padx=5)
    
    logger.info("启动格式转换面板测试")
    root.mainloop()
