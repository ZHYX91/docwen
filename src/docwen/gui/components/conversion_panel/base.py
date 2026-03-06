"""
格式转换面板基类模块

提供 ConversionPanelBase 基类，定义所有共享属性和公共方法。

主要功能：
- 初始化面板布局（格式转换区、另存为区、扩展区、装饰区）
- 管理按钮颜色和样式映射
- 提供格式标准化、按钮状态更新等公共方法
- 处理格式按钮点击事件

依赖：
- ttkbootstrap: 提供现代化的UI组件
- dpi_utils: 支持DPI缩放
- font_utils: 字体配置
- gui_utils: ToolTip组件

使用方式：
    此模块作为基类被 ConversionPanel 继承，不应直接实例化。
    所有 Mixin 类都依赖此基类定义的属性和方法。
"""

import logging
import tkinter as tk
from collections.abc import Callable
from typing import Any, ClassVar

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.i18n import t
from docwen.utils.dpi_utils import scale
from docwen.utils.font_utils import get_default_font, get_small_font
from docwen.utils.gui_utils import ToolTip

logger = logging.getLogger(__name__)


class ConversionPanelBase(tb.Frame):
    """
    格式转换面板基类

    定义所有子类（Mixin）共享的属性和公共方法。

    属性：
        config_manager: 配置管理器实例
        on_action: 操作回调函数
        current_category: 当前文件类别
        current_format: 当前文件格式
        current_file_path: 当前文件路径
        format_buttons: 格式按钮字典
        button_colors: 按钮颜色映射
        button_style_*: 按钮样式配置
    """

    # 各类别支持的格式定义
    FORMATS_BY_CATEGORY: ClassVar[dict[str, list[str]]] = {
        "document": ["DOCX", "DOC", "RTF", "ODT", "PDF"],
        "spreadsheet": ["XLSX", "XLS", "ET", "CSV", "ODS"],
        "image": ["PNG", "JPG", "BMP", "GIF", "TIF", "WebP"],
        "layout": ["PDF", "DOCX"],
    }

    # 支持有损压缩的图片格式
    COMPRESSIBLE_FORMATS: ClassVar[list[str]] = ["JPG", "JPEG", "WEBP"]

    def __init__(
        self,
        master,
        config_manager: Any = None,
        on_action: Callable | None = None,
        height: int | None = None,
        **kwargs,
    ):
        """
        初始化格式转换面板基类

        参数：
            master: 父组件
            config_manager: 配置管理器实例，用于读取选项默认值
            on_action: 操作回调函数 fn(action_type: str, file_path: str, options: dict)
            height: 格式转换section的目标高度
        """
        logger.debug("初始化格式转换面板基类")
        super().__init__(master, **kwargs)

        # 存储配置和回调
        self.config_manager = config_manager
        self.on_action = on_action

        # 文件状态
        self.current_category: str | None = None
        self.current_format: str | None = None
        self.current_file_path: str | None = None
        self.file_list: list[str] = []
        self.ui_mode: str = "single"

        # 按钮管理
        self.format_buttons: dict[str, tb.Button] = {}
        self.conversion_section_height = height

        # 校对/汇总选项变量
        self.checkbox_vars: dict[str, tk.BooleanVar] = {}
        self.merge_mode_var: tk.IntVar | None = None
        self.reference_table_var: tk.StringVar | None = None

        # 合并拆分相关状态（版式类）
        self.pdf_total_pages: int = 0
        self.page_input_var: tk.StringVar | None = None
        self.page_input_entry: tb.Entry | None = None
        self.split_pdf_button: tb.Button | None = None
        self.merge_pdfs_button: tb.Button | None = None
        self.pdf_info_var: tk.StringVar | None = None
        self.pdf_info_label: tb.Label | None = None
        self.pdf_info_title_label: tb.Label | None = None
        self.page_warning_label: tb.Label | None = None

        # 图片压缩相关变量
        self.compress_mode_var: tk.StringVar | None = None
        self.size_limit_var: tk.StringVar | None = None
        self.size_unit_var: tk.StringVar | None = None
        self.size_limit_entry: tb.Entry | None = None
        self.size_unit_combo: tb.Combobox | None = None
        self.size_warning_label: tb.Label | None = None
        self.pdf_quality_var: tk.StringVar | None = None
        self.convert_image_to_pdf_button: tb.Button | None = None

        # TIFF合并相关变量
        self.tiff_mode_var: tk.StringVar | None = None
        self.merge_tiff_button: tb.Button | None = None

        # 版式文件渲染相关变量
        self.layout_image_dpi_var: tk.StringVar | None = None
        self.convert_to_tif_button: tb.Button | None = None
        self.convert_to_jpg_button: tb.Button | None = None

        # 获取字体配置
        self.default_font, self.default_size = get_default_font()

        # 按钮颜色映射
        self.button_colors = {
            # 文档类
            "DOCX": "primary",
            "DOC": "info",
            "ODT": "success",
            "RTF": "warning",
            # 表格类
            "XLSX": "primary",
            "XLS": "info",
            "ET": "info",
            "ODS": "success",
            "CSV": "warning",
            # 图片类
            "PNG": "primary",
            "JPG": "primary",
            "BMP": "info",
            "GIF": "success",
            "TIF": "warning",
            "WebP": "danger",
            # 版式类
            "PDF": "danger",
            "XPS": "info",
            "OFD": "success",
            "CEB": "warning",
        }

        # 按钮样式定义（支持DPI缩放）
        self.button_style_3col = {"width": 8, "padding": (scale(8), scale(4))}
        self.button_style_2col = {"width": 10, "padding": (scale(8), scale(4))}
        self.button_style_1col = {"width": 16, "padding": (scale(8), scale(4))}

        # 创建界面元素
        self._create_widgets()

        logger.info("格式转换面板基类初始化完成")

    def _create_widgets(self):
        """
        创建界面元素

        构建面板的基本布局：
        - Row 0: 格式转换section
        - Row 1: 另存为section
        - Row 2: 扩展section（校对/汇总/合并拆分）
        - Row 3: 装饰区域
        """
        logger.debug("创建格式转换面板界面元素")

        # 配置grid布局
        self.grid_rowconfigure(0, weight=0)  # 格式转换
        self.grid_rowconfigure(1, weight=0)  # 另存为
        self.grid_rowconfigure(2, weight=0)  # 扩展区域
        self.grid_rowconfigure(3, weight=1)  # 装饰区域（占满剩余空间）
        self.grid_columnconfigure(0, weight=1)

        # 创建"格式转换"section
        self.conversion_frame = tb.Labelframe(self, text=t("conversion_panel.format_conversion"), bootstyle="success")
        self.conversion_frame.grid(row=0, column=0, sticky="nsew", pady=(0, scale(20)))
        self.conversion_frame.grid_rowconfigure(0, weight=1)
        self.conversion_frame.grid_columnconfigure(0, weight=1)

        # 格式转换按钮容器
        self.conversion_container = tb.Frame(self.conversion_frame, bootstyle="default")
        self.conversion_container.grid(row=0, column=0, sticky="nsew", padx=scale(10), pady=scale(10))

        # 创建"另存为"section
        self.saveas_frame = tb.Labelframe(self, text=t("conversion_panel.save_as"), bootstyle="danger")
        self.saveas_frame.grid(row=1, column=0, sticky="nsew", pady=(0, scale(20)))
        self.saveas_frame.grid_rowconfigure(0, weight=1)
        self.saveas_frame.grid_columnconfigure(0, weight=1)

        # 另存为按钮容器
        self.saveas_container = tb.Frame(self.saveas_frame, bootstyle="default")
        self.saveas_container.grid(row=0, column=0, sticky="nsew", padx=scale(10), pady=scale(10))

        # 创建扩展section（动态显示）
        self.extra_frame = tb.Labelframe(self, text="", bootstyle="warning")
        self.extra_frame.grid(row=2, column=0, sticky="nsew")
        self.extra_frame.grid_remove()  # 默认隐藏
        self.extra_frame.grid_rowconfigure(0, weight=1)
        self.extra_frame.grid_columnconfigure(0, weight=1)

        # 扩展section容器
        self.extra_container = tb.Frame(self.extra_frame, bootstyle="default")
        self.extra_container.grid(row=0, column=0, sticky="nsew", padx=scale(10), pady=scale(10))

        # 默认提示信息
        small_font_name, small_font_size = get_small_font()
        self.hint_label = tb.Label(
            self.conversion_container,
            text=t("conversion_panel.select_file_hint"),
            font=(small_font_name, small_font_size),
            bootstyle="secondary",
        )
        self.hint_label.pack(pady=scale(20))

        # 创建装饰区域（占满底部剩余空间）
        create_decoration_area = getattr(self, "_create_decoration_area", None)
        if callable(create_decoration_area):
            create_decoration_area()

        logger.debug("格式转换面板界面元素创建完成")

    def set_file_info(
        self,
        category: str,
        current_format: str,
        file_path: str | None = None,
        file_list: list[str] | None = None,
        ui_mode: str = "single",
    ):
        """
        设置文件信息并更新按钮显示

        参数：
            category: 文件类别 ('document', 'spreadsheet', 'image', 'layout')
            current_format: 当前文件格式 (如 'docx', 'png' 等)
            file_path: 文件路径（用于执行转换操作）
            file_list: 当前选项卡下的所有文件列表（批量模式使用）
            ui_mode: UI模式 ('single' 或 'batch')
        """
        logger.info(f"设置文件信息: 类别={category}, 格式={current_format}, 模式={ui_mode}")

        self.current_category = category
        self.current_format = current_format.lower() if current_format else None
        self.current_file_path = file_path
        self.file_list = file_list or []
        self.ui_mode = ui_mode

        # 清空现有按钮
        self._clear_buttons()
        self.hint_label.pack_forget()

        # 根据类别创建对应的格式按钮
        if category == "document":
            self._create_document_buttons()
        elif category == "spreadsheet":
            self._create_spreadsheet_buttons()
        elif category == "image":
            self._create_image_buttons()
        elif category == "layout":
            self._create_layout_buttons()
        else:
            logger.warning(f"未知的文件类别: {category}")
            self.hint_label.pack(pady=scale(20))

        # 更新按钮状态
        self._update_button_states()

        logger.info(f"格式转换面板已更新，显示 {len(self.format_buttons)} 个格式按钮")

    def _clear_buttons(self):
        """清空所有格式按钮和选项"""
        logger.debug("清空格式按钮和选项")

        # 清空各容器的子组件
        for widget in self.conversion_container.winfo_children():
            widget.destroy()
        for widget in self.saveas_container.winfo_children():
            widget.destroy()
        for widget in self.extra_container.winfo_children():
            widget.destroy()

        # 隐藏扩展section
        self.extra_frame.grid_remove()

        # 重置容器的grid列配置
        for col in range(10):
            self.conversion_container.grid_columnconfigure(col, weight=0, uniform="")
            self.saveas_container.grid_columnconfigure(col, weight=0, uniform="")
            self.extra_container.grid_columnconfigure(col, weight=0, uniform="")

        # 清空按钮和选项字典
        self.format_buttons.clear()
        self.checkbox_vars.clear()
        self.merge_mode_var = None
        self.reference_table_var = None

    def update_button_states(self):
        """根据当前格式更新按钮状态和tooltip"""
        self._update_button_states()

    def _update_button_states(self):
        """根据当前格式更新按钮状态和tooltip"""
        logger.debug(f"更新按钮状态，当前格式: {self.current_format}, UI模式: {self.ui_mode}")

        if not self.current_format:
            return

        # 检查压缩模式
        is_compress_mode = False
        if hasattr(self, "compress_mode_var") and self.compress_mode_var:
            is_compress_mode = self.compress_mode_var.get() == "limit_size"

        # 检查压缩输入有效性
        is_compress_input_valid = False
        if is_compress_mode and hasattr(self, "size_limit_var") and self.size_limit_var:
            input_text = self.size_limit_var.get()
            unit = self.size_unit_var.get() if hasattr(self, "size_unit_var") and self.size_unit_var else "KB"
            is_compress_input_valid = self._validate_size_input(input_text, unit)

        # 判断是否为批量模式
        is_batch_mode = self.ui_mode == "batch"

        # 获取所有相关格式
        all_formats = set()
        if is_batch_mode:
            from docwen.utils.file_type_utils import detect_actual_file_format

            for file_path in self.file_list:
                fmt = detect_actual_file_format(file_path)
                all_formats.add(self._normalize_format(fmt))
        else:
            all_formats.add(self._normalize_format(self.current_format))

        # 判断是否所有文件格式一致
        all_same_format = len(all_formats) == 1
        normalized_current = self._normalize_format(self.current_format)

        for fmt, button in self.format_buttons.items():
            normalized_fmt = self._normalize_format(fmt.lower())
            is_same_format = all_same_format and normalized_fmt == normalized_current

            # OFD格式保持禁用
            if fmt.upper() == "OFD":
                continue

            # 根据模式更新按钮状态
            if is_compress_mode and is_compress_input_valid:
                self._update_button_for_compress_mode(fmt, button, is_same_format)
            else:
                self._update_button_normal_mode(fmt, button, is_same_format, is_compress_mode)

    def _update_button_for_compress_mode(self, fmt: str, button: tb.Button, is_same_format: bool):
        """压缩模式下更新按钮状态"""
        if fmt.upper() in self.COMPRESSIBLE_FORMATS:
            original_color = self.button_colors.get(fmt, "primary")
            button.configure(state="normal", bootstyle=original_color)
            tooltip_text = self._get_button_tooltip(fmt, is_same_format, "normal")
            ToolTip(button, tooltip_text)
        else:
            button.configure(state="disabled", bootstyle="secondary")
            tooltip_text = self._get_button_tooltip(fmt, is_same_format, "disabled")
            ToolTip(button, tooltip_text)

    def _update_button_normal_mode(self, fmt: str, button: tb.Button, is_same_format: bool, is_compress_mode: bool):
        """非压缩模式下更新按钮状态"""
        if is_same_format:
            if is_compress_mode:
                original_color = self.button_colors.get(fmt, "primary")
                button.configure(state="normal", bootstyle=original_color)
            else:
                button.configure(state="disabled", bootstyle="secondary")
            tooltip_text = self._get_button_tooltip(
                fmt, is_same_format, "disabled" if not is_compress_mode else "normal"
            )
            ToolTip(button, tooltip_text)
        else:
            original_color = self.button_colors.get(fmt, "primary")
            button.configure(state="normal", bootstyle=original_color)
            tooltip_text = self._get_button_tooltip(fmt, is_same_format, "normal")
            ToolTip(button, tooltip_text)

    def _normalize_format(self, fmt: str) -> str:
        """
        标准化格式名称（处理等价格式）

        参数：
            fmt: 原始格式名称

        返回：
            str: 标准化后的格式名称（小写）
        """
        fmt = fmt.lower()
        equivalents = {"jpg": "jpeg", "jpeg": "jpeg", "tif": "tiff", "tiff": "tiff", "heif": "heic", "heic": "heic"}
        return equivalents.get(fmt, fmt)

    def _validate_size_input(self, value: str, unit: str) -> bool:
        """
        验证文件大小输入

        参数：
            value: 输入值
            unit: 单位 (KB或MB)

        返回：
            bool: 是否有效
        """
        try:
            size = int(value)
            if unit == "KB":
                return 1 <= size <= 10240
            else:
                return 1 <= size <= 100
        except ValueError:
            return False

    def _on_format_clicked(self, target_format: str):
        """
        格式按钮点击事件处理

        参数：
            target_format: 目标格式
        """
        logger.info(f"格式按钮被点击: {target_format}")

        if self.on_action and self.current_file_path:
            from docwen.utils.file_type_utils import detect_actual_file_format

            source_format = detect_actual_file_format(self.current_file_path)
            target_format_lower = target_format.lower()

            # 构建 action_type
            if target_format_lower == "pdf" or (
                self.current_category == "layout" and target_format_lower in ["docx", "doc", "odt", "rtf"]
            ):
                action_type = f"convert_{self.current_category}_to_{target_format_lower}"
            else:
                action_type = f"convert_{source_format}_to_{target_format_lower}"

            options: dict[str, object] = {"target_format": target_format_lower}

            # 添加图片压缩选项
            if self.current_category == "image" and hasattr(self, "compress_mode_var") and self.compress_mode_var:
                compress_mode = self.compress_mode_var.get()
                options["compress_mode"] = compress_mode

                if compress_mode == "limit_size" and hasattr(self, "size_limit_var") and self.size_limit_var:
                    try:
                        options["size_limit"] = int(self.size_limit_var.get())
                        options["size_unit"] = self.size_unit_var.get() if self.size_unit_var else "KB"
                    except (ValueError, AttributeError) as e:
                        logger.warning(f"无法获取压缩选项: {e}")

            logger.info(f"执行格式转换: {source_format.upper()} → {target_format_lower.upper()}")
            self.on_action(action_type, self.current_file_path, options)
        else:
            logger.warning("没有提供 on_action 回调或缺少文件路径")

    def on_format_clicked(self, target_format: str):
        self._on_format_clicked(target_format)

    def _get_button_tooltip(self, fmt: str, is_same_format: bool, button_state: str) -> str:
        """
        根据当前状态获取按钮的tooltip文本

        参数：
            fmt: 格式名称（大写）
            is_same_format: 是否与当前格式相同
            button_state: 按钮状态 ('normal' 或 'disabled')

        返回：
            str: tooltip文本
        """
        fmt_upper = fmt.upper()

        if is_same_format and button_state == "disabled":
            return t("conversion_panel.same_format_tooltip", format=fmt_upper)

        if self.current_category == "image":
            return self._get_image_button_tooltip(fmt_upper, is_same_format, button_state)
        elif self.current_category == "document":
            return self._get_document_button_tooltip(fmt_upper)
        elif self.current_category == "spreadsheet":
            return self._get_spreadsheet_button_tooltip(fmt_upper)
        elif self.current_category == "layout":
            return self._get_layout_button_tooltip(fmt_upper)
        return ""

    def _get_image_button_tooltip(self, fmt_upper: str, is_same_format: bool, button_state: str) -> str:
        """获取图片格式按钮的tooltip"""
        is_compress_mode = (
            hasattr(self, "compress_mode_var")
            and self.compress_mode_var
            and self.compress_mode_var.get() == "limit_size"
        )

        if is_compress_mode:
            if fmt_upper in self.COMPRESSIBLE_FORMATS:
                return t("conversion_panel.compressible_format_tooltip")
            else:
                format_tips = {
                    "PNG": t("conversion_panel.png_compression_tooltip"),
                    "BMP": t("conversion_panel.bmp_compression_tooltip"),
                    "GIF": t("conversion_panel.gif_compression_tooltip"),
                    "TIF": t("conversion_panel.tif_compression_tooltip"),
                    "TIFF": t("conversion_panel.tif_compression_tooltip"),
                }
                return format_tips.get(
                    fmt_upper, t("conversion_panel.other_format_compression_tooltip", format=fmt_upper)
                )
        else:
            if fmt_upper in ["JPG", "JPEG", "WEBP"]:
                return t("conversion_panel.high_quality_conversion_tooltip")
            return t("conversion_panel.lossless_conversion_tooltip")

    def _get_document_button_tooltip(self, fmt_upper: str) -> str:
        """获取文档格式按钮的tooltip"""
        return t("conversion_panel.document_conversion_tooltip")

    def _get_spreadsheet_button_tooltip(self, fmt_upper: str) -> str:
        """获取表格格式按钮的tooltip"""
        if fmt_upper == "CSV":
            return t("conversion_panel.csv_conversion_tooltip")
        return t("conversion_panel.spreadsheet_conversion_tooltip")

    def _get_layout_button_tooltip(self, fmt_upper: str) -> str:
        """获取版式格式按钮的tooltip"""
        tooltips = {
            "PDF": t("conversion_panel.pdf_conversion_tooltip")
            if self.current_format and self.current_format.lower() in ["ofd", "xps"]
            else t("conversion_panel.same_format_tooltip", format="PDF"),
            "DOCX": t("conversion_panel.docx_conversion_tooltip"),
            "DOC": t("conversion_panel.doc_conversion_tooltip"),
            "TIF": t("conversion_panel.tif_conversion_tooltip"),
            "JPG": t("conversion_panel.jpg_conversion_tooltip"),
        }
        return tooltips.get(fmt_upper, "")

    # 以下方法由Mixin类实现
    def _create_document_buttons(self):
        """创建文档类格式按钮（由DocumentSectionMixin实现）"""
        raise NotImplementedError("由DocumentSectionMixin实现")

    def _create_spreadsheet_buttons(self):
        """创建表格类格式按钮（由SpreadsheetSectionMixin实现）"""
        raise NotImplementedError("由SpreadsheetSectionMixin实现")

    def _create_image_buttons(self):
        """创建图片类格式按钮（由ImageSectionMixin实现）"""
        raise NotImplementedError("由ImageSectionMixin实现")

    def _create_layout_buttons(self):
        """创建版式类格式按钮（由LayoutSectionMixin实现）"""
        raise NotImplementedError("由LayoutSectionMixin实现")

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
        self.hint_label.pack(pady=scale(20))

        logger.info("格式转换面板已重置")
