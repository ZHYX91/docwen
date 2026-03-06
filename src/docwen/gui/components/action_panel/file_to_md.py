"""
文件转MD功能模块

提供各类文件转换为 Markdown 的按钮和选项：
- 文档转MD：DOCX/DOC等文档文件
- 表格转MD：XLSX/XLS等表格文件
- 图片转MD：PNG/JPG等图片文件
- 版式转MD：PDF/OFD等版式文件

每种类型都支持导出选项（提取图片、OCR识别），
使用公共组件 ExportOptionHandler 处理选项联动逻辑。

依赖：
- ActionPanelBase: 提供按钮样式、颜色映射等公共属性
- config_manager: 读取选项默认值
- ExportOptionHandler: 处理提取图片和OCR选项的联动

使用方式：
    此模块作为 Mixin 类被 ActionPanel 继承，不应直接实例化。
"""

import logging
import tkinter as tk
from typing import TYPE_CHECKING

import ttkbootstrap as tb

from docwen.gui.components.common import ExportOptionHandler
from docwen.gui.core.mixins_protocols import ActionPanelHost
from docwen.i18n import t
from docwen.utils.dpi_utils import scale
from docwen.utils.gui_utils import ToolTip, create_info_icon

logger = logging.getLogger(__name__)

if TYPE_CHECKING:
    from .base import ActionPanelBase as _ActionPanelBase
else:
    _ActionPanelBase = object


class FileToMdMixin(_ActionPanelBase):
    """
    文件转MD功能混入类

    提供各类文件转换为 Markdown 的所有功能：
    - setup_for_document_file: 设置文档转MD模式
    - setup_for_spreadsheet_file: 设置表格转MD模式
    - setup_for_image_file: 设置图片转MD模式
    - setup_for_layout_file: 设置版式转MD模式

    每种模式都有对应的按钮创建方法和点击事件处理方法。

    依赖基类属性：
        button_container: 按钮容器框架
        button_colors: 按钮颜色映射
        button_style_1: 按钮样式配置
        config_manager: 配置管理器
        on_action: 操作回调函数
    """

    # ==================== 文档转MD ====================

    def setup_for_document_file(self: ActionPanelHost, file_path: str, file_list: list[str] | None = None):
        """
        设置为文档文件处理模式

        为DOCX/DOC等文档文件显示：
        - 导出Markdown按钮
        - 导出选项（提取图片、图片文字识别）
        - 优化选项和序号选项

        参数：
            file_path: 文档文件路径
            file_list: 文件列表（批量模式时用于更新按钮状态）
        """
        logger.debug(f"设置文档文件处理模式: {file_path}")
        self.file_type = "document"
        self.file_path = file_path
        self.clear_buttons()
        self.clear_options()
        self._create_document_conversion_buttons()
        self.status_var.set(t("action_panel.document.ready"))
        logger.info("文档文件操作面板设置完成")

    def _create_document_conversion_buttons(self: ActionPanelHost):
        """
        创建文档文件转换按钮

        显示：
        - 第一行：导出Markdown按钮
        - 导出选项边框（提取图片、图片文字识别）
        - 优化选项和序号选项
        """
        logger.debug("创建文档转换按钮")

        # 第一行：导出Markdown按钮
        self.convert_document_to_md_button = tb.Button(
            self.button_container,
            text=t("action_panel.document.export_markdown"),
            command=self._on_convert_document_to_md_clicked,
            bootstyle=self.button_colors["success"],
            **self.button_style_1,
        )
        self.convert_document_to_md_button.grid(row=0, column=0, pady=(0, scale(10)))
        ToolTip(self.convert_document_to_md_button, t("action_panel.document.export_markdown_tooltip"))

        # 导出选项边框
        doc_export_options_frame = tb.Labelframe(
            self.button_container, text=t("action_panel.export_options"), bootstyle="info"
        )
        doc_export_options_frame.grid(row=1, column=0, sticky="ew", padx=scale(20), pady=scale(10))

        # 配置选项框架网格权重
        doc_export_options_frame.grid_rowconfigure(0, weight=1)
        doc_export_options_frame.grid_rowconfigure(1, weight=1)
        doc_export_options_frame.grid_columnconfigure(0, weight=1)

        # 从配置读取默认值
        try:
            default_extract_image = self.config_manager.get_docx_to_md_keep_images()
            default_extract_ocr = self.config_manager.get_docx_to_md_enable_ocr()
        except Exception as e:
            logger.warning(f"读取文档转MD配置失败，使用默认值: {e}")
            default_extract_image = True
            default_extract_ocr = False

        # 创建选项变量
        self.doc_extract_image_var = tk.BooleanVar(value=default_extract_image)
        self.doc_extract_ocr_var = tk.BooleanVar(value=default_extract_ocr)

        # 创建导出选项处理器
        self._doc_export_handler = ExportOptionHandler(self.doc_extract_image_var, self.doc_extract_ocr_var)

        # 多选框容器
        checkbox_container = tb.Frame(doc_export_options_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))

        # 配置行列权重
        for i in range(5):
            checkbox_container.grid_rowconfigure(i, weight=0)
        checkbox_container.grid_columnconfigure(0, weight=1)
        checkbox_container.grid_columnconfigure(1, weight=1)

        # 从配置读取序号默认值
        try:
            default_remove_numbering = self.config_manager.get_docx_to_md_remove_numbering()
            default_add_numbering = self.config_manager.get_docx_to_md_add_numbering()
            default_scheme_id = self.config_manager.get_docx_to_md_default_scheme()
        except Exception as e:
            logger.warning(f"读取文档转MD序号配置失败，使用默认值: {e}")
            default_remove_numbering = True
            default_add_numbering = False
            default_scheme_id = "hierarchical_standard"

        # 获取适用于当前语言的序号方案（从配置文件获取）
        self._doc_numbering_schemes = self.config_manager.get_localized_numbering_schemes()
        scheme_names = list(self._doc_numbering_schemes.values())
        scheme_id_to_name = self._doc_numbering_schemes

        # 确定默认序号方案名称（如果配置的方案不适用于当前语言，使用第一个可用方案）
        if default_scheme_id in scheme_id_to_name:
            default_scheme_name = scheme_id_to_name[default_scheme_id]
        elif scheme_names:
            default_scheme_name = scheme_names[0]
            default_scheme_id = next(iter(self._doc_numbering_schemes.keys()))
            logger.debug(f"配置的序号方案不适用于当前语言，使用: {default_scheme_id}")
        else:
            default_scheme_name = ""
            logger.warning("当前语言没有可用的序号方案")

        self._default_doc_numbering_scheme_id = default_scheme_id

        # 从配置读取优化默认值
        try:
            default_enable_optimization = self.config_manager.get_docx_to_md_enable_optimization()
            default_optimization_type_id = self.config_manager.get_docx_to_md_optimization_type()
        except Exception as e:
            logger.warning(f"读取优化配置失败，使用默认值: {e}")
            default_enable_optimization = True
            default_optimization_type_id = "gongwen"

        # 获取适用于当前语言的优化类型（从配置文件获取）
        self._doc_optimization_types = self.config_manager.get_localized_optimization_types(scope="document_to_md")
        optimization_type_names = list(self._doc_optimization_types.values())
        optimization_type_id_to_name = self._doc_optimization_types

        # 检查当前语言是否有可用的优化类型
        self._has_optimization_types = len(self._doc_optimization_types) > 0

        # 确定默认优化类型名称
        if default_optimization_type_id in optimization_type_id_to_name:
            default_optimization_type = optimization_type_id_to_name[default_optimization_type_id]
        elif optimization_type_names:
            default_optimization_type = optimization_type_names[0]
            default_optimization_type_id = next(iter(self._doc_optimization_types.keys()))
        else:
            default_optimization_type = ""
            # 如果没有可用的优化类型，禁用优化功能
            default_enable_optimization = False

        self._default_doc_optimization_type_id = default_optimization_type_id

        # ===== 第1行：图片选项 =====
        # 提取图片 + 信息图标
        extract_image_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_image_container.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(20)))

        self.doc_extract_image_check = tb.Checkbutton(
            extract_image_container,
            text=t("action_panel.extract_images"),
            variable=self.doc_extract_image_var,
            command=self._doc_export_handler.on_option_changed,
            bootstyle="round-toggle",
        )
        self.doc_extract_image_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        extract_image_info = create_info_icon(
            extract_image_container, t("action_panel.extract_images_tooltip"), bootstyle="info"
        )
        extract_image_info.pack(side=tk.LEFT)

        # 图片文字识别 + 信息图标
        extract_ocr_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_ocr_container.grid(row=0, column=1, sticky="w", padx=(0, scale(10)))

        self.doc_extract_ocr_check = tb.Checkbutton(
            extract_ocr_container,
            text=t("action_panel.ocr"),
            variable=self.doc_extract_ocr_var,
            command=self._doc_export_handler.on_option_changed,
            bootstyle="round-toggle",
        )
        self.doc_extract_ocr_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        extract_ocr_info = create_info_icon(extract_ocr_container, t("action_panel.ocr_tooltip"), bootstyle="info")
        extract_ocr_info.pack(side=tk.LEFT)

        # ===== 第2行：分割线 =====
        separator = tb.Separator(checkbox_container, bootstyle="info")
        separator.grid(row=1, column=0, columnspan=2, sticky="ew", padx=0, pady=scale(20))

        # ===== 第3行：针对文档类型优化（仅在当前语言有可用优化类型时显示） =====
        # 保存优化容器引用，用于后续可能的隐藏/显示
        self._optimization_container = None

        if self._has_optimization_types:
            optimization_container = tb.Frame(checkbox_container, bootstyle="default")
            optimization_container.grid(
                row=2, column=0, columnspan=2, sticky="w", padx=(scale(10), scale(10)), pady=(0, scale(5))
            )
            self._optimization_container = optimization_container

            self.doc_enable_optimization_var = tk.BooleanVar(value=default_enable_optimization)
            self.doc_enable_optimization_check = tb.Checkbutton(
                optimization_container,
                text=t("action_panel.document.optimize_for_type"),
                variable=self.doc_enable_optimization_var,
                command=self._on_doc_optimization_toggle,
                bootstyle="round-toggle",
            )
            self.doc_enable_optimization_check.pack(side=tk.LEFT, padx=(0, scale(5)))

            optimization_check_info = create_info_icon(
                optimization_container, t("action_panel.document.optimize_for_type_tooltip"), bootstyle="info"
            )
            optimization_check_info.pack(side=tk.LEFT, padx=(0, scale(10)))

            self.doc_optimization_type_var = tk.StringVar(value=default_optimization_type)
            self.doc_optimization_type_combo = tb.Combobox(
                optimization_container,
                textvariable=self.doc_optimization_type_var,
                values=optimization_type_names,
                state="readonly" if default_enable_optimization else "disabled",
                width=10,
            )
            self.doc_optimization_type_combo.pack(side=tk.LEFT)
            logger.debug(f"优化类型选项已创建，可用类型: {optimization_type_names}")
        else:
            # 当前语言没有可用的优化类型，设置默认值但不显示UI
            self.doc_enable_optimization_var = tk.BooleanVar(value=False)
            self.doc_optimization_type_var = tk.StringVar(value="")
            logger.info("当前语言没有可用的优化类型，优化选项区域已隐藏")

        # ===== 第4行：清除原有文档小标题序号 =====
        self.doc_remove_numbering_var = tk.BooleanVar(value=default_remove_numbering)
        remove_numbering_container = tb.Frame(checkbox_container, bootstyle="default")
        remove_numbering_container.grid(
            row=3, column=0, columnspan=2, sticky="w", padx=(scale(10), scale(10)), pady=(0, scale(5))
        )

        doc_remove_numbering_check = tb.Checkbutton(
            remove_numbering_container,
            text=t("action_panel.document.remove_existing_numbering"),
            variable=self.doc_remove_numbering_var,
            bootstyle="round-toggle",
        )
        doc_remove_numbering_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        remove_numbering_info = create_info_icon(
            remove_numbering_container, t("action_panel.document.remove_existing_numbering_tooltip"), bootstyle="info"
        )
        remove_numbering_info.pack(side=tk.LEFT)

        # ===== 第5行：新增小标题序号 + 下拉框 =====
        self.doc_add_numbering_var = tk.BooleanVar(value=default_add_numbering)
        add_numbering_container = tb.Frame(checkbox_container, bootstyle="default")
        add_numbering_container.grid(
            row=4, column=0, columnspan=2, sticky="w", padx=(scale(10), scale(10)), pady=(0, scale(5))
        )

        doc_add_numbering_check = tb.Checkbutton(
            add_numbering_container,
            text=t("action_panel.document.add_new_numbering"),
            variable=self.doc_add_numbering_var,
            command=self._on_doc_add_numbering_toggle,
            bootstyle="round-toggle",
        )
        doc_add_numbering_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        add_numbering_info = create_info_icon(
            add_numbering_container, t("action_panel.document.add_new_numbering_tooltip"), bootstyle="info"
        )
        add_numbering_info.pack(side=tk.LEFT, padx=(0, scale(10)))

        self.doc_numbering_scheme_var = tk.StringVar(value=default_scheme_name)
        self.doc_numbering_scheme_combo = tb.Combobox(
            add_numbering_container,
            textvariable=self.doc_numbering_scheme_var,
            values=scheme_names,
            state="disabled" if not default_add_numbering else "readonly",
            width=15,
        )
        self.doc_numbering_scheme_combo.pack(side=tk.LEFT)

        logger.debug("文档文件转换按钮创建完成（含导出选项、优化选项和序号选项）")

    def _on_doc_optimization_toggle(self: ActionPanelHost):
        """处理针对优化复选框切换事件"""
        if self.doc_enable_optimization_var.get():
            self.doc_optimization_type_combo.config(state="readonly")
            logger.debug("针对优化已启用，下拉框可选")
        else:
            self.doc_optimization_type_combo.config(state="disabled")
            logger.debug("针对优化已禁用，下拉框灰色")

    def _on_doc_add_numbering_toggle(self: ActionPanelHost):
        """处理文档转MD"添加标题序号"复选框切换事件"""
        if hasattr(self, "doc_add_numbering_var") and hasattr(self, "doc_numbering_scheme_combo"):
            if self.doc_add_numbering_var.get():
                self.doc_numbering_scheme_combo.config(state="readonly")
                logger.debug("文档转MD：添加序号已启用，序号方案下拉框可选")
            else:
                self.doc_numbering_scheme_combo.config(state="disabled")
                logger.debug("文档转MD：添加序号已禁用，序号方案下拉框灰色")

    def _on_convert_document_to_md_clicked(self: ActionPanelHost):
        """处理文档转Markdown按钮点击事件"""
        if self.on_action:
            # 获取导出选项
            extract_image = self.doc_extract_image_var.get() if hasattr(self, "doc_extract_image_var") else True
            extract_ocr = self.doc_extract_ocr_var.get() if hasattr(self, "doc_extract_ocr_var") else False

            # 获取优化选项
            enable_optimization = (
                self.doc_enable_optimization_var.get() if hasattr(self, "doc_enable_optimization_var") else False
            )
            optimization_type = (
                self.doc_optimization_type_var.get() if hasattr(self, "doc_optimization_type_var") else ""
            )

            # 确定优化类型参数（使用翻译后的名称进行反向映射）
            if enable_optimization:
                optimization_types = getattr(self, "_doc_optimization_types", {}) or {}
                name_to_id = {name: type_id for type_id, name in optimization_types.items()}
                optimize_for_type = name_to_id.get(
                    optimization_type, getattr(self, "_default_doc_optimization_type_id", "gongwen")
                )
            else:
                optimize_for_type = None

            # 构建选项字典
            options = {
                "extract_image": extract_image,
                "extract_ocr": extract_ocr,
                "optimize_for_type": optimize_for_type,
            }

            # 添加序号配置参数（使用翻译后的名称进行反向映射）
            if (
                hasattr(self, "doc_remove_numbering_var")
                and self.doc_remove_numbering_var is not None
                and self.doc_add_numbering_var is not None
            ):
                options["remove_numbering"] = self.doc_remove_numbering_var.get()
                options["add_numbering"] = self.doc_add_numbering_var.get()
                if hasattr(self, "doc_numbering_scheme_var") and self.doc_numbering_scheme_var is not None:
                    scheme_name = self.doc_numbering_scheme_var.get()
                    numbering_schemes = getattr(self, "_doc_numbering_schemes", {}) or {}
                    name_to_id = {name: scheme_id for scheme_id, name in numbering_schemes.items()}
                    options["numbering_scheme"] = name_to_id.get(
                        scheme_name, getattr(self, "_default_doc_numbering_scheme_id", "gongwen_standard")
                    )

            logger.info(
                f"文档转Markdown - 导出选项: 提取图片={extract_image}, OCR={extract_ocr}, 优化类型={optimize_for_type}"
            )
            self.on_action("convert_document_to_md", self.file_path, options)

    # ==================== 表格转MD ====================

    def setup_for_spreadsheet_file(self: ActionPanelHost, file_path: str, file_list: list[str] | None = None):
        """
        设置为表格文件处理模式

        为XLSX/XLS/CSV等表格文件显示：
        - 导出Markdown按钮
        - 导出选项（提取图片、图片文字识别）

        参数：
            file_path: 表格文件路径
            file_list: 文件列表（批量模式时用于更新按钮状态）
        """
        logger.debug(f"设置表格文件处理模式: {file_path}")
        self.file_type = "spreadsheet"
        self.file_path = file_path
        self.clear_buttons()
        self.clear_options()
        self._create_spreadsheet_to_md_button()
        self.status_var.set(t("action_panel.spreadsheet.ready"))
        logger.info("表格文件操作面板设置完成")

    def _create_spreadsheet_to_md_button(self: ActionPanelHost):
        """创建表格文件转换按钮"""
        logger.debug("创建表格转换按钮（含导出选项）")

        # 导出Markdown按钮
        self.convert_spreadsheet_to_md_button = tb.Button(
            self.button_container,
            text=t("action_panel.spreadsheet.export_markdown"),
            command=self._on_convert_spreadsheet_to_md_clicked,
            bootstyle=self.button_colors["success"],
            **self.button_style_1,
        )
        self.convert_spreadsheet_to_md_button.grid(row=0, column=0, pady=(0, scale(10)))
        ToolTip(self.convert_spreadsheet_to_md_button, t("action_panel.spreadsheet.export_markdown_tooltip"))

        # 导出选项边框
        table_export_options_frame = tb.Labelframe(
            self.button_container, text=t("action_panel.export_options"), bootstyle="info"
        )
        table_export_options_frame.grid(row=1, column=0, sticky="ew", padx=scale(20), pady=scale(10))

        table_export_options_frame.grid_rowconfigure(0, weight=1)
        table_export_options_frame.grid_columnconfigure(0, weight=1)

        # 从配置读取默认值
        try:
            default_extract_image = self.config_manager.get_xlsx_to_md_keep_images()
            default_extract_ocr = self.config_manager.get_xlsx_to_md_enable_ocr()
        except Exception as e:
            logger.warning(f"读取表格转MD配置失败，使用默认值: {e}")
            default_extract_image = True
            default_extract_ocr = False

        # 创建选项变量
        self.table_extract_image_var = tk.BooleanVar(value=default_extract_image)
        self.table_extract_ocr_var = tk.BooleanVar(value=default_extract_ocr)

        # 创建导出选项处理器
        self._table_export_handler = ExportOptionHandler(self.table_extract_image_var, self.table_extract_ocr_var)

        # 多选框容器
        checkbox_container = tb.Frame(table_export_options_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        checkbox_container.grid_columnconfigure(0, weight=0)
        checkbox_container.grid_columnconfigure(1, weight=0)

        # 提取图片 + 信息图标
        extract_image_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_image_container.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(20)))

        self.table_extract_image_check = tb.Checkbutton(
            extract_image_container,
            text=t("action_panel.extract_images"),
            variable=self.table_extract_image_var,
            command=self._table_export_handler.on_option_changed,
            bootstyle="round-toggle",
        )
        self.table_extract_image_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        extract_image_info = create_info_icon(
            extract_image_container, t("action_panel.extract_images_tooltip"), bootstyle="info"
        )
        extract_image_info.pack(side=tk.LEFT)

        # 图片文字识别 + 信息图标
        extract_ocr_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_ocr_container.grid(row=0, column=1, sticky="w", padx=(0, scale(10)))

        self.table_extract_ocr_check = tb.Checkbutton(
            extract_ocr_container,
            text=t("action_panel.ocr"),
            variable=self.table_extract_ocr_var,
            command=self._table_export_handler.on_option_changed,
            bootstyle="round-toggle",
        )
        self.table_extract_ocr_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        extract_ocr_info = create_info_icon(extract_ocr_container, t("action_panel.ocr_tooltip"), bootstyle="info")
        extract_ocr_info.pack(side=tk.LEFT)

        try:
            self._table_optimization_types = self.config_manager.get_localized_optimization_types(scope="spreadsheet_to_md")
        except Exception:
            self._table_optimization_types = {}

        table_optimization_type_names = list(self._table_optimization_types.values())
        self._has_table_optimization_types = len(self._table_optimization_types) > 0

        if self._has_table_optimization_types:
            optimization_container = tb.Frame(checkbox_container, bootstyle="default")
            optimization_container.grid(
                row=1,
                column=0,
                columnspan=2,
                sticky="w",
                padx=(scale(10), scale(10)),
                pady=(scale(5), 0),
            )

            self.table_enable_optimization_var = tk.BooleanVar(value=False)
            self.table_enable_optimization_check = tb.Checkbutton(
                optimization_container,
                text=t("action_panel.spreadsheet.optimize_for_type"),
                variable=self.table_enable_optimization_var,
                command=self._on_table_optimization_toggle,
                bootstyle="round-toggle",
            )
            self.table_enable_optimization_check.pack(side=tk.LEFT, padx=(0, scale(5)))

            optimization_check_info = create_info_icon(
                optimization_container, t("action_panel.spreadsheet.optimize_for_type_tooltip"), bootstyle="info"
            )
            optimization_check_info.pack(side=tk.LEFT, padx=(0, scale(10)))

            default_name = table_optimization_type_names[0] if table_optimization_type_names else ""
            self.table_optimization_type_var = tk.StringVar(value=default_name)
            self.table_optimization_type_combo = tb.Combobox(
                optimization_container,
                textvariable=self.table_optimization_type_var,
                values=table_optimization_type_names,
                state="disabled",
                width=scale(14),
            )
            self.table_optimization_type_combo.pack(side=tk.LEFT)
        else:
            self.table_enable_optimization_var = tk.BooleanVar(value=False)
            self.table_optimization_type_var = tk.StringVar(value="")

        logger.debug("表格文件转换按钮创建完成")

    def _on_table_optimization_toggle(self: ActionPanelHost):
        if not getattr(self, "_has_table_optimization_types", False):
            return
        if getattr(self, "table_enable_optimization_var", None) is None:
            return
        enabled = self.table_enable_optimization_var.get()
        if hasattr(self, "table_optimization_type_combo") and self.table_optimization_type_combo:
            self.table_optimization_type_combo.config(state="readonly" if enabled else "disabled")

    def _on_convert_spreadsheet_to_md_clicked(self: ActionPanelHost):
        """处理表格转Markdown按钮点击事件"""
        if self.on_action:
            extract_image = self.table_extract_image_var.get() if hasattr(self, "table_extract_image_var") else False
            extract_ocr = self.table_extract_ocr_var.get() if hasattr(self, "table_extract_ocr_var") else False

            options = {"extract_image": extract_image, "extract_ocr": extract_ocr}

            enable_optimization = (
                self.table_enable_optimization_var.get() if hasattr(self, "table_enable_optimization_var") else False
            )
            optimization_type = (
                self.table_optimization_type_var.get() if hasattr(self, "table_optimization_type_var") else ""
            )
            if enable_optimization:
                optimization_types = getattr(self, "_table_optimization_types", {}) or {}
                name_to_id = {name: type_id for type_id, name in optimization_types.items()}
                type_id = name_to_id.get(optimization_type, "")
                if type_id:
                    options["optimize_for_type"] = type_id

            logger.info(f"表格转Markdown - 导出选项: 提取图片={extract_image}, OCR={extract_ocr}")
            self.on_action("convert_spreadsheet_to_md", self.file_path, options)

    # ==================== 图片转MD ====================

    def setup_for_image_file(self: ActionPanelHost, file_path: str):
        """
        设置为图片文件处理模式

        为图片文件显示：
        - 导出Markdown按钮（OCR识别）
        - 导出选项（提取图片、图片文字识别）

        参数：
            file_path: 图片文件路径
        """
        logger.debug(f"设置图片文件处理模式: {file_path}")
        self.file_type = "image"
        self.file_path = file_path
        self.clear_buttons()
        self.clear_options()
        self._create_image_conversion_buttons()
        self.status_var.set(t("action_panel.image.ready"))
        logger.info("图片文件操作面板设置完成")

    def _create_image_conversion_buttons(self: ActionPanelHost):
        """创建图片文件转换按钮"""
        logger.debug("创建图片转换按钮")

        # 导出Markdown按钮
        self.convert_image_to_md_button = tb.Button(
            self.button_container,
            text=t("action_panel.image.export_markdown"),
            command=self._on_convert_image_to_md_clicked,
            bootstyle=self.button_colors["success"],
            **self.button_style_1,
        )
        self.convert_image_to_md_button.grid(row=0, column=0, pady=(0, scale(10)))
        ToolTip(self.convert_image_to_md_button, t("action_panel.image.export_markdown_tooltip"))

        # 导出选项边框
        export_options_frame = tb.Labelframe(
            self.button_container, text=t("action_panel.export_options"), bootstyle="info"
        )
        export_options_frame.grid(row=1, column=0, sticky="ew", padx=scale(20), pady=scale(10))

        export_options_frame.grid_rowconfigure(0, weight=1)
        export_options_frame.grid_columnconfigure(0, weight=1)

        # 从配置读取默认值
        try:
            default_extract_image = self.config_manager.get_image_to_md_keep_images()
            default_extract_ocr = self.config_manager.get_image_to_md_enable_ocr()
        except Exception as e:
            logger.warning(f"读取图片转MD配置失败，使用默认值: {e}")
            default_extract_image = True
            default_extract_ocr = False

        # 创建选项变量
        self.image_extract_image_var = tk.BooleanVar(value=default_extract_image)
        self.image_extract_ocr_var = tk.BooleanVar(value=default_extract_ocr)

        # 创建导出选项处理器，带状态变化回调
        self._image_export_handler = ExportOptionHandler(
            self.image_extract_image_var,
            self.image_extract_ocr_var,
            on_state_changed=self._on_image_export_state_changed,
        )

        # 多选框容器
        checkbox_container = tb.Frame(export_options_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        checkbox_container.grid_columnconfigure(0, weight=0)
        checkbox_container.grid_columnconfigure(1, weight=0)

        # 提取图片 + 信息图标
        extract_image_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_image_container.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(20)))

        self.extract_image_check = tb.Checkbutton(
            extract_image_container,
            text=t("action_panel.extract_images"),
            variable=self.image_extract_image_var,
            command=self._image_export_handler.on_option_changed,
            bootstyle="round-toggle",
        )
        self.extract_image_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        extract_image_info = create_info_icon(
            extract_image_container, t("action_panel.image.extract_images_tooltip"), bootstyle="info"
        )
        extract_image_info.pack(side=tk.LEFT)

        # 图片文字识别 + 信息图标
        extract_ocr_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_ocr_container.grid(row=0, column=1, sticky="w", padx=(0, scale(10)))

        self.extract_image_ocr_check = tb.Checkbutton(
            extract_ocr_container,
            text=t("action_panel.ocr"),
            variable=self.image_extract_ocr_var,
            command=self._image_export_handler.on_option_changed,
            bootstyle="round-toggle",
        )
        self.extract_image_ocr_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        extract_ocr_info = create_info_icon(
            extract_ocr_container, t("action_panel.image.ocr_tooltip"), bootstyle="info"
        )
        extract_ocr_info.pack(side=tk.LEFT)

        try:
            self._image_optimization_types = self.config_manager.get_localized_optimization_types(scope="image_to_md")
        except Exception:
            self._image_optimization_types = {}

        image_optimization_type_names = list(self._image_optimization_types.values())
        self._has_image_optimization_types = len(self._image_optimization_types) > 0

        if self._has_image_optimization_types:
            optimization_container = tb.Frame(checkbox_container, bootstyle="default")
            optimization_container.grid(
                row=1,
                column=0,
                columnspan=2,
                sticky="w",
                padx=(scale(10), scale(10)),
                pady=(scale(5), 0),
            )

            self.image_enable_optimization_var = tk.BooleanVar(value=False)
            self.image_enable_optimization_check = tb.Checkbutton(
                optimization_container,
                text=t("action_panel.image.optimize_for_type"),
                variable=self.image_enable_optimization_var,
                command=self._on_image_optimization_toggle,
                bootstyle="round-toggle",
            )
            self.image_enable_optimization_check.pack(side=tk.LEFT, padx=(0, scale(5)))

            optimization_check_info = create_info_icon(
                optimization_container, t("action_panel.image.optimize_for_type_tooltip"), bootstyle="info"
            )
            optimization_check_info.pack(side=tk.LEFT, padx=(0, scale(10)))

            default_name = image_optimization_type_names[0] if image_optimization_type_names else ""
            self.image_optimization_type_var = tk.StringVar(value=default_name)
            self.image_optimization_type_combo = tb.Combobox(
                optimization_container,
                textvariable=self.image_optimization_type_var,
                values=image_optimization_type_names,
                state="disabled",
                width=scale(14),
            )
            self.image_optimization_type_combo.pack(side=tk.LEFT)
            logger.debug(f"图片优化类型选项已创建，可用类型: {image_optimization_type_names}")
        else:
            self.image_enable_optimization_var = tk.BooleanVar(value=False)
            self.image_optimization_type_var = tk.StringVar(value="")

        logger.debug("图片转换按钮创建完成")

    def _on_image_optimization_toggle(self: ActionPanelHost):
        if not getattr(self, "_has_image_optimization_types", False):
            return
        if getattr(self, "image_enable_optimization_var", None) is None:
            return
        enabled = self.image_enable_optimization_var.get()
        if hasattr(self, "image_optimization_type_combo") and self.image_optimization_type_combo:
            self.image_optimization_type_combo.config(state="readonly" if enabled else "disabled")

    def _on_image_export_state_changed(self: ActionPanelHost):
        """处理图片导出选项状态变化"""
        # 至少勾选一个选项才启用按钮
        should_enable = self._image_export_handler.is_any_selected()
        if hasattr(self, "convert_image_to_md_button") and self.convert_image_to_md_button:
            self.convert_image_to_md_button.config(state="normal" if should_enable else "disabled")

    def _on_convert_image_to_md_clicked(self: ActionPanelHost):
        """处理图片转Markdown按钮点击事件"""
        if self.on_action:
            extract_image = self.image_extract_image_var.get() if hasattr(self, "image_extract_image_var") else True
            extract_ocr = self.image_extract_ocr_var.get() if hasattr(self, "image_extract_ocr_var") else True

            options = {"extract_image": extract_image, "extract_ocr": extract_ocr}

            enable_optimization = (
                self.image_enable_optimization_var.get() if hasattr(self, "image_enable_optimization_var") else False
            )
            optimization_type = (
                self.image_optimization_type_var.get() if hasattr(self, "image_optimization_type_var") else ""
            )
            if enable_optimization:
                optimization_types = getattr(self, "_image_optimization_types", {}) or {}
                name_to_id = {name: type_id for type_id, name in optimization_types.items()}
                type_id = name_to_id.get(optimization_type, "")
                if type_id:
                    options["optimize_for_type"] = type_id

            logger.info(f"图片转Markdown - 导出选项: 提取图片={extract_image}, OCR={extract_ocr}")
            self.on_action("convert_image_to_md", self.file_path, options)

    # ==================== 版式转MD ====================

    def setup_for_layout_file(self: ActionPanelHost, file_path: str):
        """
        设置为版式文件处理模式

        为PDF/OFD等版式文件显示导出Markdown按钮和提取选项。

        参数：
            file_path: 版式文件路径
        """
        logger.debug(f"设置版式文件处理模式: {file_path}")
        self.file_type = "layout"
        self.file_path = file_path
        self.clear_buttons()
        self.clear_options()
        self._create_layout_conversion_buttons()
        self.status_var.set(t("action_panel.layout.ready"))
        logger.info("版式文件操作面板设置完成")

    def _create_layout_conversion_buttons(self: ActionPanelHost):
        """创建PDF文件转换按钮"""
        logger.debug("创建PDF转换按钮")

        # 导出Markdown按钮
        self.convert_layout_to_md_button = tb.Button(
            self.button_container,
            text=t("action_panel.layout.export_markdown"),
            command=self._on_convert_layout_to_md_clicked,
            bootstyle=self.button_colors["success"],
            **self.button_style_1,
        )
        self.convert_layout_to_md_button.grid(row=0, column=0, pady=(0, scale(10)))
        ToolTip(self.convert_layout_to_md_button, t("action_panel.layout.export_markdown_tooltip"))

        # 提取选项边框
        extraction_frame = tb.Labelframe(
            self.button_container, text=t("action_panel.extraction_options"), bootstyle="info"
        )
        extraction_frame.grid(row=1, column=0, sticky="ew", padx=scale(20), pady=scale(10))

        extraction_frame.grid_rowconfigure(0, weight=1)
        extraction_frame.grid_columnconfigure(0, weight=1)

        # 从配置读取默认值
        try:
            default_extract_images = self.config_manager.get_layout_to_md_keep_images()
            default_extract_ocr = self.config_manager.get_layout_to_md_enable_ocr()
        except Exception as e:
            logger.warning(f"读取版式转MD配置失败，使用默认值: {e}")
            default_extract_images = True
            default_extract_ocr = False

        try:
            default_enable_optimization = self.config_manager.get_layout_to_md_enable_optimization()
            default_optimization_type_id = self.config_manager.get_layout_to_md_optimization_type()
        except Exception as e:
            logger.warning(f"读取版式优化配置失败，使用默认值: {e}")
            default_enable_optimization = False
            default_optimization_type_id = "invoice_cn"

        # 创建选项变量
        self.pdf_extract_images_var = tk.BooleanVar(value=default_extract_images)
        self.pdf_extract_ocr_var = tk.BooleanVar(value=default_extract_ocr)

        # 创建导出选项处理器
        self._pdf_export_handler = ExportOptionHandler(self.pdf_extract_images_var, self.pdf_extract_ocr_var)

        # 多选框容器
        checkbox_container = tb.Frame(extraction_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        checkbox_container.grid_columnconfigure(0, weight=0)
        checkbox_container.grid_columnconfigure(1, weight=0)

        # 提取图片 + 信息图标
        extract_images_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_images_container.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(20)))

        self.extract_images_check = tb.Checkbutton(
            extract_images_container,
            text=t("action_panel.extract_images"),
            variable=self.pdf_extract_images_var,
            command=self._pdf_export_handler.on_option_changed,
            bootstyle="round-toggle",
        )
        self.extract_images_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        extract_images_info = create_info_icon(
            extract_images_container, t("action_panel.extract_images_tooltip"), bootstyle="info"
        )
        extract_images_info.pack(side=tk.LEFT)

        # 图片文字识别 + 信息图标
        extract_ocr_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_ocr_container.grid(row=0, column=1, sticky="w", padx=(0, scale(10)))

        self.extract_ocr_check = tb.Checkbutton(
            extract_ocr_container,
            text=t("action_panel.ocr"),
            variable=self.pdf_extract_ocr_var,
            command=self._pdf_export_handler.on_option_changed,
            bootstyle="round-toggle",
        )
        self.extract_ocr_check.pack(side=tk.LEFT, padx=(0, scale(5)))

        extract_ocr_info = create_info_icon(extract_ocr_container, t("action_panel.ocr_tooltip"), bootstyle="info")
        extract_ocr_info.pack(side=tk.LEFT)

        self._layout_optimization_types = self.config_manager.get_localized_optimization_types(scope="layout_to_md")
        optimization_type_names = list(self._layout_optimization_types.values())
        optimization_type_id_to_name = self._layout_optimization_types

        self._has_layout_optimization_types = len(self._layout_optimization_types) > 0

        if default_optimization_type_id in optimization_type_id_to_name:
            default_optimization_type = optimization_type_id_to_name[default_optimization_type_id]
        elif optimization_type_names:
            default_optimization_type = optimization_type_names[0]
            default_optimization_type_id = next(iter(self._layout_optimization_types.keys()))
        else:
            default_optimization_type = ""
            default_enable_optimization = False

        self._default_layout_optimization_type_id = default_optimization_type_id

        if self._has_layout_optimization_types:
            separator = tb.Separator(checkbox_container, bootstyle="info")
            separator.grid(row=1, column=0, columnspan=2, sticky="ew", padx=0, pady=scale(20))

            optimization_container = tb.Frame(checkbox_container, bootstyle="default")
            optimization_container.grid(
                row=2, column=0, columnspan=2, sticky="w", padx=(scale(10), scale(10)), pady=(0, scale(5))
            )

            self.layout_enable_optimization_var = tk.BooleanVar(value=default_enable_optimization)
            self.layout_enable_optimization_check = tb.Checkbutton(
                optimization_container,
                text=t("action_panel.layout.optimize_for_type"),
                variable=self.layout_enable_optimization_var,
                command=self._on_layout_optimization_toggle,
                bootstyle="round-toggle",
            )
            self.layout_enable_optimization_check.pack(side=tk.LEFT, padx=(0, scale(5)))

            optimization_check_info = create_info_icon(
                optimization_container,
                t("action_panel.layout.optimize_for_type_tooltip"),
                bootstyle="info",
            )
            optimization_check_info.pack(side=tk.LEFT, padx=(0, scale(10)))

            self.layout_optimization_type_var = tk.StringVar(value=default_optimization_type)
            self.layout_optimization_type_combo = tb.Combobox(
                optimization_container,
                textvariable=self.layout_optimization_type_var,
                values=optimization_type_names,
                state="readonly" if default_enable_optimization else "disabled",
                width=10,
            )
            self.layout_optimization_type_combo.pack(side=tk.LEFT)

        logger.debug("PDF转换按钮创建完成")

    def _on_layout_optimization_toggle(self: ActionPanelHost):
        if self.layout_enable_optimization_var.get():
            self.layout_optimization_type_combo.config(state="readonly")
        else:
            self.layout_optimization_type_combo.config(state="disabled")

    def _on_convert_layout_to_md_clicked(self: ActionPanelHost):
        """处理版式文件转Markdown按钮点击事件"""
        if self.on_action:
            extract_image = self.pdf_extract_images_var.get() if hasattr(self, "pdf_extract_images_var") else False
            extract_ocr = self.pdf_extract_ocr_var.get() if hasattr(self, "pdf_extract_ocr_var") else False

            enable_optimization = (
                self.layout_enable_optimization_var.get() if hasattr(self, "layout_enable_optimization_var") else False
            )
            optimization_type = (
                self.layout_optimization_type_var.get() if hasattr(self, "layout_optimization_type_var") else ""
            )

            if enable_optimization:
                optimization_types = getattr(self, "_layout_optimization_types", {}) or {}
                name_to_id = {name: type_id for type_id, name in optimization_types.items()}
                optimize_for_type = name_to_id.get(
                    optimization_type, getattr(self, "_default_layout_optimization_type_id", "invoice_cn")
                )
            else:
                optimize_for_type = None

            options = {"extract_image": extract_image, "extract_ocr": extract_ocr, "optimize_for_type": optimize_for_type}

            logger.info(
                f"版式文件转Markdown - 提取选项: 图片={extract_image}, OCR={extract_ocr}, 优化类型={optimize_for_type}"
            )
            self.on_action("convert_layout_to_md", self.file_path, options)
