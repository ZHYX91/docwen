"""
操作面板组件 - 提供文件转换操作界面

根据文件类型动态显示不同的操作按钮和选项：
- MD文件：提供文档格式转换（DOCX/DOC/ODT/RTF）和表格格式转换（XLSX/XLS/ODS/CSV）
- 文档文件：提供导出Markdown功能
- 表格文件：提供导出Markdown功能
- 图片文件：提供OCR识别和导出Markdown功能
- 版式文件：提供导出Markdown功能

特性：
- 使用ttkbootstrap样式，支持主题切换
- 按钮居中排列，支持DPI缩放
- 支持文档生成的校对选项配置
- 支持取消正在进行的操作
"""

import logging
import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional, Dict
from gongwen_converter.utils.font_utils import get_default_font, get_title_font, get_small_font
from gongwen_converter.utils.dpi_utils import scale
from gongwen_converter.utils.gui_utils import ConditionalToolTip, ToolTip, create_info_icon

# 导入ttkbootstrap用于界面美化和样式管理
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# 配置日志
logger = logging.getLogger(__name__)


class ActionPanel(tb.Frame):
    """
    操作面板组件
    
    根据文件类型动态显示相应的转换按钮和选项。
    
    特性:
    - 使用ttkbootstrap统一管理样式
    - 按钮水平居中排列
    - 支持MD文件的多格式转换
    - 提供Markdown导出功能
    - 支持文档生成的校对选项
    - 使用grid布局管理器
    """

    def __init__(self, master, config_manager: any, on_action: Optional[Callable] = None, on_cancel: Optional[Callable] = None, **kwargs):
        """
        初始化操作面板组件
        
        参数:
            master: 父组件对象
            config_manager: 配置管理器实例，用于读取和保存校对选项
            on_action: 操作按钮点击回调函数，接收(action_type, file_path, options)参数
            on_cancel: 取消按钮点击回调函数，用于中止正在进行的操作
            **kwargs: 传递给Frame的其他参数
        """
        super().__init__(master, **kwargs)
        logger.debug("初始化操作面板组件 - 使用ttkbootstrap样式和grid布局")
        
        # 存储配置管理器和回调函数
        self.config_manager = config_manager
        self.on_action = on_action
        self.on_cancel = on_cancel
        
        # 存储文件类型和状态
        self.file_type: Optional[str] = None
        self.file_path: Optional[str] = None
        
        # 存储多选框变量
        self.checkbox_vars: Dict[str, tk.BooleanVar] = {}
        
        # 选项更新标志位（防止递归触发）
        self._updating_pdf_options: bool = False
        self._updating_doc_options: bool = False
        self._updating_table_options: bool = False
        self._updating_image_options: bool = False
        
        # 文档导出选项状态记录
        self._doc_last_image_state: bool = False
        self._doc_last_ocr_state: bool = False
        
        # 表格导出选项状态记录
        self._table_last_image_state: bool = False
        self._table_last_ocr_state: bool = False
        
        # 图片导出选项状态记录
        self._image_last_image_state: bool = False
        self._image_last_ocr_state: bool = False
        
        # 存储按钮引用（按文件类型组织）
        self.convert_docx_button: Optional[tb.Button] = None
        self.convert_doc_button: Optional[tb.Button] = None
        self.convert_odt_button: Optional[tb.Button] = None
        self.convert_rtf_button: Optional[tb.Button] = None
        self.convert_excel_button: Optional[tb.Button] = None
        self.convert_xls_button: Optional[tb.Button] = None
        self.convert_ods_button: Optional[tb.Button] = None
        self.convert_csv_button: Optional[tb.Button] = None
        self.convert_document_to_md_button: Optional[tb.Button] = None
        self.convert_spreadsheet_to_md_button: Optional[tb.Button] = None
        self.convert_image_to_md_button: Optional[tb.Button] = None
        self.convert_layout_to_md_button: Optional[tb.Button] = None
        self.cancel_button: Optional[tb.Button] = None
        
        # 获取字体配置（优化：只在初始化时获取一次）
        self.default_font, self.default_size = get_default_font()
        self.title_font, self.title_size = get_title_font()
        self.small_font, self.small_size = get_small_font()
        
        # 创建支持DPI的按钮样式配置
        self.button_style_3 = {
            'width': 12,
            'padding': (scale(10), scale(5))
        }
        self.button_style_2 = {
            'width': 16,
            'padding': (scale(10), scale(5))
        }
        self.button_style_1 = {
            'width': 20,
            'padding': (scale(10), scale(5))
        }
        
        # 按钮颜色映射（使用ttkbootstrap样式名称）
        self.button_colors = {
            'primary': 'primary',      # 主题色：用于主要格式转换按钮 (DOCX/XLSX)
            'secondary': 'secondary',  # 通常灰色：用于"取消"按钮
            'success': 'success',      # 通常绿色：用于开放格式转换按钮 (ODT/ODS) 与 通用格式转换按钮 (Markdown)
            'info': 'info',            # 通常蓝色：用于旧版格式转换按钮 (DOC/XLS) 与 "汇总表格"按钮、"校对"按钮
            'warning': 'warning',      # 通常橙色：用于CSV, RTF
            'danger': 'danger'         # 通常红色：用于PDF
        }
        
        # 配置grid布局权重
        self.grid_rowconfigure(0, weight=1)  # 主内容区域
        self.grid_columnconfigure(0, weight=1)  # 单列
        
        # 创建界面元素
        self._create_widgets()
        
        # 默认隐藏组件
        self.hide()
        
        logger.info("操作面板组件初始化完成")
        
    def _create_widgets(self):
        """
        创建界面元素
        
        构建操作面板的基本结构：
        - 主框架：包含所有内容的容器
        - 状态标签：显示当前操作状态或错误信息
        - 按钮框架：动态显示操作按钮
        - 取消按钮：用于中止操作（默认隐藏）
        
        所有组件使用grid布局管理，支持DPI缩放。
        """
        logger.debug("创建操作面板界面元素 - 使用grid布局")

        # 创建主框架 - 使用ttkbootstrap卡片样式
        self.main_frame = tb.Frame(self, bootstyle="default")
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=scale(5))
        
        # 配置主框架网格权重（从3行改为2行布局）
        self.main_frame.grid_rowconfigure(0, weight=0)  # 状态区域
        self.main_frame.grid_rowconfigure(1, weight=0)  # 按钮区域（包含选项）
        self.main_frame.grid_columnconfigure(0, weight=1)  # 单列

        # 创建状态标签（放在最上方）
        self.status_var = tk.StringVar(value="")
        self.status_label = tb.Label(
            self.main_frame,
            textvariable=self.status_var,
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            wraplength=scale(400),  # 使用DPI适配
            anchor=tk.CENTER  # 文本居中
        )
        self.status_label.grid(row=0, column=0, sticky="ew", pady=(0, scale(10)))

        # 创建按钮框架 - 使用grid布局，现在包含选项区域
        self.button_frame = tb.Frame(self.main_frame, bootstyle="default")
        self.button_frame.grid(row=1, column=0, sticky="ew", pady=(0, scale(5)))

        # 配置按钮框架网格权重（现在包含按钮和选项）
        self.button_frame.grid_rowconfigure(0, weight=0)  # 按钮容器
        self.button_frame.grid_rowconfigure(1, weight=0)  # 选项容器
        self.button_frame.grid_columnconfigure(0, weight=1)  # 单列

        # 创建按钮容器框架 - 用于水平居中
        self.button_container = tb.Frame(self.button_frame, bootstyle="default")
        self.button_container.grid(row=0, column=0, sticky="ew", pady=(0, scale(5)))
        self.button_container.grid_columnconfigure(0, weight=1) # 全局配置居中
        
        # 创建取消按钮（默认隐藏）
        self.cancel_button_container = tb.Frame(self.button_frame, bootstyle="default")
        self.cancel_button_container.grid(row=0, column=0, sticky="ew")
        self.cancel_button_container.grid_remove() # 默认隐藏

        self.cancel_button = tb.Button(
            self.cancel_button_container, text="❌ 取消", command=self._on_cancel_clicked,
            bootstyle=self.button_colors['danger'], **self.button_style_1
        )
        self.cancel_button.pack() # Use pack to center a single button easily

        # 配置按钮容器框架网格权重（用于居中）
        self.button_frame.grid_columnconfigure(0, weight=1)
        self.button_frame.grid_rowconfigure(0, weight=1)
        
        
        logger.debug("操作面板界面元素创建完成")
        
    def setup_for_md_to_document(self, file_path: str):
        """
        设置为MD转文档处理模式
        
        为MD文件显示文档转换按钮（DOCX/DOC/ODT/RTF）和校对选项。
        
        参数:
            file_path: MD文件路径
        """
        logger.debug(f"设置MD转文档处理模式: {file_path}")
        self.file_type = "docx"
        self.file_path = file_path
        self._clear_buttons()
        self._clear_options()
        self._create_md_to_document_buttons()
        self._create_md_to_document_options()
        self.status_var.set(f"准备处理文本文件 - 转为文档")
        logger.info("MD转文档操作面板设置完成")
    
    def setup_for_document_file(self, file_path: str, file_list: list = None):
        """
        设置为文档文件处理模式
        
        为DOCX/DOC等文档文件显示：
        - 导出Markdown按钮
        - 导出选项（提取图片、图片文字识别）
        
        参数:
            file_path: 文档文件路径
            file_list: 文件列表（批量模式时用于更新按钮状态）
        """
        logger.debug(f"设置文档文件处理模式: {file_path}")
        self.file_type = "document"
        self.file_path = file_path
        self._clear_buttons()
        self._clear_options()
        self._create_document_conversion_buttons()
        
        self.status_var.set(f"准备处理文档文件")
        logger.info("文档文件操作面板设置完成")
    
    def setup_for_md_to_spreadsheet(self, file_path: str):
        """
        设置为MD转表格处理模式
        
        为MD文件显示表格转换按钮（XLSX/XLS/ODS/CSV）。
        
        参数:
            file_path: MD文件路径
        """
        logger.debug(f"设置MD转表格处理模式: {file_path}")
        self.file_type = "xlsx"
        self.file_path = file_path
        self._clear_buttons()
        self._clear_options()
        self._create_md_to_spreadsheet_buttons()
        self.status_var.set(f"准备处理文本文件 - 转为表格")
        logger.info("MD转表格操作面板设置完成")

    def setup_for_spreadsheet_file(self, file_path: str, file_list: list = None):
        """
        设置为表格文件处理模式
        
        为XLSX/XLS/CSV等表格文件显示：
        - 导出Markdown按钮
        - 导出选项（提取图片、图片文字识别）
        
        参数:
            file_path: 表格文件路径
            file_list: 文件列表（批量模式时用于更新按钮状态）
        """
        logger.debug(f"设置表格文件处理模式: {file_path}")
        self.file_type = "spreadsheet"
        self.file_path = file_path
        self._clear_buttons()
        self._clear_options()
        self._create_spreadsheet_to_md_button()
        
        self.status_var.set("准备处理表格文件")
        logger.info("表格文件操作面板设置完成")
    
    def setup_for_image_file(self, file_path: str):
        """
        设置为图片文件处理模式
        
        为图片文件显示：
        - 导出Markdown按钮（OCR识别）
        - 导出选项（提取图片、图片文字识别）
        
        参数:
            file_path: 图片文件路径
        """
        logger.debug(f"设置图片文件处理模式: {file_path}")
        self.file_type = "image"
        self.file_path = file_path
        self._clear_buttons()
        self._clear_options()
        self._create_image_conversion_buttons()
        
        self.status_var.set("准备处理图片文件")
        logger.info("图片文件操作面板设置完成")
    
    def setup_for_layout_file(self, file_path: str):
        """
        设置为版式文件处理模式
        
        为PDF/OFD等版式文件显示导出Markdown按钮和提取选项。
        
        参数:
            file_path: 版式文件路径
        """
        logger.debug(f"设置版式文件处理模式: {file_path}")
        self.file_type = "layout"
        self.file_path = file_path
        self._clear_buttons()
        self._clear_options()
        self._create_layout_conversion_buttons()
        
        self.status_var.set("准备处理版式文件")
        logger.info("版式文件操作面板设置完成")
    
    def _create_md_to_document_buttons(self):
        """
        创建MD转文档格式的按钮
        
        显示两行按钮：
        - 第一行：生成DOCX（主色调）、生成DOC（蓝色调）
        - 第二行：生成ODT（绿色调）、生成RTF（橙色调）
        
        按钮水平居中排列，支持DPI缩放。
        """
        logger.debug("创建MD到文档系列的按钮 - 两行布局")
        
        # 第一行：生成DOCX | 生成DOC
        first_row_frame = tb.Frame(self.button_container, bootstyle="default")
        first_row_frame.grid(row=0, column=0, pady=(0, scale(10)))
        
        self.convert_docx_button = tb.Button(
            first_row_frame, text="📝 生成 DOCX", command=self._on_convert_docx_clicked,
            bootstyle=self.button_colors['primary'], **self.button_style_2
        )
        self.convert_docx_button.grid(row=0, column=0, padx=(0, scale(25)))
        ToolTip(self.convert_docx_button, "转换为Word/WPS文档格式（推荐）")
        
        self.convert_doc_button = tb.Button(
            first_row_frame, text="📝 生成 DOC", command=self._on_convert_doc_clicked,
            bootstyle=self.button_colors['info'], **self.button_style_2
        )
        self.convert_doc_button.grid(row=0, column=1)
        # 为DOC按钮添加tooltip提示
        ToolTip(self.convert_doc_button, "需要通过本地安装的 WPS、Microsoft Office 或 LibreOffice 进行转换，用户可自行设置使用软件的优先级。")
        
        # 第二行：生成ODT | 生成RTF
        second_row_frame = tb.Frame(self.button_container, bootstyle="default")
        second_row_frame.grid(row=1, column=0, pady=(0, scale(10)))
        
        self.convert_odt_button = tb.Button(
            second_row_frame, text="📝 生成 ODT", command=self._on_convert_odt_clicked,
            bootstyle=self.button_colors['success'], **self.button_style_2
        )
        self.convert_odt_button.grid(row=0, column=0, padx=(0, scale(25)))
        # 为ODT按钮添加tooltip提示
        ToolTip(self.convert_odt_button, "需要通过本地安装的 Microsoft Office 或 LibreOffice 进行转换，用户可自行设置使用软件的优先级。")
        
        self.convert_rtf_button = tb.Button(
            second_row_frame, text="📝 生成 RTF", command=self._on_convert_rtf_clicked,
            bootstyle=self.button_colors['warning'], **self.button_style_2
        )
        self.convert_rtf_button.grid(row=0, column=1)
        # 为RTF按钮添加tooltip提示
        ToolTip(self.convert_rtf_button, "需要通过本地安装的 WPS、Microsoft Office 或 LibreOffice 进行转换，用户可自行设置使用软件的优先级。")
        
        logger.debug("MD到文档系列按钮创建完成 - 两行布局。")

    def _create_document_conversion_buttons(self):
        """
        创建文档文件转换按钮
        
        显示简化的按钮布局：
        - 第一行：导出Markdown按钮
        - 第二行：导出选项（提取图片、图片文字识别）
        - 第三行：针对优化选项（复选框 + 下拉框）
        
        所有按钮居中排列，支持DPI缩放。
        """
        logger.debug("创建简化的文档转换按钮")

        # 第一行：导出Markdown按钮（始终可用）
        self.convert_document_to_md_button = tb.Button(
            self.button_container, text="✏️ 导出 Markdown", command=self._on_convert_document_to_md_clicked,
            bootstyle=self.button_colors['success'], **self.button_style_1
        )
        self.convert_document_to_md_button.grid(row=0, column=0, pady=(0, scale(10)))
        ToolTip(self.convert_document_to_md_button, "将文档转换为Markdown格式\n可选：提取图片、图片文字识别（OCR）")
        
        # 第二行：导出选项边框
        doc_export_options_frame = tb.Labelframe(
            self.button_container, 
            text="导出选项",
            bootstyle="info"
        )
        doc_export_options_frame.grid(row=1, column=0, sticky="ew", padx=scale(20), pady=scale(10))
        
        # 配置选项框架网格权重（用于居中内部容器）
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
        
        # 创建两个独立的多选框变量
        self.doc_extract_image_var = tk.BooleanVar(value=default_extract_image)
        self.doc_extract_ocr_var = tk.BooleanVar(value=default_extract_ocr)
        
        # 初始化状态记录
        self._doc_last_image_state = default_extract_image
        self._doc_last_ocr_state = default_extract_ocr
        
        # 多选框容器，放在边框内（支持2行布局）
        checkbox_container = tb.Frame(doc_export_options_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置多选框容器的行列权重
        checkbox_container.grid_rowconfigure(0, weight=0)  # 第1行：图片选项
        checkbox_container.grid_rowconfigure(1, weight=0)  # 第2行：分割线
        checkbox_container.grid_rowconfigure(2, weight=0)  # 第3行：优化选项
        checkbox_container.grid_columnconfigure(0, weight=0)
        checkbox_container.grid_columnconfigure(1, weight=0)
        
        # ===== 第1行：图片选项 =====
        # 提取图片 + 信息图标
        extract_image_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_image_container.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(20)))
        
        self.doc_extract_image_check = tb.Checkbutton(
            extract_image_container, text="提取图片",
            variable=self.doc_extract_image_var,
            command=self._on_doc_export_option_changed,
            bootstyle="round-toggle"
        )
        self.doc_extract_image_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        extract_image_info = create_info_icon(
            extract_image_container,
            "无损提取嵌入的图片文件，并在Markdown文件中添加链接。",
            bootstyle="info"
        )
        extract_image_info.pack(side=tk.LEFT)
        
        # 图片文字识别 + 信息图标
        extract_ocr_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_ocr_container.grid(row=0, column=1, sticky="w", padx=(0, scale(10)))
        
        self.doc_extract_ocr_check = tb.Checkbutton(
            extract_ocr_container, text="图片文字识别",
            variable=self.doc_extract_ocr_var,
            command=self._on_doc_export_option_changed,
            bootstyle="round-toggle"
        )
        self.doc_extract_ocr_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        extract_ocr_info = create_info_icon(
            extract_ocr_container,
            "对嵌入的图片进行OCR（离线），\n"
            "结果输出至单独的Markdown文件，\n"
            "并将文件链接添加至主要Markdown文件中。\n"
            "耗时可能较长",
            bootstyle="info"
        )
        extract_ocr_info.pack(side=tk.LEFT)
        
        # ===== 第2行：分割线 =====
        separator = tb.Separator(checkbox_container, bootstyle="info")
        separator.grid(row=1, column=0, columnspan=2, sticky="ew", padx=0, pady=scale(20))
        
        # ===== 第3行：优化选项（单行布局）=====
        optimization_container = tb.Frame(checkbox_container, bootstyle="default")
        optimization_container.grid(row=2, column=0, columnspan=2, sticky="", padx=(scale(10), scale(10)))
        
        # 复选框
        self.doc_enable_optimization_var = tk.BooleanVar(value=True)  # 默认勾选
        self.doc_enable_optimization_check = tb.Checkbutton(
            optimization_container, text="针对优化",
            variable=self.doc_enable_optimization_var,
            command=self._on_doc_optimization_toggle,
            bootstyle="round-toggle"
        )
        self.doc_enable_optimization_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        # 信息图标
        optimization_check_info = create_info_icon(
            optimization_container,
            "启用针对特定文档类型的优化转换\n"
            "公文：识别公文元素，生成YAML元数据",
            bootstyle="info"
        )
        optimization_check_info.pack(side=tk.LEFT, padx=(0, scale(10)))
        
        # 优化类型下拉框
        self.doc_optimization_type_var = tk.StringVar(value="公文")
        self.doc_optimization_type_combo = tb.Combobox(
            optimization_container,
            textvariable=self.doc_optimization_type_var,
            values=["公文"],  # 未来可扩展: ["公文", "合同", "论文"]
            state="readonly",
            width=10
        )
        self.doc_optimization_type_combo.pack(side=tk.LEFT)
        
        logger.debug("文档文件转换按钮创建完成（含导出选项和优化选项）。")

    def _create_spreadsheet_to_md_button(self):
        """
        创建表格文件转换按钮
        
        显示简化的按钮布局：
        - 第一行：导出Markdown按钮
        - 第二行：导出选项（提取图片、图片文字识别）
        
        所有按钮居中排列，支持DPI缩放。
        """
        logger.debug("创建表格转换按钮（含导出选项）")
        
        # 第一行：导出Markdown按钮（始终可用）
        self.convert_spreadsheet_to_md_button = tb.Button(
            self.button_container, text="✏️ 导出 Markdown", command=self._on_convert_spreadsheet_to_md_clicked,
            bootstyle=self.button_colors['success'], **self.button_style_1
        )
        self.convert_spreadsheet_to_md_button.grid(row=0, column=0, pady=(0, scale(10)))
        ToolTip(self.convert_spreadsheet_to_md_button, "将表格转换为Markdown格式\n可选：提取图片、图片文字识别（OCR）")
        
        # 第二行：导出选项边框
        table_export_options_frame = tb.Labelframe(
            self.button_container, 
            text="导出选项",
            bootstyle="info"
        )
        table_export_options_frame.grid(row=1, column=0, sticky="ew", padx=scale(20), pady=scale(10))
        
        # 配置选项框架网格权重（用于居中内部容器）
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
        
        # 创建两个独立的多选框变量
        self.table_extract_image_var = tk.BooleanVar(value=default_extract_image)
        self.table_extract_ocr_var = tk.BooleanVar(value=default_extract_ocr)
        
        # 初始化状态记录
        self._table_last_image_state = default_extract_image
        self._table_last_ocr_state = default_extract_ocr
        
        # 多选框容器，放在边框内
        checkbox_container = tb.Frame(table_export_options_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置多选框容器的列权重
        checkbox_container.grid_columnconfigure(0, weight=0)
        checkbox_container.grid_columnconfigure(1, weight=0)
        
        # 提取图片 + 信息图标
        extract_image_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_image_container.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(20)))
        
        self.table_extract_image_check = tb.Checkbutton(
            extract_image_container, text="提取图片",
            variable=self.table_extract_image_var,
            command=self._on_table_export_option_changed,
            bootstyle="round-toggle"
        )
        self.table_extract_image_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        extract_image_info = create_info_icon(
            extract_image_container,
            "无损提取嵌入的图片文件，并在Markdown文件中添加链接。",
            bootstyle="info"
        )
        extract_image_info.pack(side=tk.LEFT)
        
        # 图片文字识别 + 信息图标
        extract_ocr_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_ocr_container.grid(row=0, column=1, sticky="w", padx=(0, scale(10)))
        
        self.table_extract_ocr_check = tb.Checkbutton(
            extract_ocr_container, text="图片文字识别",
            variable=self.table_extract_ocr_var,
            command=self._on_table_export_option_changed,
            bootstyle="round-toggle"
        )
        self.table_extract_ocr_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        extract_ocr_info = create_info_icon(
            extract_ocr_container,
            "对嵌入的图片进行OCR（离线），\n"
            "结果输出至单独的Markdown文件，\n"
            "并将文件链接添加至主要Markdown文件中。\n"
            "耗时可能较长",
            bootstyle="info"
        )
        extract_ocr_info.pack(side=tk.LEFT)
        
        logger.debug("表格文件转换按钮创建完成。")

    def _create_md_to_document_options(self):
        """
        创建MD转文档的校对选项
        
        显示四个校对选项的复选框：
        - 标点配对
        - 错别字校对
        - 符号校对
        - 敏感词匹配
        
        选项排列为2列2行，居中显示，默认值从配置文件读取。
        """
        logger.debug("创建MD转文档的生成选项 - 优化边距和对齐")

        # 创建生成选项边框 - 添加左右边距避免铺满面板
        self.md_to_doc_options_frame = tb.Labelframe(
            self.button_container, 
            text="生成选项",
            bootstyle="info"
        )
        self.md_to_doc_options_frame.grid(row=2, column=0, sticky="ew", padx=scale(20), pady=scale(10))
        
        # 配置选项框架网格权重
        self.md_to_doc_options_frame.grid_rowconfigure(0, weight=1)
        self.md_to_doc_options_frame.grid_columnconfigure(0, weight=1)

        # 创建一个容器来包裹所有的多选框，以便将它们作为一个整体居中
        checkbox_container = tb.Frame(self.md_to_doc_options_frame)
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置复选框容器的列权重，确保选项平均分布
        checkbox_container.grid_columnconfigure(0, weight=1)
        checkbox_container.grid_columnconfigure(1, weight=1)

        default_options = self._get_default_options()
        options = [
            ("标点配对", "symbol_pairing", "检查文档中的标点符号配对情况，如括号、引号等是否成对出现"),
            ("错别字校对", "typos_rule", "检查文档中的错别字，根据用户自定义的错别字词库进行匹配"),
            ("符号校对", "symbol_correction", "检查文档中的标点符号使用规范，如全角和半角符号误用等"),
            ("敏感词匹配", "sensitive_word", "检查文档中的敏感词，根据用户自定义的敏感词库进行匹配")
        ]
        
        for i, (text, key, tooltip_text) in enumerate(options):
            var = tk.BooleanVar(value=default_options.get(key, False))
            self.checkbox_vars[key] = var
            
            # 创建容器用于放置复选框和信息图标
            option_frame = tb.Frame(checkbox_container, bootstyle="default")
            row, col = divmod(i, 2)
            option_frame.grid(row=row, column=col, sticky="", padx=scale(10), pady=scale(5))
            
            # 创建复选框
            checkbox = tb.Checkbutton(
                option_frame, text=text, variable=var,
                command=self._on_option_changed, bootstyle="round-toggle"
            )
            checkbox.pack(side=tk.LEFT, padx=(0, scale(5)))
            
            # 添加信息图标
            info_icon = create_info_icon(option_frame, tooltip_text, bootstyle="info")
            info_icon.pack(side=tk.LEFT)
        
        logger.debug("MD转文档生成选项创建完成 - 优化边距和对齐")
    
    def _create_md_to_spreadsheet_buttons(self):
        """
        创建MD转电子表格格式的按钮
        
        显示两行按钮：
        - 第一行：生成XLSX（主色调）、生成XLS（蓝色调）
        - 第二行：生成ODS（绿色调）、生成CSV（橙色调）
        
        按钮水平居中排列，支持DPI缩放。
        """
        logger.debug("创建MD到电子表格系列的按钮 - 两行布局")
        
        # 第一行：生成XLSX | 生成XLS
        first_row_frame = tb.Frame(self.button_container, bootstyle="default")
        first_row_frame.grid(row=0, column=0, pady=(0, scale(10)))
        
        self.convert_excel_button = tb.Button(
            first_row_frame, text="📊 生成 XLSX", command=self._on_convert_excel_clicked,
            bootstyle=self.button_colors['primary'], **self.button_style_2
        )
        self.convert_excel_button.grid(row=0, column=0, padx=(0, scale(25)))
        ToolTip(self.convert_excel_button, "转换为Excel/WPS表格格式（推荐）")
        
        self.convert_xls_button = tb.Button(
            first_row_frame, text="📊 生成 XLS", command=self._on_convert_xls_clicked,
            bootstyle=self.button_colors['info'], **self.button_style_2
        )
        self.convert_xls_button.grid(row=0, column=1)
        # 为XLS按钮添加tooltip提示
        ToolTip(self.convert_xls_button, "需要通过本地安装的 WPS、Microsoft Office 或 LibreOffice 进行转换，用户可自行设置使用软件的优先级。")
        
        # 第二行：生成ODS | 生成CSV
        second_row_frame = tb.Frame(self.button_container, bootstyle="default")
        second_row_frame.grid(row=1, column=0, pady=(0, scale(10)))
        
        self.convert_ods_button = tb.Button(
            second_row_frame, text="📊 生成 ODS", command=self._on_convert_ods_clicked,
            bootstyle=self.button_colors['success'], **self.button_style_2
        )
        self.convert_ods_button.grid(row=0, column=0, padx=(0, scale(25)))
        # 为ODS按钮添加tooltip提示
        ToolTip(self.convert_ods_button, "需要通过本地安装的 Microsoft Office 或 LibreOffice 进行转换，用户可自行设置使用软件的优先级。")
        
        self.convert_csv_button = tb.Button(
            second_row_frame, text="📊 生成 CSV", command=self._on_convert_csv_clicked,
            bootstyle=self.button_colors['warning'], **self.button_style_2
        )
        self.convert_csv_button.grid(row=0, column=1)
        ToolTip(self.convert_csv_button, "转换为纯文本表格格式，可能需要本地Office软件预处理")
        
        logger.debug("MD到电子表格系列按钮创建完成 - 两行布局。")
    
    def _create_image_conversion_buttons(self):
        """
        创建图片文件转换按钮
        
        显示按钮和选项：
        - 第一行：导出Markdown按钮
        - 第二行：导出选项（提取图片、图片文字识别）
        
        按钮居中排列，支持DPI缩放。
        """
        logger.debug("创建图片转换按钮")
        
        # 第一行：导出Markdown按钮（初始启用，默认两个选项都勾选）
        self.convert_image_to_md_button = tb.Button(
            self.button_container, text="✏️ 导出 Markdown", command=self._on_convert_image_to_md_clicked,
            bootstyle=self.button_colors['success'], **self.button_style_1
        )
        self.convert_image_to_md_button.grid(row=0, column=0, pady=(0, scale(10)))
        ToolTip(self.convert_image_to_md_button, "将图片转换为Markdown格式\n可选：提取图片、图片文字识别（OCR）")
        
        # 第二行：导出选项边框
        export_options_frame = tb.Labelframe(
            self.button_container, 
            text="导出选项",
            bootstyle="info"
        )
        export_options_frame.grid(row=1, column=0, sticky="ew", padx=scale(20), pady=scale(10))
        
        # 配置选项框架网格权重（用于居中内部容器）
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
        
        # 创建两个独立的多选框变量
        self.image_extract_image_var = tk.BooleanVar(value=default_extract_image)
        self.image_extract_ocr_var = tk.BooleanVar(value=default_extract_ocr)
        
        # 初始化状态记录
        self._image_last_image_state = default_extract_image
        self._image_last_ocr_state = default_extract_ocr
        
        # 多选框容器，放在边框内
        checkbox_container = tb.Frame(export_options_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置多选框容器的列权重
        checkbox_container.grid_columnconfigure(0, weight=0)
        checkbox_container.grid_columnconfigure(1, weight=0)
        
        # 提取图片 + 信息图标
        extract_image_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_image_container.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(20)))
        
        self.extract_image_check = tb.Checkbutton(
            extract_image_container, text="提取图片",
            variable=self.image_extract_image_var,
            command=self._on_image_export_option_changed,
            bootstyle="round-toggle"
        )
        self.extract_image_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        extract_image_info = create_info_icon(
            extract_image_container,
            "在Markdown文件中插入图片链接。",
            bootstyle="info"
        )
        extract_image_info.pack(side=tk.LEFT)
        
        # 图片文字识别 + 信息图标
        extract_ocr_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_ocr_container.grid(row=0, column=1, sticky="w", padx=(0, scale(10)))
        
        self.extract_image_ocr_check = tb.Checkbutton(
            extract_ocr_container, text="图片文字识别",
            variable=self.image_extract_ocr_var,
            command=self._on_image_export_option_changed,
            bootstyle="round-toggle"
        )
        self.extract_image_ocr_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        extract_ocr_info = create_info_icon(
            extract_ocr_container,
            "对图片进行OCR（离线），耗时可能较长",
            bootstyle="info"
        )
        extract_ocr_info.pack(side=tk.LEFT)
        
        logger.debug("图片转换按钮创建完成")
    
    def _create_layout_conversion_buttons(self):
        """
        创建PDF文件转换按钮
        
        显示按钮和选项：
        - 第一行：导出Markdown按钮（始终可用）
        - 第二行：提取选项（2个独立多选框：提取图片、图片文字识别）
        
        逻辑说明：
        - 不提供"提取文字"选项，内部总是提取文本
        - 两个都不勾选 → 只提取文字
        - 勾选"提取图片" → 提取文字 + 提取图片
        - 勾选OCR → 自动勾选"提取图片"，提取文字 + 提取图片 + OCR
        
        所有按钮居中排列，支持DPI缩放。
        """
        logger.debug("创建PDF转换按钮 - 简化设计v2.2")
        
        # 第一行：导出Markdown按钮（始终可用）
        self.convert_layout_to_md_button = tb.Button(
            self.button_container, text="✏️ 导出 Markdown", command=self._on_convert_layout_to_md_clicked,
            bootstyle=self.button_colors['success'], **self.button_style_1
        )
        self.convert_layout_to_md_button.grid(row=0, column=0, pady=(0, scale(10)))
        ToolTip(self.convert_layout_to_md_button, "将版式文件转换为Markdown格式\n可选：提取图片、图片文字识别（OCR）")
        
        # 第二行：提取选项边框
        extraction_frame = tb.Labelframe(
            self.button_container, 
            text="提取选项",
            bootstyle="info"
        )
        extraction_frame.grid(row=1, column=0, sticky="ew", padx=scale(20), pady=scale(10))
        
        # 配置选项框架网格权重（用于居中内部容器）
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
        
        # 创建2个独立的多选框变量
        self.pdf_extract_images_var = tk.BooleanVar(value=default_extract_images)
        self.pdf_extract_ocr_var = tk.BooleanVar(value=default_extract_ocr)
        
        # 初始化状态记录（用于联动逻辑）
        self._pdf_last_images_state = default_extract_images
        self._pdf_last_ocr_state = default_extract_ocr
        
        # 多选框容器，放在边框内
        checkbox_container = tb.Frame(extraction_frame, bootstyle="default")
        checkbox_container.grid(row=0, column=0, sticky="", padx=scale(10), pady=scale(10))
        
        # 配置多选框容器的列权重
        checkbox_container.grid_columnconfigure(0, weight=0)
        checkbox_container.grid_columnconfigure(1, weight=0)
        
        # 提取图片 + 信息图标
        extract_images_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_images_container.grid(row=0, column=0, sticky="w", padx=(scale(10), scale(20)))
        
        self.extract_images_check = tb.Checkbutton(
            extract_images_container, text="提取图片",
            variable=self.pdf_extract_images_var,
            command=self._on_pdf_extraction_option_changed,
            bootstyle="round-toggle"
        )
        self.extract_images_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        extract_images_info = create_info_icon(
            extract_images_container,
            "无损提取嵌入的图片文件，并在Markdown文件中添加链接。",
            bootstyle="info"
        )
        extract_images_info.pack(side=tk.LEFT)
        
        # 图片文字识别 + 信息图标
        extract_ocr_container = tb.Frame(checkbox_container, bootstyle="default")
        extract_ocr_container.grid(row=0, column=1, sticky="w", padx=(0, scale(10)))
        
        self.extract_ocr_check = tb.Checkbutton(
            extract_ocr_container, text="图片文字识别",
            variable=self.pdf_extract_ocr_var,
            command=self._on_pdf_extraction_option_changed,
            bootstyle="round-toggle"
        )
        self.extract_ocr_check.pack(side=tk.LEFT, padx=(0, scale(5)))
        
        extract_ocr_info = create_info_icon(
            extract_ocr_container,
            "对嵌入的图片进行OCR（离线），\n"
            "结果输出至单独的Markdown文件，\n"
            "并将文件链接添加至主要Markdown文件中。\n"
            "耗时可能较长",
            bootstyle="info"
        )
        extract_ocr_info.pack(side=tk.LEFT)
        
        logger.debug("PDF转换按钮创建完成 - 2个多选框")
    
    def _get_default_options(self) -> Dict[str, bool]:
        """
        从配置文件获取默认校对选项
        
        读取symbol_settings、typos_settings和sensitive_settings配置，
        返回各校对选项的默认启用状态。
        
        返回:
            Dict[str, bool]: 校对选项字典，包含symbol_pairing、symbol_correction、
                           typos_rule、sensitive_word的启用状态
        """
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
            logger.error(f"获取默认选项失败: {str(e)}")
            return {"symbol_pairing": True, "symbol_correction": True, "typos_rule": True, "sensitive_word": False}
            
    def refresh_options(self):
        """刷新校对选项状态"""
        logger.debug("刷新选项状态以匹配当前配置")
        default_options = self._get_default_options()
        for key, var in self.checkbox_vars.items():
            if key in default_options:
                var.set(default_options[key])
                logger.debug(f"更新选项 {key} = {default_options[key]}")
        logger.info("选项状态已刷新")

    def _on_option_changed(self):
        """处理校对选项变更事件"""
        pass
    
    def _on_convert_docx_clicked(self):
        if self.on_action: self.on_action("convert_md_to_docx", self.file_path, self._get_selected_options())
    
    def _on_convert_doc_clicked(self):
        if self.on_action: self.on_action("convert_md_to_doc", self.file_path, self._get_selected_options())
    
    def _on_doc_optimization_toggle(self):
        """
        处理针对优化复选框切换事件
        
        控制下拉框的启用/禁用状态：
        - 勾选时：下拉框启用
        - 不勾选时：下拉框禁用（灰色）
        """
        if self.doc_enable_optimization_var.get():
            # 启用下拉框
            self.doc_optimization_type_combo.config(state="readonly")
            logger.debug("针对优化已启用，下拉框可选")
        else:
            # 禁用下拉框
            self.doc_optimization_type_combo.config(state="disabled")
            logger.debug("针对优化已禁用，下拉框灰色")
    
    def _on_convert_document_to_md_clicked(self):
        """
        处理文档转Markdown按钮点击事件
        
        根据UI选项构建转换参数：
        
        导出选项：
        - extract_image: 是否提取图片
        - extract_ocr: 是否进行OCR识别
        
        优化选项：
        - 不勾选"针对优化" → optimize_for_type = None（简化模式）
        - 勾选"针对优化" + 选"公文" → optimize_for_type = "gongwen"（公文模式）
        - 勾选"针对优化" + 选其他 → optimize_for_type = "contract"/"thesis"（预留）
        
        简化模式特点：
        - 基于Word样式（Title/Subtitle/Heading 1-6）转换
        - 无公文元素识别
        - YAML只有标题和副标题
        
        公文模式特点：
        - 三轮公文元素识别
        - 生成14个字段的YAML元数据
        - 附件内容单独输出
        """
        if self.on_action:
            # 获取导出选项
            extract_image = self.doc_extract_image_var.get() if hasattr(self, 'doc_extract_image_var') else True
            extract_ocr = self.doc_extract_ocr_var.get() if hasattr(self, 'doc_extract_ocr_var') else False
            
            # 获取优化选项
            enable_optimization = self.doc_enable_optimization_var.get() if hasattr(self, 'doc_enable_optimization_var') else False
            optimization_type = self.doc_optimization_type_var.get() if hasattr(self, 'doc_optimization_type_var') else "公文"
            
            # 确定优化类型参数
            if enable_optimization:
                # 映射中文 → 英文标识
                type_map = {
                    "公文": "gongwen",
                    "合同": "contract",
                    "论文": "thesis"
                }
                optimize_for_type = type_map.get(optimization_type, "gongwen")
            else:
                # 不优化 → 简化模式（None）
                optimize_for_type = None
            
            # 构建选项字典
            options = {
                'extract_image': extract_image,
                'extract_ocr': extract_ocr,
                'optimize_for_type': optimize_for_type
            }
            
            logger.info(
                f"文档转Markdown - 导出选项: "
                f"提取图片={extract_image}, OCR={extract_ocr}, "
                f"优化类型={optimize_for_type}"
            )
            self.on_action("convert_document_to_md", self.file_path, options)
    
    def _on_cancel_clicked(self):
        """
        处理取消按钮点击事件
        
        禁用取消按钮以防止重复点击，然后调用取消回调函数中止当前操作。
        """
        logger.info("取消按钮被点击")
        # 禁用按钮防止重复点击
        self.cancel_button.config(state="disabled")
        if self.on_cancel:
            self.on_cancel()
            
    def _on_convert_odt_clicked(self):
        """处理生成ODT按钮点击事件"""
        if self.on_action: self.on_action("convert_md_to_odt", self.file_path, self._get_selected_options())

    def _on_convert_rtf_clicked(self):
        """处理生成RTF按钮点击事件"""
        if self.on_action: self.on_action("convert_md_to_rtf", self.file_path, self._get_selected_options())
    
    def _on_convert_excel_clicked(self):
        if self.on_action: self.on_action("convert_md_to_xlsx", self.file_path, self._get_selected_options())

    def _on_convert_xls_clicked(self):
        if self.on_action: self.on_action("convert_md_to_xls", self.file_path, {})

    def _on_convert_ods_clicked(self):
        """处理生成ODS按钮点击事件"""
        if self.on_action: self.on_action("convert_md_to_ods", self.file_path, {})

    def _on_convert_csv_clicked(self):
        """处理生成CSV按钮点击事件"""
        if self.on_action: self.on_action("convert_md_to_csv", self.file_path, {})

    def _on_convert_spreadsheet_to_md_clicked(self):
        """处理表格转Markdown按钮点击事件"""
        if self.on_action:
            # 获取导出选项
            extract_image = self.table_extract_image_var.get() if hasattr(self, 'table_extract_image_var') else False
            extract_ocr = self.table_extract_ocr_var.get() if hasattr(self, 'table_extract_ocr_var') else False
            
            # 构建选项字典
            options = {
                'extract_image': extract_image,
                'extract_ocr': extract_ocr
            }
            
            logger.info(
                f"表格转Markdown - 导出选项: "
                f"提取图片={extract_image}, OCR={extract_ocr}"
            )
            self.on_action("convert_spreadsheet_to_md", self.file_path, options)
    
    def _on_convert_image_to_md_clicked(self):
        """处理图片转Markdown按钮点击事件"""
        if self.on_action:
            # 获取两个选项的值
            extract_image = self.image_extract_image_var.get() if hasattr(self, 'image_extract_image_var') else True
            extract_ocr = self.image_extract_ocr_var.get() if hasattr(self, 'image_extract_ocr_var') else True
            
            # 构建选项字典
            options = {
                'extract_image': extract_image,
                'extract_ocr': extract_ocr
            }
            
            logger.info(
                f"图片转Markdown - 导出选项: "
                f"提取图片={extract_image}, OCR={extract_ocr}"
            )
            self.on_action("convert_image_to_md", self.file_path, options)
    
    def _on_pdf_extraction_option_changed(self):
        """
        处理PDF提取选项变更事件（v2.2简化版）
        
        实现联动逻辑：
        1. 勾选OCR时，自动勾选"提取图片"（OCR从无到有）
        2. 取消"提取图片"时，自动取消OCR（提取图片从有到无）
        
        使用状态变化检测避免死循环联动
        """
        if self._updating_pdf_options:
            return
        
        try:
            self._updating_pdf_options = True
            
            extract_images = self.pdf_extract_images_var.get()
            extract_ocr = self.pdf_extract_ocr_var.get()
            
            # 检测状态变化
            images_changed = (extract_images != self._pdf_last_images_state)
            ocr_changed = (extract_ocr != self._pdf_last_ocr_state)
            
            # 场景A：用户勾选OCR（从无到有）
            if extract_ocr and not self._pdf_last_ocr_state and ocr_changed:
                if not extract_images:
                    logger.debug("PDF导出：勾选OCR，自动勾选提取图片")
                    self.pdf_extract_images_var.set(True)
            
            # 场景B：用户取消提取图片（从有到无）
            elif not extract_images and self._pdf_last_images_state and images_changed:
                if extract_ocr:
                    logger.debug("PDF导出：取消提取图片，自动取消OCR")
                    self.pdf_extract_ocr_var.set(False)
            
            # 更新状态记录
            self._pdf_last_images_state = self.pdf_extract_images_var.get()
            self._pdf_last_ocr_state = self.pdf_extract_ocr_var.get()
        
        finally:
            self._updating_pdf_options = False
    
    
    def _on_convert_layout_to_md_clicked(self):
        """
        处理版式文件转Markdown按钮点击事件（v2.2简化版）
        
        使用pymupdf4llm导出，受提取选项影响。
        两个选项控制行为：
        - 两个都不勾选：只提取文字
        - 勾选提取图片：提取文字 + 提取图片
        - 勾选OCR：提取文字 + 提取图片 + OCR（自动勾选提取图片）
        """
        if self.on_action:
            # 获取2个选项的值
            extract_images = self.pdf_extract_images_var.get() if hasattr(self, 'pdf_extract_images_var') else False
            extract_ocr = self.pdf_extract_ocr_var.get() if hasattr(self, 'pdf_extract_ocr_var') else False
            
            # 构建选项字典（只传递2个参数）
            options = {
                'extract_images': extract_images,
                'extract_ocr': extract_ocr
            }
            
            logger.info(
                f"版式文件转Markdown - 提取选项: "
                f"图片={extract_images}, OCR={extract_ocr}"
            )
            self.on_action("convert_layout_to_md", self.file_path, options)
    
    def _on_doc_export_option_changed(self):
        """
        处理文档导出选项变更事件
        
        实现联动逻辑（类似PDF）：
        1. 勾选OCR时，自动勾选"提取图片"（OCR从无到有）
        2. 取消"提取图片"时，自动取消OCR（提取图片从有到无）
        
        使用状态变化检测避免死循环联动
        """
        if self._updating_doc_options:
            return
        
        try:
            self._updating_doc_options = True
            
            extract_image = self.doc_extract_image_var.get()
            extract_ocr = self.doc_extract_ocr_var.get()
            
            # 检测状态变化
            image_changed = (extract_image != self._doc_last_image_state)
            ocr_changed = (extract_ocr != self._doc_last_ocr_state)
            
            # 场景A：用户勾选OCR（从无到有）
            if extract_ocr and not self._doc_last_ocr_state and ocr_changed:
                if not extract_image:
                    logger.debug("文档导出：勾选OCR，自动勾选提取图片")
                    self.doc_extract_image_var.set(True)
            
            # 场景B：用户取消提取图片（从有到无）
            elif not extract_image and self._doc_last_image_state and image_changed:
                if extract_ocr:
                    logger.debug("文档导出：取消提取图片，自动取消OCR")
                    self.doc_extract_ocr_var.set(False)
            
            # 更新状态记录
            self._doc_last_image_state = self.doc_extract_image_var.get()
            self._doc_last_ocr_state = self.doc_extract_ocr_var.get()
        
        finally:
            self._updating_doc_options = False
    
    def _on_table_export_option_changed(self):
        """
        处理表格导出选项变更事件
        
        实现联动逻辑（类似PDF）：
        1. 勾选OCR时，自动勾选"提取图片"（OCR从无到有）
        2. 取消"提取图片"时，自动取消OCR（提取图片从有到无）
        
        使用状态变化检测避免死循环联动
        """
        if self._updating_table_options:
            return
        
        try:
            self._updating_table_options = True
            
            extract_image = self.table_extract_image_var.get()
            extract_ocr = self.table_extract_ocr_var.get()
            
            # 检测状态变化
            image_changed = (extract_image != self._table_last_image_state)
            ocr_changed = (extract_ocr != self._table_last_ocr_state)
            
            # 场景A：用户勾选OCR（从无到有）
            if extract_ocr and not self._table_last_ocr_state and ocr_changed:
                if not extract_image:
                    logger.debug("表格导出：勾选OCR，自动勾选提取图片")
                    self.table_extract_image_var.set(True)
            
            # 场景B：用户取消提取图片（从有到无）
            elif not extract_image and self._table_last_image_state and image_changed:
                if extract_ocr:
                    logger.debug("表格导出：取消提取图片，自动取消OCR")
                    self.table_extract_ocr_var.set(False)
            
            # 更新状态记录
            self._table_last_image_state = self.table_extract_image_var.get()
            self._table_last_ocr_state = self.table_extract_ocr_var.get()
        
        finally:
            self._updating_table_options = False
    
    def _on_image_export_option_changed(self):
        """
        处理图片导出选项变更事件
        
        实现联动逻辑（类似PDF）：
        1. 勾选OCR时，自动勾选"提取图片"（OCR从无到有）
        2. 取消"提取图片"时，自动取消OCR（提取图片从有到无）
        3. 至少勾选一个选项时，"导出Markdown"按钮才可用
        
        使用状态变化检测避免死循环联动
        """
        if self._updating_image_options:
            return
        
        try:
            self._updating_image_options = True
            
            extract_image = self.image_extract_image_var.get()
            extract_ocr = self.image_extract_ocr_var.get()
            
            # 检测状态变化
            image_changed = (extract_image != self._image_last_image_state)
            ocr_changed = (extract_ocr != self._image_last_ocr_state)
            
            # 场景A：用户勾选OCR（从无到有）
            if extract_ocr and not self._image_last_ocr_state and ocr_changed:
                if not extract_image:
                    logger.debug("图片导出：勾选OCR，自动勾选提取图片")
                    self.image_extract_image_var.set(True)
                    extract_image = True  # 更新本地变量
            
            # 场景B：用户取消提取图片（从有到无）
            elif not extract_image and self._image_last_image_state and image_changed:
                if extract_ocr:
                    logger.debug("图片导出：取消提取图片，自动取消OCR")
                    self.image_extract_ocr_var.set(False)
                    extract_ocr = False  # 更新本地变量
            
            # 更新状态记录
            self._image_last_image_state = self.image_extract_image_var.get()
            self._image_last_ocr_state = self.image_extract_ocr_var.get()
            
            # 规则3：至少勾选一个选项才启用按钮
            final_extract_image = self.image_extract_image_var.get()
            final_extract_ocr = self.image_extract_ocr_var.get()
            should_enable = final_extract_image or final_extract_ocr
            
            if hasattr(self, 'convert_image_to_md_button') and self.convert_image_to_md_button:
                self.convert_image_to_md_button.config(state="normal" if should_enable else "disabled")
                
                if should_enable:
                    options = []
                    if final_extract_image:
                        options.append('图片')
                    if final_extract_ocr:
                        options.append('OCR')
                    logger.debug(f"导出Markdown按钮已启用：{' + '.join(options)}")
                else:
                    logger.debug("导出Markdown按钮已禁用：未选择任何选项")
        
        finally:
            self._updating_image_options = False
    
    def _get_selected_options(self) -> Dict[str, bool]:
        """获取当前选中的选项"""
        return {key: var.get() for key, var in self.checkbox_vars.items()}
    
    def _clear_buttons(self):
        """清空按钮区域"""
        for widget in self.button_container.winfo_children():
            widget.destroy()
        self.convert_docx_button = self.convert_doc_button = None
        self.convert_odt_button = self.convert_rtf_button = None
        self.convert_excel_button = self.convert_xls_button = None
        self.convert_ods_button = self.convert_csv_button = None
        self.convert_document_to_md_button = None
        self.convert_spreadsheet_to_md_button = None
        self.convert_image_to_md_button = None
        self.convert_layout_to_md_button = None
    
    def _clear_options(self):
        """清空选项区域"""
        self.checkbox_vars = {}
    
    def show(self):
        self.grid(row=0, column=0, sticky="nsew", padx=scale(1), pady=scale(5))
    
    def hide(self):
        self.grid_remove()

    def show_cancel_button(self):
        """
        显示取消按钮
        
        在开始长时间操作时调用。
        隐藏所有操作按钮，只显示取消按钮，允许用户中止操作。
        """
        self.button_container.grid_remove()
        self.cancel_button_container.grid()
        self.cancel_button.config(state="normal")

    def hide_cancel_button(self):
        """
        隐藏取消按钮
        
        在操作完成或被取消后调用。
        显示所有操作按钮，隐藏取消按钮。
        """
        self.cancel_button_container.grid_remove()
        self.button_container.grid()
    
    def set_status(self, message: str, is_error: bool = False):
        """
        设置状态消息
        
        在状态标签中显示提示或错误信息。
        
        参数:
            message: 要显示的消息文本
            is_error: 是否为错误消息，True时显示为红色
        """
        bootstyle = "danger" if is_error else "secondary"
        self.status_var.set(f"错误: {message}" if is_error else message)
        self.status_label.configure(bootstyle=bootstyle)
    
    def clear_status(self):
        self.set_status("")
    
    def get_selected_options(self) -> Dict[str, bool]:
        return self._get_selected_options()
    
    def enable_all_buttons(self):
        """
        启用所有操作按钮
        
        在操作完成或取消后调用此方法，恢复所有按钮的可用状态。
        只处理存在的按钮（因为按钮是根据文件类型动态创建的）。
        
        注意：此方法不包括取消按钮。取消按钮的状态由 show_cancel_button() 
             和 hide_cancel_button() 单独管理。
        """
        logger.debug("启用所有操作按钮")
        
        # 获取所有可能的按钮引用
        buttons = [
            self.convert_docx_button,
            self.convert_doc_button,
            self.convert_odt_button,
            self.convert_rtf_button,
            self.convert_excel_button,
            self.convert_xls_button,
            self.convert_ods_button,
            self.convert_csv_button,
            self.convert_document_to_md_button,
            self.convert_spreadsheet_to_md_button,
            self.convert_image_to_md_button,
            self.convert_layout_to_md_button,
        ]
        
        # 遍历所有按钮，启用存在的按钮
        enabled_count = 0
        for button in buttons:
            if button is not None:
                button.config(state="normal")
                enabled_count += 1
        
        logger.debug(f"已启用 {enabled_count} 个按钮")
    
    def refresh_style(self, theme_name: str = None):
        """
        刷新组件样式
        
        在主题切换后调用，更新所有子组件的样式以匹配新主题。
        
        参数:
            theme_name: 新主题名称（当前未使用）
        """
        self.main_frame.configure(bootstyle="default")
        self.status_label.configure(bootstyle="secondary")
        self.button_frame.configure(bootstyle="default")
        if self.status_var.get().startswith("错误:"):
            self.status_label.configure(bootstyle="danger")
