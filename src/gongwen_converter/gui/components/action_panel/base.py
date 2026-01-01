"""
操作面板基类模块

提供 ActionPanel 组件的基础框架，包括：
- 组件初始化和布局管理
- 公共属性和样式配置
- 按钮/选项区域的清理方法
- 显示/隐藏/状态管理方法

依赖：
- ttkbootstrap: 提供现代化UI组件
- config_manager: 读取配置选项

使用方式：
    此模块作为 Mixin 基类被 ActionPanel 继承，不应直接实例化。
"""

import logging
import tkinter as tk
from typing import Callable, Optional, Dict

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.utils.font_utils import get_default_font, get_title_font, get_small_font
from gongwen_converter.utils.dpi_utils import scale

logger = logging.getLogger(__name__)


class ActionPanelBase(tb.Frame):
    """
    操作面板基类
    
    提供操作面板的基础框架和公共功能：
    - 组件初始化和grid布局管理
    - 字体、按钮样式、颜色映射等公共属性
    - 按钮和选项区域的清理方法
    - 显示/隐藏、取消按钮管理
    - 状态消息显示
    
    属性：
        config_manager: 配置管理器实例
        on_action: 操作按钮点击回调函数
        on_cancel: 取消按钮点击回调函数
        file_type: 当前文件类型
        file_path: 当前文件路径
        checkbox_vars: 校对选项变量字典
    """
    
    def __init__(
        self,
        master,
        config_manager,
        on_action: Optional[Callable] = None,
        on_cancel: Optional[Callable] = None,
        **kwargs
    ):
        """
        初始化操作面板基类
        
        参数：
            master: 父组件对象
            config_manager: 配置管理器实例，用于读取和保存选项配置
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
        
        # 导出选项处理器（由子类创建）
        self._doc_export_handler = None
        self._table_export_handler = None
        self._image_export_handler = None
        self._pdf_export_handler = None
        
        # 标题序号选项变量（文档转MD）
        self.doc_remove_numbering_var: Optional[tk.BooleanVar] = None
        self.doc_add_numbering_var: Optional[tk.BooleanVar] = None
        self.doc_numbering_scheme_var: Optional[tk.StringVar] = None
        self.doc_numbering_scheme_combo: Optional[tb.Combobox] = None
        
        # 标题序号选项变量（MD转文档）
        self.md_remove_numbering_var: Optional[tk.BooleanVar] = None
        self.md_add_numbering_var: Optional[tk.BooleanVar] = None
        self.md_numbering_scheme_var: Optional[tk.StringVar] = None
        self.md_numbering_scheme_combo: Optional[tb.Combobox] = None
        
        # 存储按钮引用
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
        
        # 获取字体配置
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
            'primary': 'primary',      # 主题色
            'secondary': 'secondary',  # 灰色
            'success': 'success',      # 绿色
            'info': 'info',            # 蓝色
            'warning': 'warning',      # 橙色
            'danger': 'danger'         # 红色
        }
        
        # 配置grid布局权重
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
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
        """
        logger.debug("创建操作面板界面元素 - 使用grid布局")

        # 创建主框架
        self.main_frame = tb.Frame(self, bootstyle="default")
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=scale(5))
        
        # 配置主框架网格权重
        self.main_frame.grid_rowconfigure(0, weight=0)  # 状态区域
        self.main_frame.grid_rowconfigure(1, weight=0)  # 按钮区域
        self.main_frame.grid_columnconfigure(0, weight=1)

        # 创建状态标签
        self.status_var = tk.StringVar(value="")
        self.status_label = tb.Label(
            self.main_frame,
            textvariable=self.status_var,
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            wraplength=scale(400),
            anchor=tk.CENTER
        )
        self.status_label.grid(row=0, column=0, sticky="ew", pady=(0, scale(10)))

        # 创建按钮框架
        self.button_frame = tb.Frame(self.main_frame, bootstyle="default")
        self.button_frame.grid(row=1, column=0, sticky="ew", pady=(0, scale(5)))

        # 配置按钮框架网格权重
        self.button_frame.grid_rowconfigure(0, weight=0)
        self.button_frame.grid_rowconfigure(1, weight=0)
        self.button_frame.grid_columnconfigure(0, weight=1)

        # 创建按钮容器框架
        self.button_container = tb.Frame(self.button_frame, bootstyle="default")
        self.button_container.grid(row=0, column=0, sticky="ew", pady=(0, scale(5)))
        self.button_container.grid_columnconfigure(0, weight=1)
        
        # 创建取消按钮容器（默认隐藏）
        self.cancel_button_container = tb.Frame(self.button_frame, bootstyle="default")
        self.cancel_button_container.grid(row=0, column=0, sticky="ew")
        self.cancel_button_container.grid_remove()

        self.cancel_button = tb.Button(
            self.cancel_button_container,
            text="❌ 取消",
            command=self._on_cancel_clicked,
            bootstyle=self.button_colors['danger'],
            **self.button_style_1
        )
        self.cancel_button.pack()

        # 配置按钮容器框架网格权重
        self.button_frame.grid_columnconfigure(0, weight=1)
        self.button_frame.grid_rowconfigure(0, weight=1)
        
        logger.debug("操作面板界面元素创建完成")
    
    def _on_cancel_clicked(self):
        """
        处理取消按钮点击事件
        
        禁用取消按钮以防止重复点击，然后调用取消回调函数中止当前操作。
        """
        logger.info("取消按钮被点击")
        self.cancel_button.config(state="disabled")
        if self.on_cancel:
            self.on_cancel()
    
    def _clear_buttons(self):
        """
        清空按钮区域
        
        销毁按钮容器中的所有子组件，并重置按钮引用为 None。
        """
        for widget in self.button_container.winfo_children():
            widget.destroy()
        
        self.convert_docx_button = None
        self.convert_doc_button = None
        self.convert_odt_button = None
        self.convert_rtf_button = None
        self.convert_excel_button = None
        self.convert_xls_button = None
        self.convert_ods_button = None
        self.convert_csv_button = None
        self.convert_document_to_md_button = None
        self.convert_spreadsheet_to_md_button = None
        self.convert_image_to_md_button = None
        self.convert_layout_to_md_button = None
    
    def _clear_options(self):
        """
        清空选项区域
        
        重置校对选项变量字典和导出选项处理器。
        """
        self.checkbox_vars = {}
        self._doc_export_handler = None
        self._table_export_handler = None
        self._image_export_handler = None
        self._pdf_export_handler = None
    
    def show(self):
        """显示操作面板"""
        self.grid(row=0, column=0, sticky="nsew", padx=scale(1), pady=scale(5))
    
    def hide(self):
        """隐藏操作面板"""
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
        
        参数：
            message: 要显示的消息文本
            is_error: 是否为错误消息，True时显示为红色
        """
        bootstyle = "danger" if is_error else "secondary"
        self.status_var.set(f"错误: {message}" if is_error else message)
        self.status_label.configure(bootstyle=bootstyle)
    
    def clear_status(self):
        """清除状态消息"""
        self.set_status("")
    
    def enable_all_buttons(self):
        """
        启用所有操作按钮
        
        在操作完成或取消后调用此方法，恢复所有按钮的可用状态。
        """
        logger.debug("启用所有操作按钮")
        
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
        
        参数：
            theme_name: 新主题名称（当前未使用）
        """
        self.main_frame.configure(bootstyle="default")
        self.status_label.configure(bootstyle="secondary")
        self.button_frame.configure(bootstyle="default")
        
        if self.status_var.get().startswith("错误:"):
            self.status_label.configure(bootstyle="danger")
    
    def get_selected_options(self) -> Dict[str, any]:
        """
        获取当前选中的选项
        
        返回包含校对选项和序号配置的字典。
        
        返回：
            Dict[str, any]: 包含所有选项的字典
        """
        # 基础校对选项
        options = {key: var.get() for key, var in self.checkbox_vars.items()}
        
        # 方案名称到ID的映射
        scheme_name_to_id = {
            "公文标准": "gongwen_standard",
            "层级数字标准": "hierarchical_standard",
            "法律条文标准": "legal_standard"
        }
        
        # 添加文档转MD的序号参数
        if hasattr(self, 'doc_remove_numbering_var') and self.doc_remove_numbering_var is not None:
            options['doc_remove_numbering'] = self.doc_remove_numbering_var.get()
            options['doc_add_numbering'] = self.doc_add_numbering_var.get()
            if hasattr(self, 'doc_numbering_scheme_var') and self.doc_numbering_scheme_var is not None:
                scheme_name = self.doc_numbering_scheme_var.get()
                options['doc_numbering_scheme'] = scheme_name_to_id.get(scheme_name, "gongwen_standard")
        
        # 添加MD转文档的序号参数
        if hasattr(self, 'md_remove_numbering_var') and self.md_remove_numbering_var is not None:
            options['md_remove_numbering'] = self.md_remove_numbering_var.get()
            options['md_add_numbering'] = self.md_add_numbering_var.get()
            if hasattr(self, 'md_numbering_scheme_var') and self.md_numbering_scheme_var is not None:
                scheme_name = self.md_numbering_scheme_var.get()
                options['md_numbering_scheme'] = scheme_name_to_id.get(scheme_name, "gongwen_standard")
        
        return options
