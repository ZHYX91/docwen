"""
MD转表格功能模块

提供 Markdown 文件转换为表格格式的按钮：
- 转换按钮：XLSX、XLS、ODS、CSV

依赖：
- ActionPanelBase: 提供按钮样式、颜色映射等公共属性
- config_manager: 读取选项默认值

使用方式：
    此模块作为 Mixin 类被 ActionPanel 继承，不应直接实例化。
"""

import logging

import ttkbootstrap as tb

from gongwen_converter.utils.dpi_utils import scale
from gongwen_converter.utils.gui_utils import ToolTip

logger = logging.getLogger(__name__)


class MdToSpreadsheetMixin:
    """
    MD转表格功能混入类
    
    提供 Markdown 文件转换为表格格式的所有功能：
    - setup_for_md_to_spreadsheet: 设置MD转表格模式
    - _create_md_to_spreadsheet_buttons: 创建转换按钮
    - 点击事件处理方法
    
    依赖基类属性：
        button_container: 按钮容器框架
        button_colors: 按钮颜色映射
        button_style_2: 按钮样式配置
        on_action: 操作回调函数
    """
    
    def setup_for_md_to_spreadsheet(self, file_path: str):
        """
        设置为MD转表格处理模式
        
        为MD文件显示表格转换按钮（XLSX/XLS/ODS/CSV）。
        
        参数：
            file_path: MD文件路径
        """
        logger.debug(f"设置MD转表格处理模式: {file_path}")
        self.file_type = "xlsx"
        self.file_path = file_path
        self._clear_buttons()
        self._clear_options()
        self._create_md_to_spreadsheet_buttons()
        self.status_var.set("准备处理文本文件 - 转为表格")
        logger.info("MD转表格操作面板设置完成")
    
    def _create_md_to_spreadsheet_buttons(self):
        """
        创建MD转电子表格格式的按钮
        
        显示两行按钮：
        - 第一行：生成XLSX（主色调）、生成XLS（蓝色调）
        - 第二行：生成ODS（绿色调）、生成CSV（橙色调）
        """
        logger.debug("创建MD到电子表格系列的按钮 - 两行布局")
        
        # 第一行：生成XLSX | 生成XLS
        first_row_frame = tb.Frame(self.button_container, bootstyle="default")
        first_row_frame.grid(row=0, column=0, pady=(0, scale(10)))
        
        self.convert_excel_button = tb.Button(
            first_row_frame,
            text="📊 生成 XLSX",
            command=self._on_convert_excel_clicked,
            bootstyle=self.button_colors['primary'],
            **self.button_style_2
        )
        self.convert_excel_button.grid(row=0, column=0, padx=(0, scale(25)))
        ToolTip(self.convert_excel_button, "转换为Excel/WPS表格格式（推荐）")
        
        self.convert_xls_button = tb.Button(
            first_row_frame,
            text="📊 生成 XLS",
            command=self._on_convert_xls_clicked,
            bootstyle=self.button_colors['info'],
            **self.button_style_2
        )
        self.convert_xls_button.grid(row=0, column=1)
        ToolTip(
            self.convert_xls_button,
            "需要通过本地安装的 WPS、Microsoft Office 或 LibreOffice 进行转换，"
            "用户可自行设置使用软件的优先级。"
        )
        
        # 第二行：生成ODS | 生成CSV
        second_row_frame = tb.Frame(self.button_container, bootstyle="default")
        second_row_frame.grid(row=1, column=0, pady=(0, scale(10)))
        
        self.convert_ods_button = tb.Button(
            second_row_frame,
            text="📊 生成 ODS",
            command=self._on_convert_ods_clicked,
            bootstyle=self.button_colors['success'],
            **self.button_style_2
        )
        self.convert_ods_button.grid(row=0, column=0, padx=(0, scale(25)))
        ToolTip(
            self.convert_ods_button,
            "需要通过本地安装的 Microsoft Office 或 LibreOffice 进行转换，"
            "用户可自行设置使用软件的优先级。"
        )
        
        self.convert_csv_button = tb.Button(
            second_row_frame,
            text="📊 生成 CSV",
            command=self._on_convert_csv_clicked,
            bootstyle=self.button_colors['warning'],
            **self.button_style_2
        )
        self.convert_csv_button.grid(row=0, column=1)
        ToolTip(self.convert_csv_button, "转换为纯文本表格格式，可能需要本地Office软件预处理")
        
        logger.debug("MD到电子表格系列按钮创建完成 - 两行布局")
    
    def _on_convert_excel_clicked(self):
        """处理生成XLSX按钮点击事件"""
        if self.on_action:
            self.on_action("convert_md_to_xlsx", self.file_path, self.get_selected_options())
    
    def _on_convert_xls_clicked(self):
        """处理生成XLS按钮点击事件"""
        if self.on_action:
            self.on_action("convert_md_to_xls", self.file_path, {})
    
    def _on_convert_ods_clicked(self):
        """处理生成ODS按钮点击事件"""
        if self.on_action:
            self.on_action("convert_md_to_ods", self.file_path, {})
    
    def _on_convert_csv_clicked(self):
        """处理生成CSV按钮点击事件"""
        if self.on_action:
            self.on_action("convert_md_to_csv", self.file_path, {})
