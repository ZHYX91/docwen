"""
公共组件包

提供可重用的UI组件工厂和辅助类：
- ButtonFactory: 按钮创建工厂，统一管理按钮样式和颜色
- ExportOptionHandler: 导出选项联动处理器，处理"提取图片"和"OCR"选项的联动逻辑

使用示例:
    from docwen.gui.components.common import ButtonFactory, ExportOptionHandler

    # 创建格式转换按钮
    button = ButtonFactory.create_format_button(parent, 'DOCX', command)

    # 创建选项联动处理器
    handler = ExportOptionHandler(image_var, ocr_var)
"""

from .button_factory import ButtonFactory
from .option_handler import ExportOptionHandler, NumberingOptionHandler

__all__ = ["ButtonFactory", "ExportOptionHandler", "NumberingOptionHandler"]
