"""
操作面板组件子包

提供 ActionPanel 组件，用于根据文件类型动态显示不同的操作按钮和选项。

主要功能：
- MD文件：提供文档格式转换（DOCX/DOC/ODT/RTF）和表格格式转换（XLSX/XLS/ODS/CSV）
- 文档文件：提供导出Markdown功能
- 表格文件：提供导出Markdown功能
- 图片文件：提供OCR识别和导出Markdown功能
- 版式文件：提供导出Markdown功能

模块结构：
- base.py: ActionPanelBase 基类，提供基础框架和公共功能
- md_to_document.py: MdToDocumentMixin，MD转文档功能
- md_to_spreadsheet.py: MdToSpreadsheetMixin，MD转表格功能
- file_to_md.py: FileToMdMixin，各类文件转MD功能

使用方式：
    from gongwen_converter.gui.components.action_panel import ActionPanel
    
    panel = ActionPanel(master, config_manager, on_action=callback)
    panel.setup_for_md_to_document(file_path)
"""

from .base import ActionPanelBase
from .md_to_document import MdToDocumentMixin
from .md_to_spreadsheet import MdToSpreadsheetMixin
from .file_to_md import FileToMdMixin


class ActionPanel(
    FileToMdMixin,
    MdToSpreadsheetMixin,
    MdToDocumentMixin,
    ActionPanelBase
):
    """
    操作面板组件
    
    根据文件类型动态显示相应的转换按钮和选项。
    
    继承顺序说明（MRO）：
    ActionPanel -> FileToMdMixin -> MdToSpreadsheetMixin 
    -> MdToDocumentMixin -> ActionPanelBase -> Frame -> object
    
    Mixin 类按功能分组：
    - ActionPanelBase: 基础框架、公共属性、显示/隐藏管理
    - MdToDocumentMixin: MD转DOCX/DOC/ODT/RTF 按钮和校对选项
    - MdToSpreadsheetMixin: MD转XLSX/XLS/ODS/CSV 按钮
    - FileToMdMixin: 文档/表格/图片/版式 转MD 按钮和导出选项
    
    特性：
    - 使用ttkbootstrap统一管理样式
    - 按钮水平居中排列
    - 支持MD文件的多格式转换
    - 提供Markdown导出功能
    - 支持文档生成的校对选项
    - 使用grid布局管理器
    
    使用示例：
        # 创建面板
        panel = ActionPanel(
            master=parent,
            config_manager=config,
            on_action=self._on_action_callback,
            on_cancel=self._on_cancel_callback
        )
        
        # 根据文件类型设置模式
        panel.setup_for_md_to_document(md_file_path)
        # 或
        panel.setup_for_document_file(docx_file_path)
        
        # 显示面板
        panel.show()
    """
    pass


__all__ = ['ActionPanel']
