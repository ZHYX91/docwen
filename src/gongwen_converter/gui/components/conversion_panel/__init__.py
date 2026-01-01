"""
格式转换面板组件包

提供 ConversionPanel 类，用于显示格式转换按钮和相关选项。

支持的文件类别：
- 文档类（document）: DOCX/DOC/ODT/RTF → 互转 + PDF
- 表格类（spreadsheet）: XLSX/XLS/ODS/CSV → 互转 + PDF
- 图片类（image）: PNG/JPG/BMP/GIF/TIF/WebP → 互转 + PDF + TIFF合并
- 版式类（layout）: PDF/OFD/XPS → DOCX/DOC + TIF/JPG + 合并拆分

使用示例:
    from gongwen_converter.gui.components.conversion_panel import ConversionPanel
    
    panel = ConversionPanel(
        parent,
        config_manager=config_manager,
        on_action=handle_action
    )
    panel.set_file_info('document', 'docx', file_path)

模块结构:
    - base.py: ConversionPanelBase 基类，定义共享属性和公共方法
    - document_section.py: 文档类转换功能（校对）
    - spreadsheet_section.py: 表格类转换功能（汇总）
    - image_section.py: 图片类转换功能（压缩、合并）
    - layout_section.py: 版式类转换功能（合并拆分）
    - decoration.py: 装饰性绘图功能
"""

from .base import ConversionPanelBase
from .document_section import DocumentSectionMixin
from .spreadsheet_section import SpreadsheetSectionMixin
from .image_section import ImageSectionMixin
from .layout_section import LayoutSectionMixin
from .decoration import DecorationMixin


class ConversionPanel(
    DecorationMixin,
    LayoutSectionMixin,
    ImageSectionMixin,
    SpreadsheetSectionMixin,
    DocumentSectionMixin,
    ConversionPanelBase
):
    """
    格式转换面板组件
    
    通过多继承组合各功能模块，提供完整的格式转换面板功能。
    
    MRO（方法解析顺序）：
    ConversionPanel -> DecorationMixin -> LayoutSectionMixin 
    -> ImageSectionMixin -> SpreadsheetSectionMixin 
    -> DocumentSectionMixin -> ConversionPanelBase -> Frame -> object
    
    属性（继承自 ConversionPanelBase）：
        config_manager: 配置管理器实例
        on_action: 操作回调函数
        current_category: 当前文件类别
        current_format: 当前文件格式
        current_file_path: 当前文件路径
        format_buttons: 格式按钮字典
    
    公共方法：
        set_file_info(): 设置文件信息并更新显示
        show(): 显示面板
        hide(): 隐藏面板
        reset(): 重置面板状态
        refresh_decoration(): 刷新装饰区域（主题切换后调用）
        set_reference_table(): 设置基准表格（表格类）
        set_pdf_info(): 设置PDF信息（版式类）
    """
    pass


__all__ = ['ConversionPanel']
