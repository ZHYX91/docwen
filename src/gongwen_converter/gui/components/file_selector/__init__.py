"""
文件选择器组件子包

提供批量文件管理的核心组件：
- FileSelector: 批量文件列表组件（单类别）
- FileInfo: 文件信息封装类
- TabbedFileSelector: 选项卡式批量文件列表（多类别）

模块结构：
- file_info.py: FileInfo 类，封装单个文件的信息
- core.py: FileSelectorCore 类，提供核心文件管理功能
- drag_drop_mixin.py: DragDropMixin 类，提供拖拽支持
- dialogs.py: 对话框辅助函数
- tabbed.py: TabbedFileSelector 类，选项卡式多类别管理

使用示例:
    from gongwen_converter.gui.components.file_selector import FileSelector, FileInfo, TabbedFileSelector
    
    # 使用单类别文件选择器
    selector = FileSelector(
        parent,
        on_file_selected=on_select_callback,
        on_file_removed=on_remove_callback
    )
    selector.add_file("/path/to/file.docx")
    
    # 使用多类别选项卡式选择器
    tabbed_selector = TabbedFileSelector(
        parent,
        on_tab_changed=on_tab_callback,
        on_file_selected=on_select_callback
    )
    tabbed_selector.add_files(["/path/to/file1.docx", "/path/to/file2.xlsx"])
"""

import logging

from .file_info import FileInfo
from .core import FileSelectorCore
from .drag_drop_mixin import DragDropMixin

logger = logging.getLogger(__name__)


class FileSelector(DragDropMixin, FileSelectorCore):
    """
    批量文件列表组件（组合类）
    
    通过多继承组合 FileSelectorCore 和 DragDropMixin，
    提供完整的文件选择器功能，包括拖拽支持。
    
    继承顺序（MRO）：
    FileSelector -> DragDropMixin -> FileSelectorCore -> object
    
    DragDropMixin 在前，其 _setup_drag_drop 方法会覆盖 FileSelectorCore 的占位方法。
    
    功能特性：
    - 滚动文件列表显示
    - 文件添加/移除/选中
    - 同类文件验证（基于实际文件格式）
    - 拖拽支持（需要 tkinterdnd2）
    - 状态管理和更新
    - 键盘导航（上下箭头调整顺序）
    
    使用示例:
        selector = FileSelector(
            parent,
            on_file_removed=on_remove,
            on_file_opened=on_open,
            on_list_cleared=on_clear,
            on_file_selected=on_select
        )
        
        # 添加文件
        success, error = selector.add_file("/path/to/file.docx", auto_select=True)
        
        # 批量添加
        added, failed = selector.add_files(["/path/to/file1.docx", "/path/to/file2.docx"])
        
        # 获取选中文件
        selected = selector.get_selected_file()
    """
    pass


# 延迟导入 TabbedFileSelector 以避免循环依赖
# TabbedFileSelector 内部会导入 FileSelector
from .tabbed import TabbedFileSelector


__all__ = [
    'FileSelector',
    'FileInfo',
    'TabbedFileSelector',
]

logger.debug("file_selector 子包初始化完成")
