"""
GUI组件包 - 可重用的用户界面组件

提供各类可重用的UI组件，包括：
- 文件拖拽区域（支持单文件和批量模式）
- 模板选择器
- 操作面板（动作按钮和选项）
- 格式转换面板（支持4种文件类别）
- 关于对话框
- 批量文件列表
- 状态栏
- 分页文件管理器

所有组件基于ttkbootstrap构建，支持主题切换和DPI缩放。

子包结构：
- common/: 公共组件（ButtonFactory, ExportOptionHandler）
- conversion_panel/: 格式转换面板（拆分后的模块化结构）
"""

# 导出核心组件类，供外部模块直接导入使用
from .file_drop import FileDropArea
from .action_panel import ActionPanel
from .about_dialog import AboutDialog
from .conversion_panel import ConversionPanel

# 模块导出列表
__all__ = [
    'FileDropArea',
    'ActionPanel',
    'AboutDialog',
    'ConversionPanel'
]

# 包初始化日志
import logging
logger = logging.getLogger(__name__)
logger.info("GUI组件包初始化完成 - 所有UI组件已加载")
