"""
按钮创建工厂模块

提供统一的按钮创建方法，管理按钮样式和颜色映射。

主要功能：
- 创建格式转换按钮（支持不同列数的样式）
- 创建操作按钮（带图标）
- 统一管理按钮颜色映射

依赖：
- ttkbootstrap: 提供现代化的按钮样式
- dpi_utils: 支持DPI缩放

使用示例：
    from gongwen_converter.gui.components.common import ButtonFactory
    
    # 创建格式转换按钮
    button = ButtonFactory.create_format_button(
        parent, 'DOCX', command, style_size='2col'
    )
    
    # 创建操作按钮
    button = ButtonFactory.create_action_button(
        parent, '📝 生成 DOCX', command, style='primary'
    )
"""

import logging
from typing import Callable, Dict, Optional, Any

import ttkbootstrap as tb
from gongwen_converter.utils.dpi_utils import scale

logger = logging.getLogger(__name__)


class ButtonFactory:
    """
    按钮创建工厂
    
    统一管理按钮的样式和颜色，提供创建格式转换按钮和操作按钮的方法。
    
    属性：
        COLORS: 格式到颜色的映射字典
        STYLES: 按钮样式配置字典
    """
    
    # 按钮颜色映射（格式名称 → ttkbootstrap样式）
    COLORS: Dict[str, str] = {
        # 文档类
        'DOCX': 'primary',    # 主题色：主要格式
        'DOC': 'info',        # 蓝色：旧版格式
        'ODT': 'success',     # 绿色：开放格式
        'RTF': 'warning',     # 橙色：富文本格式
        # 表格类
        'XLSX': 'primary',    # 主题色：主要格式
        'XLS': 'info',        # 蓝色：旧版格式
        'ET': 'info',         # 蓝色：WPS格式
        'ODS': 'success',     # 绿色：开放格式
        'CSV': 'warning',     # 橙色：通用格式
        # 图片类
        'PNG': 'primary',     # 主题色：常用格式
        'JPG': 'primary',     # 主题色：常用格式
        'JPEG': 'primary',    # 主题色：常用格式
        'BMP': 'info',        # 蓝色：基础格式
        'GIF': 'success',     # 绿色：动画格式
        'TIF': 'warning',     # 橙色：专业格式
        'TIFF': 'warning',    # 橙色：专业格式
        'WebP': 'danger',     # 红色：现代格式
        # 版式类
        'PDF': 'danger',      # 红色：PDF格式
        'XPS': 'info',        # 蓝色：版式格式
        'OFD': 'success',     # 绿色：国产格式
        'CEB': 'warning',     # 橙色：电子书格式
    }
    
    # 按钮样式配置（支持DPI缩放）
    # 注意：这些值在运行时会根据DPI进行缩放
    STYLES: Dict[str, Dict[str, Any]] = {
        '3col': {
            'width': 8,
            'padding_x': 8,
            'padding_y': 4
        },
        '2col': {
            'width': 10,
            'padding_x': 8,
            'padding_y': 4
        },
        '1col': {
            'width': 16,
            'padding_x': 8,
            'padding_y': 4
        },
        'action': {
            'width': 20,
            'padding_x': 10,
            'padding_y': 5
        },
        'action_medium': {
            'width': 16,
            'padding_x': 10,
            'padding_y': 5
        },
        'action_small': {
            'width': 12,
            'padding_x': 10,
            'padding_y': 5
        }
    }
    
    @classmethod
    def get_color(cls, fmt: str) -> str:
        """
        获取指定格式的按钮颜色
        
        参数：
            fmt: 格式名称（不区分大小写）
            
        返回：
            str: ttkbootstrap样式名称，如果未找到则返回'primary'
        """
        return cls.COLORS.get(fmt.upper(), 'primary')
    
    @classmethod
    def get_style(cls, style_size: str) -> Dict[str, Any]:
        """
        获取指定尺寸的按钮样式配置
        
        参数：
            style_size: 样式尺寸名称 ('3col', '2col', '1col', 'action', 等)
            
        返回：
            Dict: 包含width和padding的样式配置字典
        """
        style_config = cls.STYLES.get(style_size, cls.STYLES['2col'])
        return {
            'width': style_config['width'],
            'padding': (scale(style_config['padding_x']), scale(style_config['padding_y']))
        }
    
    @classmethod
    def create_format_button(
        cls,
        parent: tb.Frame,
        fmt: str,
        command: Callable,
        style_size: str = '2col',
        text: Optional[str] = None,
        bootstyle: Optional[str] = None,
        state: str = 'normal'
    ) -> tb.Button:
        """
        创建格式转换按钮
        
        根据格式名称自动选择颜色，支持自定义文本和样式。
        
        参数：
            parent: 父组件
            fmt: 格式名称（如'DOCX', 'PNG'等）
            command: 点击回调函数
            style_size: 按钮尺寸样式 ('3col', '2col', '1col')
            text: 自定义按钮文本，默认使用格式名称
            bootstyle: 自定义按钮颜色样式，默认根据格式自动选择
            state: 按钮状态 ('normal' 或 'disabled')
            
        返回：
            tb.Button: 创建的按钮对象
        """
        button_text = text if text else fmt
        button_style = bootstyle if bootstyle else cls.get_color(fmt)
        style_config = cls.get_style(style_size)
        
        button = tb.Button(
            parent,
            text=button_text,
            command=command,
            bootstyle=button_style,
            state=state,
            **style_config
        )
        
        logger.debug(f"创建格式按钮: {fmt} - 样式: {button_style}, 尺寸: {style_size}")
        return button
    
    @classmethod
    def create_action_button(
        cls,
        parent: tb.Frame,
        text: str,
        command: Callable,
        bootstyle: str = 'primary',
        style_size: str = 'action',
        state: str = 'normal'
    ) -> tb.Button:
        """
        创建操作按钮
        
        用于创建带图标的操作按钮，如"📝 生成 DOCX"。
        
        参数：
            parent: 父组件
            text: 按钮文本（可包含emoji图标）
            command: 点击回调函数
            bootstyle: 按钮颜色样式
            style_size: 按钮尺寸样式 ('action', 'action_medium', 'action_small')
            state: 按钮状态 ('normal' 或 'disabled')
            
        返回：
            tb.Button: 创建的按钮对象
        """
        style_config = cls.get_style(style_size)
        
        button = tb.Button(
            parent,
            text=text,
            command=command,
            bootstyle=bootstyle,
            state=state,
            **style_config
        )
        
        logger.debug(f"创建操作按钮: {text} - 样式: {bootstyle}, 尺寸: {style_size}")
        return button
    
    @classmethod
    def create_disabled_button(
        cls,
        parent: tb.Frame,
        text: str,
        style_size: str = '2col',
        tooltip_text: Optional[str] = None
    ) -> tb.Button:
        """
        创建禁用状态的按钮
        
        用于显示暂不支持的功能按钮。
        
        参数：
            parent: 父组件
            text: 按钮文本
            style_size: 按钮尺寸样式
            tooltip_text: 提示文本（如"暂不支持"）
            
        返回：
            tb.Button: 创建的禁用按钮对象
        """
        style_config = cls.get_style(style_size)
        
        button = tb.Button(
            parent,
            text=text,
            command=lambda: None,
            bootstyle='secondary',
            state='disabled',
            **style_config
        )
        
        # 如果提供了tooltip文本，添加提示
        if tooltip_text:
            from gongwen_converter.utils.gui_utils import ToolTip
            ToolTip(button, tooltip_text)
        
        logger.debug(f"创建禁用按钮: {text}")
        return button
