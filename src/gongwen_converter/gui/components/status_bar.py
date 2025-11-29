"""
状态栏组件模块
实现4行滚动消息队列的状态栏，支持时间戳、颜色编码和定位按钮
"""

import logging
import tkinter as tk
from typing import List, Optional, Callable
from datetime import datetime

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.utils.font_utils import get_default_font, get_small_font
from gongwen_converter.utils.icon_utils import load_image_icon

logger = logging.getLogger(__name__)

class StatusBar(tb.Frame):
    """
    独立的状态栏组件
    管理4行滚动消息队列，支持时间戳、颜色编码和定位按钮
    """
    
    def __init__(self, master, line_count: int = 4, on_location_clicked: Optional[Callable] = None, **kwargs):
        """
        初始化状态栏组件
        
        参数:
            master: 父组件
            line_count: 消息行数 (默认4行)
            on_location_clicked: 定位按钮点击回调函数
            kwargs: 传递给tb.Frame的额外参数
        """
        super().__init__(master, **kwargs)
        logger.debug(f"初始化状态栏组件，行数: {line_count}")
        
        # 存储配置
        self.line_count = line_count
        self.on_location_clicked = on_location_clicked
        
        # 获取字体配置
        self.default_font, self.default_size = get_default_font()
        self.small_font, self.small_size = get_small_font()
        
        # 初始化消息队列
        self.message_queue = [""] * line_count
        self.color_queue = ["secondary"] * line_count
        self.location_queue = [False] * line_count
        self.file_paths = [None] * line_count  # 存储每行对应的文件路径
        
        # 创建界面元素
        self._create_widgets()
        
        logger.info("状态栏组件初始化完成")
    
    def _create_widgets(self):
        """创建状态栏界面元素"""
        logger.debug("创建状态栏界面元素")
        
        # 配置网格权重，确保状态栏可以垂直扩展
        for i in range(self.line_count):
            self.grid_rowconfigure(i, weight=1)
        self.grid_columnconfigure(0, weight=1)  # 消息标签列
        self.grid_columnconfigure(1, weight=0)  # 定位按钮列
        
        # 创建每行的消息标签和定位按钮
        self.message_vars = []
        self.message_labels = []
        self.location_buttons = []
        
        for i in range(self.line_count):
            # 创建消息变量和标签
            message_var = tk.StringVar(value=self.message_queue[i])
            message_label = tb.Label(
                self,
                textvariable=message_var,
                font=(self.small_font, self.small_size),
                bootstyle=self.color_queue[i],
                anchor="w",
                padding=(5, 2)
            )
            message_label.grid(row=i, column=0, sticky="ew", padx=(0, 5))
            
            # 加载并创建定位按钮（支持DPI缩放）
            from gongwen_converter.utils.dpi_utils import scale
            icon_size = scale(16)
            location_icon = load_image_icon("location_icon.png", master=self, size=(icon_size, icon_size))
            location_button = tb.Button(
                self,
                image=location_icon,
                bootstyle="secondary-link",  # 使用link样式以更好地显示图片
                command=lambda row=i: self._on_location_button_clicked(row)
            )
            if location_icon:
                location_button.image = location_icon  # 保持对图片的引用
            location_button.grid(row=i, column=1, sticky="e", padx=(0, 5))
            
            # 初始隐藏定位按钮
            location_button.grid_remove()
            
            # 存储组件引用
            self.message_vars.append(message_var)
            self.message_labels.append(message_label)
            self.location_buttons.append(location_button)
        
        logger.debug("状态栏界面元素创建完成")
    
    def add_message(self, message: str, message_type: str = "secondary", 
                   show_location: bool = False, file_path: Optional[str] = None):
        """
        添加状态消息
        
        参数:
            message: 状态消息
            message_type: 消息类型 ('success', 'danger', 'warning', 'secondary')
            show_location: 是否显示定位按钮
            file_path: 关联的文件路径（用于定位按钮）
        """
        logger.debug(f"添加状态消息: {message} (类型: {message_type}, 定位: {show_location})")
        
        try:
            # 获取当前时间并格式化为"HH:MM:SS"
            current_time = datetime.now().strftime("%H:%M:%S")
            timestamped_message = f"{current_time} {message}"
            
            # 滚动消息队列：移除最旧消息，添加最新消息
            self.message_queue.pop(0)
            self.message_queue.append(timestamped_message)
            
            self.color_queue.pop(0)
            self.color_queue.append(message_type)
            
            self.location_queue.pop(0)
            self.location_queue.append(show_location)
            
            self.file_paths.pop(0)
            self.file_paths.append(file_path)
            
            # 更新所有行的显示
            for i in range(self.line_count):
                self.message_vars[i].set(self.message_queue[i])
                self.message_labels[i].configure(bootstyle=self.color_queue[i])
                
                # 更新定位按钮显示状态
                if self.location_queue[i]:
                    self.location_buttons[i].grid()
                else:
                    self.location_buttons[i].grid_remove()
            
            logger.debug("状态消息添加完成")
            
        except Exception as e:
            logger.error(f"添加状态消息失败: {str(e)}")
    
    def clear_all(self):
        """清空状态栏所有信息"""
        logger.debug("清空状态栏所有信息")
        
        try:
            # 清空所有队列
            self.message_queue = [""] * self.line_count
            self.color_queue = ["secondary"] * self.line_count
            self.location_queue = [False] * self.line_count
            self.file_paths = [None] * self.line_count
            
            # 更新显示
            for i in range(self.line_count):
                self.message_vars[i].set("")
                self.message_labels[i].configure(bootstyle="secondary")
                self.location_buttons[i].grid_remove()
            
            logger.debug("状态栏已清空")
            
        except Exception as e:
            logger.error(f"清空状态栏失败: {str(e)}")
    
    def set_final_output_path(self, file_path: str):
        """
        设置最终输出文件路径（用于定位按钮）
        
        参数:
            file_path: 最终输出文件路径
        """
        logger.debug(f"设置最终输出文件路径: {file_path}")
        self.final_output_path = file_path
    
    def _on_location_button_clicked(self, row: int):
        """
        处理定位按钮点击事件
        
        参数:
            row: 按钮所在行索引
        """
        logger.debug(f"定位按钮被点击 - 行{row}")
        
        # 优先使用该行特定的文件路径
        file_path = self.file_paths[row]
        if not file_path and hasattr(self, 'final_output_path'):
            file_path = self.final_output_path
        
        if file_path and self.on_location_clicked:
            self.on_location_clicked(file_path)
        else:
            logger.warning("无法定位文件：文件路径不存在或回调函数未设置")
    
    def refresh_style(self):
        """刷新组件样式以匹配当前主题"""
        logger.debug("刷新状态栏组件样式")
        
        # 重新应用当前颜色
        for i in range(self.line_count):
            self.message_labels[i].configure(bootstyle=self.color_queue[i])
        
        # 重置定位按钮样式
        for button in self.location_buttons:
            button.configure(bootstyle="secondary")
