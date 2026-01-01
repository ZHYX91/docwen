"""
状态栏组件模块
实现可滚动的消息历史记录显示，支持时间戳、颜色编码和定位按钮
"""

import logging
import tkinter as tk
from typing import Optional, Callable
from datetime import datetime
from collections import deque

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.utils.font_utils import get_default_font, get_small_font
from gongwen_converter.utils.icon_utils import load_image_icon
from gongwen_converter.utils.dpi_utils import scale

logger = logging.getLogger(__name__)

class StatusBar(tb.Frame):
    """
    可滚动的状态栏组件
    
    功能特性：
    - 支持最多100条消息历史记录（自动丢弃最早的消息）
    - 可滚动查看历史消息（始终显示滚动条）
    - 自动占满剩余垂直空间
    - 支持时间戳、颜色编码和定位按钮
    - 新消息添加时自动滚动到底部
    
    技术实现：
    - 使用 collections.deque(maxlen=100) 存储消息历史
    - 使用 Canvas + Scrollbar 实现可滚动容器
    - 动态创建 Label 显示每条消息
    """
    
    def __init__(self, master, on_location_clicked: Optional[Callable] = None, **kwargs):
        """
        初始化状态栏组件
        
        参数:
            master: 父组件
            on_location_clicked: 定位按钮点击回调函数，签名为 (file_path: str) -> None
            kwargs: 传递给tb.Frame的额外参数
        """
        super().__init__(master, **kwargs)
        logger.debug("初始化可滚动状态栏组件")
        
        # 存储配置
        self.on_location_clicked = on_location_clicked
        self.max_messages = 100  # 最大消息历史记录数
        
        # 获取字体配置
        self.default_font, self.default_size = get_default_font()
        self.small_font, self.small_size = get_small_font()
        
        # 初始化消息历史队列（使用deque自动限制长度）
        # 每个元素是一个字典: {"message": str, "type": str, "show_location": bool, "file_path": str}
        self.message_history = deque(maxlen=self.max_messages)
        logger.debug(f"消息队列初始化完成，最大容量: {self.max_messages}")
        
        # 存储最终输出文件路径（用于定位按钮的备用路径）
        self.final_output_path = None
        
        # 创建滚动容器和界面元素
        self._create_scrollable_container()
        
        logger.info("可滚动状态栏组件初始化完成")
    
    def _create_scrollable_container(self):
        """
        创建可滚动容器
        
        结构:
            StatusBar (self)
            ├── Canvas (self.canvas)
            │   └── Frame (self.scrollable_frame) - 实际的消息容器
            └── Scrollbar (self.scrollbar)
        
        滚动条始终显示，提供明确的视觉反馈
        """
        logger.debug("创建可滚动容器")
        
        # 配置主容器的网格权重
        self.grid_rowconfigure(0, weight=1)  # 让Canvas占满垂直空间
        self.grid_columnconfigure(0, weight=1)  # Canvas列可扩展
        self.grid_columnconfigure(1, weight=0)  # Scrollbar列固定宽度
        
        # 创建Canvas（用于滚动）
        self.canvas = tb.Canvas(self, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        logger.debug("Canvas创建完成")
        
        # 自定义滚动命令（带边界检查，防止滚动超出内容区域）
        def bounded_yview(*args):
            """带边界检查的滚动命令"""
            if args[0] == 'scroll':
                # 点击滚动条按钮时的滚动操作：检查是否已到边界
                top, bottom = self.canvas.yview()
                direction = int(args[1])
                if top <= 0 and direction < 0:  # 已在顶部，不能再向上滚动
                    return
                if bottom >= 1 and direction > 0:  # 已在底部，不能再向下滚动
                    return
            self.canvas.yview(*args)
        
        # 创建垂直滚动条（始终显示，使用带边界检查的命令）
        self.scrollbar = tb.Scrollbar(
            self, 
            orient="vertical", 
            command=bounded_yview
        )
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        logger.debug("滚动条创建完成（始终显示）")
        
        # 创建滚动框架（实际的消息容器）
        self.scrollable_frame = tb.Frame(self.canvas)
        self.scrollable_frame.grid_columnconfigure(0, weight=1)  # 消息列可扩展
        logger.debug("滚动框架创建完成")
        
        # 将滚动框架放入Canvas
        self.canvas_frame = self.canvas.create_window(
            (0, 0), 
            window=self.scrollable_frame, 
            anchor="nw"
        )
        
        # 配置Canvas的滚动命令
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 绑定事件：当滚动框架大小改变时更新滚动区域
        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        
        # 绑定事件：当Canvas大小改变时调整滚动框架宽度
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # 绑定鼠标滚轮事件
        self._bind_mousewheel_events()
        
        logger.debug("可滚动容器创建完成")
    
    def _on_frame_configure(self, event=None):
        """
        滚动框架配置改变事件处理
        更新Canvas的滚动区域以匹配框架大小
        """
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _on_canvas_configure(self, event):
        """
        Canvas配置改变事件处理
        调整滚动框架宽度以匹配Canvas宽度
        """
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_frame, width=canvas_width)
    
    def _bind_mousewheel_events(self):
        """
        绑定鼠标滚轮事件到所有相关组件
        支持Windows和Linux系统
        """
        def on_mousewheel(event):
            """处理Windows系统的鼠标滚轮事件"""
            # 检查是否已在边界（避免过度滚动）
            yview = self.canvas.yview()
            delta = int(-1 * (event.delta / 120))
            
            if delta < 0 and yview[0] <= 0:
                return  # 已在顶部
            elif delta > 0 and yview[1] >= 1:
                return  # 已在底部
            
            self.canvas.yview_scroll(delta, "units")
        
        def on_linux_scroll_up(event):
            """处理Linux系统向上滚动"""
            if self.canvas.yview()[0] <= 0:
                return  # 已在顶部
            self.canvas.yview_scroll(-1, "units")
        
        def on_linux_scroll_down(event):
            """处理Linux系统向下滚动"""
            if self.canvas.yview()[1] >= 1:
                return  # 已在底部
            self.canvas.yview_scroll(1, "units")
        
        # 绑定到主组件
        self.bind("<MouseWheel>", on_mousewheel)
        self.bind("<Button-4>", on_linux_scroll_up)
        self.bind("<Button-5>", on_linux_scroll_down)
        
        # 绑定到Canvas
        self.canvas.bind("<MouseWheel>", on_mousewheel)
        self.canvas.bind("<Button-4>", on_linux_scroll_up)
        self.canvas.bind("<Button-5>", on_linux_scroll_down)
        
        # 绑定到滚动框架
        self.scrollable_frame.bind("<MouseWheel>", on_mousewheel)
        self.scrollable_frame.bind("<Button-4>", on_linux_scroll_up)
        self.scrollable_frame.bind("<Button-5>", on_linux_scroll_down)
        
        logger.debug("鼠标滚轮事件绑定完成")
    
    def add_message(self, message: str, message_type: str = "secondary", 
                   show_location: bool = False, file_path: Optional[str] = None):
        """
        添加状态消息到历史记录
        
        新消息会添加到列表底部，并自动滚动到最新消息。
        如果消息数量超过100条，最早的消息会被自动丢弃。
        
        参数:
            message: 状态消息文本
            message_type: 消息类型/颜色 ('success', 'danger', 'warning', 'secondary', 'info')
            show_location: 是否显示定位按钮
            file_path: 关联的文件路径（用于定位按钮）
        """
        logger.debug(f"添加状态消息: {message} (类型: {message_type}, 定位: {show_location})")
        
        try:
            # 获取当前时间并格式化为"HH:MM:SS"
            current_time = datetime.now().strftime("%H:%M:%S")
            timestamped_message = f"{current_time} {message}"
            
            # 添加到消息历史队列（deque会自动丢弃最早的消息）
            message_data = {
                "message": timestamped_message,
                "type": message_type,
                "show_location": show_location,
                "file_path": file_path
            }
            self.message_history.append(message_data)
            logger.debug(f"消息已添加到历史队列，当前队列长度: {len(self.message_history)}")
            
            # 重新渲染所有消息
            self._render_all_messages()
            
            # 自动滚动到底部（显示最新消息）
            self.update_idletasks()  # 确保布局更新完成
            self.canvas.yview_moveto(1.0)  # 滚动到底部
            logger.debug("已自动滚动到最新消息")
            
        except Exception as e:
            logger.error(f"添加状态消息失败: {str(e)}", exc_info=True)
    
    def _render_all_messages(self):
        """
        重新渲染所有消息
        
        清空当前显示，然后根据消息历史队列重新创建所有Label和按钮
        """
        logger.debug("开始重新渲染所有消息")
        
        # 清空滚动框架中的所有子组件
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        # 为每条消息创建一行（Label + 可选的定位按钮）
        for row_index, msg_data in enumerate(self.message_history):
            self._create_message_row(row_index, msg_data)
        
        logger.debug(f"消息渲染完成，共 {len(self.message_history)} 条")
    
    def _create_message_row(self, row_index: int, msg_data: dict):
        """
        创建单条消息的显示行
        
        参数:
            row_index: 行索引（用于grid布局）
            msg_data: 消息数据字典
        """
        # 创建消息Label
        message_label = tb.Label(
            self.scrollable_frame,
            text=msg_data["message"],
            font=(self.small_font, self.small_size),
            bootstyle=msg_data["type"],
            anchor="w",
            padding=(5, 2)
        )
        message_label.grid(row=row_index, column=0, sticky="ew", padx=(0, 5))
        
        # 如果需要显示定位按钮
        if msg_data["show_location"]:
            # 加载定位图标
            icon_size = scale(16)
            location_icon = load_image_icon("location_icon.png", master=self.scrollable_frame, size=(icon_size, icon_size))
            
            # 创建定位按钮
            location_button = tb.Button(
                self.scrollable_frame,
                image=location_icon,
                bootstyle="secondary-link",
                command=lambda fp=msg_data["file_path"]: self._on_location_button_clicked(fp)
            )
            
            if location_icon:
                location_button.image = location_icon  # 保持对图片的引用
            
            location_button.grid(row=row_index, column=1, sticky="e", padx=(0, 5))
            logger.debug(f"第 {row_index} 行已添加定位按钮")
    
    def clear_all(self):
        """
        清空所有消息历史记录
        """
        logger.debug("清空所有消息历史")
        
        try:
            # 清空消息队列
            self.message_history.clear()
            
            # 清空显示
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            
            logger.info("状态栏已清空")
            
        except Exception as e:
            logger.error(f"清空状态栏失败: {str(e)}", exc_info=True)
    
    def set_final_output_path(self, file_path: str):
        """
        设置最终输出文件路径（用于定位按钮的备用路径）
        
        参数:
            file_path: 最终输出文件路径
        """
        logger.debug(f"设置最终输出文件路径: {file_path}")
        self.final_output_path = file_path
    
    def _on_location_button_clicked(self, file_path: Optional[str]):
        """
        处理定位按钮点击事件
        
        参数:
            file_path: 要定位的文件路径
        """
        logger.debug(f"定位按钮被点击，文件路径: {file_path}")
        
        # 如果没有指定路径，使用最终输出路径
        if not file_path and hasattr(self, 'final_output_path'):
            file_path = self.final_output_path
        
        # 调用回调函数
        if file_path and self.on_location_clicked:
            logger.info(f"执行定位操作: {file_path}")
            self.on_location_clicked(file_path)
        else:
            logger.warning("无法定位文件：文件路径不存在或回调函数未设置")
    
    def refresh_style(self):
        """
        刷新组件样式以匹配当前主题
        主题切换时调用此方法
        """
        logger.debug("刷新状态栏组件样式")
        
        # 重新渲染所有消息以应用新主题
        self._render_all_messages()
        
        logger.debug("状态栏样式刷新完成")
