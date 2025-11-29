"""
批量文件列表组件 - 批量模式下的文件管理界面

提供完整的批量文件管理功能：
- 文件列表显示（序号、文件名、路径、大小）
- 文件操作按钮（移至顶部、打开位置、移除）
- 文件选中机制（单选，支持回调）
- 智能文件分类（根据实际格式，非扩展名）
- 格式验证和警告提示
- 完成状态显示（转换完成后显示图标）

特性：
- 支持同类文件批量添加（文本/文档/表格/图片/PDF）
- 多数决策算法（混合文件时自动选择主要类别）
- 实时状态更新（待处理/处理中/已完成/失败）
- 支持DPI缩放和主题切换
"""

import os
import logging
import subprocess
import tkinter as tk
from typing import List, Dict, Callable, Optional, Tuple
from gongwen_converter.utils.font_utils import get_default_font, get_small_font, get_micro_font
from gongwen_converter.utils.path_utils import get_file_size_formatted, collect_files_from_folder
from gongwen_converter.utils.icon_utils import load_image_icon
from gongwen_converter.utils.file_type_utils import (
    get_file_category, get_category_name, is_supported_file,
    get_file_info, get_actual_file_category, activate_optimal_tab
)

# 导入ttkbootstrap用于界面美化和样式管理
import ttkbootstrap as tb
from ttkbootstrap.constants import *

logger = logging.getLogger(__name__)

# 尝试导入拖拽支持库
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKINTERDND2_AVAILABLE = True
    logger.debug("tkinterdnd2库已成功导入")
except ImportError:
    TkinterDnD = None
    DND_FILES = None
    TKINTERDND2_AVAILABLE = False
    logger.warning("tkinterdnd2未安装，拖拽功能将不可用")


class FileInfo:
    """
    文件信息类
    
    封装单个文件的完整信息：
    - 基本信息：路径、文件名、大小
    - 格式信息：实际格式、类别、扩展名
    - 验证信息：是否支持、是否有效、警告消息
    - 状态信息：选中状态、处理状态、输出路径
    
    此类用于在批量文件列表中统一管理文件信息。
    """
    
    def __init__(self, file_path: str):
        """
        初始化文件信息
        
        读取文件的基本信息和实际格式，进行格式验证。
        
        参数:
            file_path: 文件的完整路径
        """
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.file_size = get_file_size_formatted(file_path)
        
        # 获取完整的文件信息
        file_info = get_file_info(file_path)
        self.actual_format = file_info['actual_format']
        self.actual_category = file_info['actual_category']
        self.extension = file_info['extension']
        self.extension_category = file_info['extension_category']
        self.is_supported = file_info['is_supported']
        self.is_valid = file_info['is_valid']
        self.warning_message = file_info['warning_message']
        
        # 选中状态
        self.is_selected = False
        
        # 处理状态
        self.status = 'pending'  # pending/processing/completed/skipped/failed
        self.output_path = None  # 输出文件路径
        
        # 新增：跳过和失败的详细信息
        self.skip_reason = None  # 跳过原因（如："已是XLS格式"）
        self.error_message = None  # 失败原因（详细错误信息）
        
        logger.debug(f"创建FileInfo: {self.file_path}, 实际类别: {self.actual_category}, 实际格式: {self.actual_format}")
    
    def to_dict(self) -> Dict:
        """
        转换为字典格式
        
        将文件信息转换为字典，便于传递和序列化。
        兼容现有代码中使用字典的地方。
        
        返回:
            Dict: 包含所有关键文件信息的字典
        """
        return {
            'path': self.file_path,
            'size': self.file_size,
            'status': self.status,
            'output_path': self.output_path,
            'actual_category': self.actual_category,
            'actual_format': self.actual_format,
            'is_selected': self.is_selected,
            'warning_message': self.warning_message
        }
    
    def __str__(self) -> str:
        return f"FileInfo({self.file_path}, 类别:{self.actual_category}, 选中:{self.is_selected})"


class FileSelector:
    """
    批量文件列表组件
    
    批量模式下的核心文件管理界面，提供：
    - 滚动文件列表显示
    - 文件添加和移除
    - 文件选中和回调
    - 同类文件验证
    - 状态管理和更新
    - 空状态提示
    
    支持单文件和批量文件两种模式，自动进行文件分类和验证。
    """
    
    def __init__(self, master, on_file_removed: Optional[Callable] = None, 
                 on_file_opened: Optional[Callable] = None,
                 on_list_cleared: Optional[Callable] = None,
                 on_file_selected: Optional[Callable] = None,
                 **kwargs):
        """
        初始化批量文件列表组件
        
        参数:
            master: 父组件对象
            on_file_removed: 文件移除时的回调函数(file_path)
            on_file_opened: 文件打开时的回调函数(file_path)
            on_list_cleared: 列表清空时的回调函数（用于重置UI状态）
            on_file_selected: 文件选中时的回调函数(file_info)
            **kwargs: 传递给Frame的其他参数
        """
        logger.debug("初始化批量文件列表组件")
        
        # 存储父组件
        self.master = master
        
        # 存储回调函数
        self.on_file_removed = on_file_removed
        self.on_file_opened = on_file_opened
        self.on_list_cleared = on_list_cleared
        self.on_file_selected = on_file_selected
        
        # 文件列表数据结构：使用FileInfo对象存储完整文件信息
        self.files: List[FileInfo] = []
        self.current_category: Optional[str] = None  # 当前文件类别：text/document/spreadsheet/image/layout
        self.selected_file: Optional[FileInfo] = None  # 当前选中的文件
        
        # UI引用管理：保存每个文件项的frame引用，用于增量更新
        self.file_item_frames: Dict[str, tk.Frame] = {}
        
        # 拖拽状态
        self.is_dragging = False
        self.drag_enabled = True
        
        # 获取字体配置
        self.default_font, self.default_size = get_default_font()
        self.small_font, self.small_size = get_small_font()
        self.micro_font, self.micro_size = get_micro_font()
        
        # 创建界面元素
        self._create_widgets()
        
        # 设置拖拽支持
        self._setup_drag_drop()
        
        logger.info("批量文件列表组件初始化完成")
    
    def _create_widgets(self):
        """
        创建界面元素
        
        构建文件列表的基本结构：
        - Canvas + Scrollbar（滚动区域）
        - 内部框架（用于放置文件项）
        - 绑定滚动事件
        
        初始状态显示空状态提示。
        
        结构与 TemplateSelector 保持一致，直接在 self 下创建滚动区域。
        """
        logger.debug("创建批量文件列表界面元素")
        
        # 配置 master 的 grid 布局
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=0)  # 滚动条列
        
        # 创建文件列表区域（直接在 master 下）
        self._create_file_list_area()
        
        logger.debug("批量文件列表界面元素创建完成")
    
    def _create_file_list_area(self):
        """
        创建文件列表滚动区域
        
        使用Canvas和Scrollbar实现可滚动的文件列表：
        - Canvas作为滚动容器
        - Scrollbar提供滚动条
        - 内部Frame容纳文件项
        - 自动调整滚动区域大小
        
        结构参考 TemplateSelector，直接在 self 下创建，无中间容器层。
        """
        logger.debug("创建文件列表区域")
        
        # 创建Canvas和Scrollbar（直接在 master 下）
        self.canvas = tk.Canvas(
            self.master, 
            bg="SystemButtonFace",  # 使用系统默认背景色
            highlightthickness=0,
            takefocus=1  # 允许Canvas接收焦点，以便响应键盘事件
        )
        logger.info(f"Canvas创建完成，takefocus={self.canvas.cget('takefocus')}")
        
        self.scrollbar = tb.Scrollbar(
            self.master, 
            orient=tk.VERTICAL, 
            command=self.canvas.yview,
            bootstyle="round-warning"
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 创建内部框架（用于放置文件项）
        self.inner_frame = tb.Frame(self.canvas, bootstyle="default")
        
        # 将内部框架添加到Canvas
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.inner_frame, anchor="nw"
        )
        
        # 布局Canvas和Scrollbar（使用 grid）
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        
        # 绑定事件
        self.inner_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        # 绑定键盘事件（上下箭头键移动文件位置）
        self.canvas.bind("<Up>", self._on_arrow_up)
        self.canvas.bind("<Down>", self._on_arrow_down)
        self.canvas.bind("<FocusIn>", self._on_canvas_focus_in)
        self.canvas.bind("<FocusOut>", self._on_canvas_focus_out)
        self.canvas.bind("<Button-1>", self._on_canvas_click)
        logger.info("键盘事件绑定完成: <Up>, <Down>, <FocusIn>, <FocusOut>, <Button-1>")
        
        # 初始为空状态
        self._show_empty_state()
    
    def _on_frame_configure(self, event):
        """内部框架大小变化时更新Canvas滚动区域"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _on_canvas_configure(self, event):
        """Canvas大小变化时调整内部框架宽度"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
    
    def _on_canvas_focus_in(self, event):
        """Canvas获得焦点时的处理"""
        logger.info("✓ Canvas获得焦点，键盘事件已激活")
    
    def _on_canvas_focus_out(self, event):
        """Canvas失去焦点时的处理"""
        logger.info("✗ Canvas失去焦点")
    
    def _on_canvas_click(self, event):
        """Canvas被点击时的处理 - 确保获取焦点以响应键盘事件"""
        logger.info("Canvas被点击，设置焦点")
        self.canvas.focus_set()
    
    def _on_arrow_up(self, event):
        """
        处理上箭头键事件：将选中文件向上移动一位
        
        如果有选中文件且不在第一位，则与上一个文件交换位置。
        交换后更新UI并保持选中状态和可见性。
        """
        if not self.selected_file:
            logger.debug("没有选中文件，忽略上箭头键")
            return "break"
        
        if len(self.files) <= 1:
            logger.debug("文件数量不足，无需移动")
            return "break"
        
        # 查找选中文件的索引
        try:
            current_index = self.files.index(self.selected_file)
        except ValueError:
            logger.warning("选中文件不在列表中")
            return "break"
        
        # 检查是否已经在第一位
        if current_index == 0:
            logger.debug("文件已在第一位，无法继续上移")
            return "break"
        
        # 与上一个文件交换位置
        self.files[current_index], self.files[current_index - 1] = \
            self.files[current_index - 1], self.files[current_index]
        
        logger.info(f"文件向上移动: {self.selected_file.file_name} (位置 {current_index + 1} -> {current_index})")
        
        # 重新渲染列表
        self._show_file_list()
        
        # 确保选中的文件可见
        self._scroll_to_selected()
        
        return "break"
    
    def _on_arrow_down(self, event):
        """
        处理下箭头键事件：将选中文件向下移动一位
        
        如果有选中文件且不在最后一位，则与下一个文件交换位置。
        交换后更新UI并保持选中状态和可见性。
        """
        if not self.selected_file:
            logger.debug("没有选中文件，忽略下箭头键")
            return "break"
        
        if len(self.files) <= 1:
            logger.debug("文件数量不足，无需移动")
            return "break"
        
        # 查找选中文件的索引
        try:
            current_index = self.files.index(self.selected_file)
        except ValueError:
            logger.warning("选中文件不在列表中")
            return "break"
        
        # 检查是否已经在最后一位
        if current_index == len(self.files) - 1:
            logger.debug("文件已在最后一位，无法继续下移")
            return "break"
        
        # 与下一个文件交换位置
        self.files[current_index], self.files[current_index + 1] = \
            self.files[current_index + 1], self.files[current_index]
        
        logger.info(f"文件向下移动: {self.selected_file.file_name} (位置 {current_index + 1} -> {current_index + 2})")
        
        # 重新渲染列表
        self._show_file_list()
        
        # 确保选中的文件可见
        self._scroll_to_selected()
        
        return "break"
    
    def _scroll_to_selected(self):
        """
        滚动Canvas以确保选中的文件项可见
        
        计算选中文件项的位置，如果超出可视区域则滚动到合适位置。
        """
        if not self.selected_file:
            return
        
        try:
            # 查找选中文件的索引
            current_index = self.files.index(self.selected_file)
            
            # 获取Canvas和滚动区域信息
            canvas_height = self.canvas.winfo_height()
            scroll_region = self.canvas.bbox("all")
            if not scroll_region:
                return
            
            total_height = scroll_region[3] - scroll_region[1]
            if total_height <= 0:
                return
            
            # 估算每个文件项的平均高度
            from gongwen_converter.utils.dpi_utils import scale
            item_pady = scale(5)
            estimated_item_height = scale(60)  # 估算值，包含边距
            
            # 计算选中文件项的大致位置
            item_top = current_index * estimated_item_height
            item_bottom = item_top + estimated_item_height
            
            # 获取当前可见区域
            current_view_top = self.canvas.yview()[0] * total_height
            current_view_bottom = self.canvas.yview()[1] * total_height
            
            # 判断是否需要滚动
            if item_top < current_view_top:
                # 文件项在可见区域上方，滚动到顶部
                scroll_position = max(0, item_top / total_height)
                self.canvas.yview_moveto(scroll_position)
                logger.debug(f"向上滚动以显示选中文件")
            elif item_bottom > current_view_bottom:
                # 文件项在可见区域下方，滚动到底部
                scroll_position = max(0, (item_bottom - canvas_height) / total_height)
                self.canvas.yview_moveto(scroll_position)
                logger.debug(f"向下滚动以显示选中文件")
        
        except Exception as e:
            logger.debug(f"滚动到选中文件失败: {e}")
    
    def _update_file_item_selection(self, file_info: FileInfo, is_selected: bool):
        """
        增量更新单个文件项的选中状态（只改变边框）
        
        用于同一选项卡内点击选择文件时，避免完整刷新导致的闪烁。
        只更新文件项的边框样式，不重建整个UI。
        
        参数:
            file_info: 要更新的文件信息对象
            is_selected: 是否选中
        """
        if file_info.file_path not in self.file_item_frames:
            logger.warning(f"找不到文件项UI引用: {file_info.file_path}")
            return
        
        frame = self.file_item_frames[file_info.file_path]
        
        try:
            from gongwen_converter.utils.dpi_utils import scale
            border_width = scale(2)
            
            if is_selected:
                # 选中状态：添加橙色边框
                style = tb.Style.get_instance()
                warning_color = style.colors.warning
                parent_bg = style.colors.bg
                
                # 注意：frame必须是tk.Frame才支持highlightthickness
                # 如果frame是tb.Frame，需要特殊处理
                if isinstance(frame, tk.Frame) and not isinstance(frame, tb.Frame):
                    frame.configure(
                        bg=parent_bg,
                        highlightbackground=warning_color,
                        highlightcolor=warning_color,
                        highlightthickness=border_width
                    )
                else:
                    # 如果是tb.Frame，需要销毁并重建（不常见，因为选中的都是tk.Frame）
                    logger.warning(f"文件项frame类型不支持增量更新: {type(frame)}")
                    # 回退到完整刷新
                    self._show_file_list()
                    return
            else:
                # 未选中状态：移除边框
                # 需要重新配置为tb.Frame的样式
                if isinstance(frame, tk.Frame) and not isinstance(frame, tb.Frame):
                    frame.configure(highlightthickness=0)
                else:
                    logger.warning(f"文件项frame类型不支持增量更新: {type(frame)}")
                    # 回退到完整刷新
                    self._show_file_list()
                    return
            
            logger.debug(f"增量更新文件项选中状态成功: {file_info.file_path}, 选中={is_selected}")
            
        except Exception as e:
            logger.error(f"增量更新文件项选中状态失败: {e}")
            # 失败则回退到完整刷新
            logger.info("回退到完整刷新")
            self._show_file_list()
    
    def _show_empty_state(self):
        """
        显示空状态提示
        
        当列表中没有文件时显示提示信息，
        引导用户拖拽文件到拖拽区开始批量处理。
        """
        logger.debug("显示空状态")
        
        # 清除现有内容
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        
        # 创建空状态提示
        from gongwen_converter.utils.dpi_utils import scale
        empty_padding = scale(20)
        empty_label = tb.Label(
            self.inner_frame,
            text="暂无文件\n拖拽文件（文件夹）进入，开始批量处理",
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            justify=tk.CENTER,
            anchor=tk.CENTER,
            padding=empty_padding
        )
        empty_label.pack(expand=True, fill=tk.BOTH)
    
    def _show_file_list(self):
        """
        显示文件列表（完整刷新）
        
        清除现有内容，重新渲染所有文件项。
        每个文件项包含序号/图标、文件信息和操作按钮。
        
        注意：完整刷新会清空UI引用字典，重新创建所有组件。
        """
        logger.debug("显示文件列表（完整刷新）")
        
        # 清空UI引用字典（避免内存泄漏）
        self.file_item_frames.clear()
        
        # 清除现有内容
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        
        # 按顺序创建所有文件项，传递索引
        for index, file_info in enumerate(self.files):
            self._create_file_item(file_info, index)
    
    def _create_file_item(self, file_info: FileInfo, index: int):
        """
        创建单个文件项
        
        构建文件项的完整界面：
        - 序号/选中图标/完成图标
        - 文件名和路径标签
        - 格式警告（如果有）
        - 文件大小
        - 操作按钮（置顶、打开、移除）
        
        参数:
            file_info: 文件信息对象
            index: 文件在列表中的索引（从0开始）
        """
        # 添加DPI缩放支持
        from gongwen_converter.utils.dpi_utils import scale
        icon_size = scale(16)
        border_width = scale(2)  # 边框宽度也需要DPI缩放
        # 边距缩放
        item_padx = scale(2)
        item_pady = scale(5) 
        index_padx_left = scale(5)
        index_padx_right = scale(10)
        size_padx = scale(2)  # 文件大小列边距（高DPI适配）
        button_padx = scale(2)
        
        # 根据选中状态设置框架样式（使用边框而不是背景色）
        # 【关键修复】统一使用 tk.Frame，这样所有文件项都支持 highlightthickness 参数
        # 不能混用 tk.Frame 和 tb.Frame，因为 tb.Frame 不支持 highlightthickness
        item_frame = tk.Frame(self.inner_frame)
        
        try:
            # 获取当前主题的颜色
            style = tb.Style.get_instance()
            warning_color = style.colors.warning
            parent_bg = style.colors.bg  # 从样式系统获取背景色
            
            if file_info.is_selected:
                # 选中状态：显示橙色边框
                item_frame.configure(
                    bg=parent_bg,
                    highlightbackground=warning_color,
                    highlightcolor=warning_color,
                    highlightthickness=border_width
                )
            else:
                # 未选中状态：无边框
                item_frame.configure(
                    bg=parent_bg,
                    highlightthickness=0
                )
        except Exception as e:
            logger.debug(f"设置边框颜色失败，使用默认: {e}")
            # 如果获取主题颜色失败，使用系统默认背景色
            if file_info.is_selected:
                item_frame.configure(
                    bg='SystemButtonFace',
                    highlightthickness=border_width
                )
            else:
                item_frame.configure(
                    bg='SystemButtonFace',
                    highlightthickness=0
                )
        
        item_frame.pack(fill=tk.X, padx=item_padx, pady=item_pady)
        
        # 保存文件项的UI引用（用于增量更新）
        self.file_item_frames[file_info.file_path] = item_frame
        
        # 配置网格布局
        item_frame.grid_columnconfigure(0, weight=0)  # 序号/选中图标列
        item_frame.grid_columnconfigure(1, weight=1)  # 文件信息（文件名和路径）
        item_frame.grid_columnconfigure(2, weight=0)  # 文件大小
        item_frame.grid_columnconfigure(3, weight=0)  # 打开按钮
        item_frame.grid_columnconfigure(4, weight=0)  # 移除按钮
        
        # 序号/完成图标/跳过图标/失败图标列
        # 逻辑：根据状态显示不同的内容
        # - completed: 显示完成图标，点击打开输出文件
        # - skipped: 显示跳过图标，点击显示跳过原因
        # - failed: 显示失败图标，点击显示错误详情
        # - 其他状态: 显示序号
        if file_info.status == 'completed':
            # 显示完成图标（转换完成）
            complete_icon = load_image_icon("complete_icon.png", master=item_frame, size=(icon_size, icon_size))
            if complete_icon:
                index_label = tb.Label(
                    item_frame,
                    image=complete_icon,
                    bootstyle="default",
                    anchor=tk.CENTER,
                    cursor="hand2"
                )
                index_label.image = complete_icon  # 保持对图片的引用
                # 绑定点击事件打开输出文件
                target_path = file_info.output_path if file_info.output_path else file_info.file_path
                index_label.bind("<Button-1>", lambda e, path=target_path: self._open_file(path))
                logger.debug(f"文件 {index + 1} 已完成，点击可打开: {target_path}")
            else:
                # 如果图标加载失败，显示序号
                index_label = tb.Label(
                    item_frame,
                    text=str(index + 1),
                    font=(self.small_font, self.small_size),
                    bootstyle="default",
                    anchor=tk.CENTER,
                    cursor="hand2"
                )
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._select_file(fi))
        elif file_info.status == 'skipped':
            # 显示跳过图标
            skip_icon = load_image_icon("skip_icon.png", master=item_frame, size=(icon_size, icon_size))
            if skip_icon:
                index_label = tb.Label(
                    item_frame,
                    image=skip_icon,
                    bootstyle="default",
                    anchor=tk.CENTER,
                    cursor="hand2"
                )
                index_label.image = skip_icon  # 保持对图片的引用
                # 绑定点击事件显示跳过原因对话框
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._show_skip_dialog(fi))
                logger.debug(f"文件 {index + 1} 已跳过，点击可查看原因")
            else:
                # 如果图标加载失败，显示序号
                index_label = tb.Label(
                    item_frame,
                    text=str(index + 1),
                    font=(self.small_font, self.small_size),
                    bootstyle="warning",
                    anchor=tk.CENTER,
                    cursor="hand2"
                )
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._show_skip_dialog(fi))
        elif file_info.status == 'failed':
            # 显示失败图标
            fail_icon = load_image_icon("fail_icon.png", master=item_frame, size=(icon_size, icon_size))
            if fail_icon:
                index_label = tb.Label(
                    item_frame,
                    image=fail_icon,
                    bootstyle="default",
                    anchor=tk.CENTER,
                    cursor="hand2"
                )
                index_label.image = fail_icon  # 保持对图片的引用
                # 绑定点击事件显示错误详情对话框
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._show_error_dialog(fi))
                logger.debug(f"文件 {index + 1} 处理失败，点击可查看详情")
            else:
                # 如果图标加载失败，显示序号
                index_label = tb.Label(
                    item_frame,
                    text=str(index + 1),
                    font=(self.small_font, self.small_size),
                    bootstyle="danger",
                    anchor=tk.CENTER,
                    cursor="hand2"
                )
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._show_error_dialog(fi))
        else:
            # 显示序号（待处理、处理中、选中状态等）
            index_label = tb.Label(
                item_frame,
                text=str(index + 1),
                font=(self.small_font, self.small_size),
                bootstyle="default",
                anchor=tk.CENTER,
                cursor="hand2"
            )
            index_label.bind("<Button-1>", lambda e, fi=file_info: self._select_file(fi))
        
        index_label.grid(row=0, column=0, sticky="ew", padx=(index_padx_left, index_padx_right))
        
        # 文件信息框架（包含文件名和路径）
        info_frame = tb.Frame(item_frame, bootstyle="default")
        info_frame.grid(row=0, column=1, sticky="ew")
        
        # 文件名标签（可点击）
        name_label = tb.Label(
            info_frame,
            text=file_info.file_name,
            font=(self.small_font, self.small_size),
            bootstyle="warning",
            anchor=tk.W,
            cursor="hand2"
        )
        name_label.pack(fill=tk.X)
        name_label.bind("<Button-1>", lambda e, fi=file_info: self._select_file(fi))
        
        # 文件路径 - 只显示目录（不包含文件名）
        path_label = tb.Label(
            info_frame,
            text=os.path.dirname(file_info.file_path),
            font=(self.micro_font, self.micro_size),
            bootstyle="secondary",
            anchor=tk.W
        )
        path_label.pack(fill=tk.X)
        
        # 格式验证和警告显示
        if file_info.warning_message:
            warning_label = tb.Label(
                info_frame,
                text=file_info.warning_message,
                font=(self.micro_font, self.micro_size),
                bootstyle="warning",
                anchor=tk.W
            )
            warning_label.pack(fill=tk.X)
            logger.debug(f"文件格式不匹配: {file_info.file_path} - {file_info.warning_message}")
        
        # 文件大小列
        size_label = tb.Label(
            item_frame,
            text=file_info.file_size,
            font=(self.micro_font, self.micro_size),
            bootstyle="secondary",
            anchor=tk.CENTER
        )
        size_label.grid(row=0, column=2, sticky="ew", padx=size_padx)
        
        # 打开按钮 - 打开原始文件所在文件夹
        open_icon = load_image_icon("location_icon.png", master=item_frame, size=(icon_size, icon_size))
        open_button = tb.Button(
            item_frame,
            image=open_icon,
            bootstyle="secondary-link",
            command=lambda fi=file_info: self._open_file(fi.file_path)
        )
        if open_icon:
            open_button.image = open_icon  # 保持对图片的引用
        open_button.grid(row=0, column=3, sticky="e", padx=0) 
        
        # 移除按钮
        remove_icon = load_image_icon("remove_icon.png", master=item_frame, size=(icon_size, icon_size))
        remove_button = tb.Button(
            item_frame,
            image=remove_icon,
            bootstyle="danger-link",
            command=lambda fi=file_info: self._remove_file(fi)
        )
        if remove_icon:
            remove_button.image = remove_icon  # 保持对图片的引用
        remove_button.grid(row=0, column=4, sticky="e", padx=0)
    
    def _select_file(self, file_info: FileInfo):
        """
        选中指定文件（使用增量更新避免闪烁）
        
        取消当前选中的文件，选中新文件，只更新边框样式而不重建整个列表。
        触发on_file_selected回调通知外部组件。
        
        参数:
            file_info: 要选中的文件信息对象
        """
        logger.debug(f"选中文件: {file_info.file_path}")
        
        # 记录之前选中的文件
        previous_selected = self.selected_file
        
        # 【关键修复】即使是已选中的文件，也要触发回调以确保UI同步
        # 这是为了处理以下场景：
        # 1. 用户切换到其他选项卡
        # 2. 再次拖入同一文件（文件已在列表中且已选中）
        # 3. 此时需要触发回调更新右栏UI（显示该文件的格式转换选项）
        # 如果不触发回调，右栏会继续显示其他选项卡文件的UI
        
        if previous_selected and previous_selected.file_path == file_info.file_path:
            logger.debug("重新选中已选中的文件，触发回调以同步UI")
            # 直接触发回调，不更新UI（因为UI已经是选中状态）
            if self.on_file_selected:
                self.on_file_selected(file_info)
            return
        
        # 更新数据层状态
        for file in self.files:
            file.is_selected = False
        file_info.is_selected = True
        self.selected_file = file_info
        
        # 使用增量更新UI（只改变边框，不重建整个列表）
        if previous_selected:
            self._update_file_item_selection(previous_selected, False)
        self._update_file_item_selection(file_info, True)
        
        # 【关键】让Canvas获取焦点，以便响应键盘事件
        logger.info("文件选中后，主动让Canvas获取焦点")
        self.canvas.focus_set()
        
        # 调用选中回调
        if self.on_file_selected:
            self.on_file_selected(file_info)
        
        logger.info(f"文件已选中（增量更新）: {file_info.file_path}")
    
    def _on_file_clicked(self, file_path: str):
        """处理文件点击事件"""
        logger.debug(f"文件被点击: {file_path}")
        if self.on_file_opened:
            self.on_file_opened(file_path)
    
    def _open_file(self, file_path: str):
        """打开文件所在文件夹并选中文件"""
        logger.debug(f"打开文件所在文件夹: {file_path}")
        
        try:
            if os.path.exists(file_path):
                # Windows: 使用 explorer /select 命令打开文件夹并选中文件
                subprocess.Popen(['explorer', '/select,', os.path.normpath(file_path)])
                logger.info(f"已打开文件夹并选中文件: {file_path}")
            else:
                logger.error(f"文件不存在: {file_path}")
        except Exception as e:
            logger.error(f"打开文件夹失败: {e}")
        
        # 同时调用回调函数（如果有）
        if self.on_file_opened:
            self.on_file_opened(file_path)
    
    def _remove_file(self, file_info: FileInfo):
        """
        从列表中移除文件，如果移除的是选中文件，自动选中另一个文件
        
        参数:
            file_info: 要移除的文件信息
        """
        logger.debug(f"移除文件: {file_info.file_path}")
        
        # 记录被移除文件的索引和是否被选中
        try:
            current_index = self.files.index(file_info)
            was_selected = self.selected_file and self.selected_file.file_path == file_info.file_path
        except ValueError:
            current_index = -1
            was_selected = False
            logger.warning(f"要移除的文件不在列表中: {file_info.file_path}")
        
        # 从列表中移除
        self.files = [f for f in self.files if f.file_path != file_info.file_path]
        
        # 如果移除的是选中文件，需要处理选中状态
        if was_selected:
            if self.files:
                # 列表中还有文件，自动选中一个
                # 优先选中被删除文件的前一个，如果没有则选中第一个
                new_index = max(0, current_index - 1) if current_index > 0 else 0
                new_selected = self.files[new_index]
                new_selected.is_selected = True
                self.selected_file = new_selected
                logger.info(f"自动选中文件: {new_selected.file_path} (位置: {new_index + 1})")
            else:
                # 列表为空，清空选中状态
                self.selected_file = None
        
        # 更新显示
        if self.files:
            self._show_file_list()
            # 如果选择了新文件，触发选中回调
            if was_selected and self.selected_file:
                if self.on_file_selected:
                    self.on_file_selected(self.selected_file)
                    logger.debug(f"触发文件选中回调: {self.selected_file.file_path}")
        else:
            self._show_empty_state()
            # 清空时重置类别
            self.current_category = None
            
            # 列表已清空，触发UI重置回调
            logger.info("批量列表已清空，触发UI重置回调")
            if self.on_list_cleared:
                self.on_list_cleared()
        
        # 调用文件移除回调函数
        if self.on_file_removed:
            self.on_file_removed(file_info.file_path)
    
    def _show_skip_dialog(self, file_info: FileInfo):
        """
        显示文件跳过原因对话框
        
        参数:
            file_info: 被跳过的文件信息
        """
        from tkinter import messagebox
        
        skip_reason = file_info.skip_reason or "未知原因"
        file_dir = os.path.dirname(file_info.file_path)
        
        message = f"文件已跳过转换\n\n"
        message += f"原因：{skip_reason}\n\n"
        message += f"原文件：{file_info.file_name}\n"
        message += f"位置：{file_dir}"
        
        logger.info(f"显示跳过对话框: {file_info.file_name}")
        
        # 显示信息对话框
        result = messagebox.showinfo(
            "文件已跳过",
            message,
            parent=self.master
        )
        
        # 询问是否打开文件位置
        if messagebox.askyesno("打开位置", "是否打开原文件位置？", parent=self.master):
            self._open_file(file_info.file_path)
    
    def _show_error_dialog(self, file_info: FileInfo):
        """
        显示文件处理失败详情对话框
        
        参数:
            file_info: 处理失败的文件信息
        """
        from tkinter import messagebox
        
        error_message = file_info.error_message or "未知错误"
        
        message = f"文件转换失败\n\n"
        message += f"文件：{file_info.file_name}\n"
        message += f"错误：{error_message}\n\n"
        message += f"建议：\n"
        message += f"1. 检查文件是否已损坏\n"
        message += f"2. 确认相关软件正常运行\n"
        message += f"3. 查看日志获取详细信息"
        
        logger.info(f"显示错误对话框: {file_info.file_name}")
        
        # 显示错误对话框
        messagebox.showerror(
            "转换失败",
            message,
            parent=self.master
        )
    
    def _can_add_file(self, file_path: str) -> Tuple[bool, str]:
        """
        检查文件是否可以添加到列表
        
        验证规则：
        - 文件必须是支持的类型
        - 列表为空时可以添加任何支持的文件
        - 列表不为空时只能添加相同类别的文件
        
        参数:
            file_path: 要检查的文件路径
            
        返回:
            Tuple[bool, str]: (是否可以添加, 错误消息)
        """
        # 获取文件实际类别
        file_category = get_actual_file_category(file_path)
        
        if file_category == 'unknown':
            return False, "不支持的文件类型"
        
        # 如果列表为空，可以添加任何类别，但要设置当前类别
        if not self.files:
            self.current_category = file_category  # 设置当前类别
            logger.debug(f"设置当前类别为: {file_category}")
            return True, ""
        
        # 如果列表不为空，检查类别是否一致
        if file_category != self.current_category:
            current_name = get_category_name(self.current_category)
            new_name = get_category_name(file_category)
            return False, f"只能添加{current_name}文件，不能混合{new_name}文件"
        
        return True, ""
    
    def add_file(self, file_path: str, auto_select: bool = False) -> Tuple[bool, str]:
        """
        添加单个文件到列表
        
        智能处理文件已存在的情况：
        - 文件不存在：正常添加到列表
        - 文件已存在且未选中：选中该文件，触发UI更新
        - 文件已存在且已选中：静默处理，不做任何操作
        
        执行流程：
        1. 检查文件是否已存在，如已存在则智能处理选中状态
        2. 检查文件类别是否匹配当前列表类别
        3. 创建FileInfo对象
        4. 添加到列表并更新显示
        5. 可选：自动选中文件并触发回调
        
        参数:
            file_path: 要添加的文件路径
            auto_select: 是否自动选中该文件（首个文件或显式指定时为True）
            
        返回:
            Tuple[bool, str]: (是否成功, 消息)
                - 成功：(True, "")
                - 失败：(False, 错误消息)
        """
        logger.debug(f"添加文件到批量列表: {file_path}")
        
        # 检查文件是否已存在
        existing_file = next((f for f in self.files if f.file_path == file_path), None)
        if existing_file:
            logger.debug(f"文件已存在于列表中: {file_path}")
            
            # 智能处理：文件已存在时，始终调用_select_file以确保UI同步
            # 即使文件已被选中，也需要重新选中以触发UI刷新
            # 这确保了当用户在不同选项卡之间切换后，UI能正确显示当前选项卡的文件
            logger.info(f"文件已存在，重新选中以同步UI: {file_path}")
            self._select_file(existing_file)
            return True, ""  # 返回成功，UI已通过_select_file更新
        
        # 检查是否可以添加（分类验证）
        can_add, error_msg = self._can_add_file(file_path)
        if not can_add:
            logger.warning(f"文件类别不匹配: {error_msg}")
            return False, error_msg
        
        # 创建FileInfo对象
        file_info = FileInfo(file_path)
        
        # 如果是第一个文件或者自动选中，则选中该文件
        should_select = auto_select or len(self.files) == 0
        if should_select:
            # 【关键修复】先取消所有文件的选中状态，确保只有一个文件被选中
            for f in self.files:
                f.is_selected = False
            
            file_info.is_selected = True
            self.selected_file = file_info
            logger.debug(f"自动选中文件: {file_path}")
        
        self.files.append(file_info)
        logger.info(f"文件已添加到批量列表: {file_path}")
        
        # 更新显示
        self._show_file_list()
        
        # 【关键修复】如果文件被自动选中，触发选中回调
        if should_select and self.on_file_selected:
            logger.debug(f"触发文件选中回调: {file_path}")
            self.on_file_selected(file_info)
        
        return True, ""
    
    def _determine_category_by_majority(self, file_paths: List[str]) -> Tuple[str, List[str], List[str]]:
        """
        使用多数决策算法确定文件类别
        
        当批量添加混合类型文件时，统计各类别文件数量，
        选择数量最多的类别作为列表类别。
        
        参数:
            file_paths: 要分析的文件路径列表
            
        返回:
            Tuple[str, List[str], List[str]]: 
                - 选定的文件类别
                - 匹配该类别的文件列表
                - 不匹配的文件列表
        """
        # 统计各类别文件数量（包含所有5种类别）
        category_count = {'text': 0, 'document': 0, 'spreadsheet': 0, 'layout': 0, 'image': 0}
        categorized_files = {'text': [], 'document': [], 'spreadsheet': [], 'layout': [], 'image': []}
        
        for file_path in file_paths:
            category = get_actual_file_category(file_path)
            if category in category_count:
                category_count[category] += 1
                categorized_files[category].append(file_path)
        
        # 选择数量最多的类别（排除数量为0的类别）
        non_zero_categories = {cat: count for cat, count in category_count.items() if count > 0}
        if not non_zero_categories:
            # 如果没有有效类别，默认选择第一个类别
            selected_category = 'text'
        else:
            selected_category = max(non_zero_categories, key=non_zero_categories.get)
        
        # 返回选定类别、匹配文件和不匹配文件
        matching_files = categorized_files[selected_category]
        non_matching_files = []
        for category, files in categorized_files.items():
            if category != selected_category and files:
                non_matching_files.extend(files)
        
        logger.info(f"根据多数原则选定类别: {selected_category}, 匹配文件: {len(matching_files)}, 不匹配文件: {len(non_matching_files)}")
        return selected_category, matching_files, non_matching_files

    def add_files(self, file_paths: List[str], auto_select_first: bool = False) -> Tuple[List[str], List[Tuple[str, str]]]:
        """
        批量添加多个文件到列表
        
        智能处理逻辑：
        - 列表为空：使用多数决策算法确定类别
        - 列表不为空：只添加匹配当前类别的文件
        
        参数:
            file_paths: 要添加的文件路径列表
            auto_select_first: 是否自动选中第一个成功添加的文件
            
        返回:
            Tuple[List[str], List[Tuple[str, str]]]: 
                - 成功添加的文件路径列表
                - 失败的文件列表（文件路径和错误消息的元组）
        """
        logger.debug(f"批量添加文件到列表: {len(file_paths)} 个文件")
        
        added_files = []
        failed_files = []
        
        # 如果列表为空，使用智能类别决策
        if not self.files:
            logger.info("列表为空，使用智能类别决策")
            selected_category, matching_files, non_matching_files = self._determine_category_by_majority(file_paths)
            
            # 设置当前类别
            self.current_category = selected_category
            logger.info(f"设置当前类别为: {selected_category}")
            
            # 添加匹配类别的文件
            for i, file_path in enumerate(matching_files):
                # 第一个文件自动选中
                auto_select = auto_select_first and i == 0
                success, error_msg = self.add_file(file_path, auto_select)
                if success:
                    added_files.append(file_path)
                else:
                    failed_files.append((file_path, error_msg))
            
            # 记录不匹配的文件
            for file_path in non_matching_files:
                category_names = {'text': '文本类', 'document': '文档类', 'spreadsheet': '表格类', 'layout': '版式类', 'image': '图片类'}
                selected_name = category_names.get(selected_category, selected_category)
                file_category = get_actual_file_category(file_path)
                file_category_name = category_names.get(file_category, file_category)
                failed_files.append((file_path, f"类型不匹配（当前列表为{selected_name}，文件为{file_category_name}）"))
        
        else:
            # 列表不为空，按原有逻辑处理
            logger.info("列表不为空，按原有类别检查逻辑处理")
            for file_path in file_paths:
                success, error_msg = self.add_file(file_path)
                if success:
                    added_files.append(file_path)
                else:
                    failed_files.append((file_path, error_msg))
        
        logger.info(f"批量添加完成: {len(added_files)}/{len(file_paths)} 个文件成功")
        if failed_files:
            logger.warning(f"有 {len(failed_files)} 个文件添加失败")
        
        return added_files, failed_files
    
    def clear_all_files(self):
        """清空所有文件"""
        logger.debug("清空批量文件列表")
        
        self.files.clear()
        self.current_category = None  # 重置类别
        self.selected_file = None  # 重置选中文件
        self._show_empty_state()
        
        logger.info("批量文件列表已清空")
    
    def get_files(self) -> List[str]:
        """
        获取所有文件路径
        
        返回:
            List[str]: 文件路径列表（原始文件路径）
        """
        return [f.file_path for f in self.files]
    
    def get_file_objects(self) -> List[FileInfo]:
        """
        获取所有文件对象
        
        返回:
            List[FileInfo]: 文件信息对象列表
        """
        return self.files.copy()
    
    def get_selected_file(self) -> Optional[FileInfo]:
        """
        获取当前选中的文件
        
        返回:
            Optional[FileInfo]: 选中的文件信息，如果没有选中则返回None
        """
        return self.selected_file
    
    def get_selected_file_path(self) -> Optional[str]:
        """
        获取当前选中文件的路径
        
        返回:
            Optional[str]: 选中文件的路径，如果没有选中则返回None
        """
        return self.selected_file.file_path if self.selected_file else None
    
    def get_file_count(self) -> int:
        """
        获取文件数量
        
        返回:
            int: 文件数量
        """
        return len(self.files)
    
    def get_file_by_path(self, file_path: str) -> Optional[FileInfo]:
        """
        根据文件路径获取FileInfo对象
        
        用于从批量列表中获取文件的缓存信息（如actual_format）。
        
        参数:
            file_path: 文件路径
            
        返回:
            Optional[FileInfo]: 找到的FileInfo对象，如果不存在则返回None
        """
        for file_info in self.files:
            if file_info.file_path == file_path:
                return file_info
        return None
    
    def has_files(self) -> bool:
        """
        检查是否有文件
        
        返回:
            bool: 是否有文件
        """
        return len(self.files) > 0
    
    def get_current_category(self) -> Optional[str]:
        """
        获取当前文件类别
        
        返回:
            Optional[str]: 当前文件类别，如果列表为空则返回None
        """
        return self.current_category
    
    def update_file_status(self, file_path: str, status: str, output_path: Optional[str] = None, 
                           skip_reason: Optional[str] = None, error_message: Optional[str] = None):
        """
        更新文件的处理状态
        
        在文件转换过程中更新其状态，并重新渲染列表以显示新状态。
        
        参数:
            file_path: 原始文件路径
            status: 新的处理状态
                - 'pending': 待处理
                - 'processing': 处理中
                - 'completed': 已完成
                - 'skipped': 已跳过
                - 'failed': 失败
            output_path: 输出文件路径（status为completed时提供）
            skip_reason: 跳过原因（status为skipped时提供）
            error_message: 错误信息（status为failed时提供）
        """
        logger.debug(f"更新文件状态: {file_path} -> {status}")
        if output_path:
            logger.debug(f"输出路径: {output_path}")
        if skip_reason:
            logger.debug(f"跳过原因: {skip_reason}")
        if error_message:
            logger.debug(f"错误信息: {error_message}")
        
        # 查找并更新文件状态
        for file_info in self.files:
            if file_info.file_path == file_path:
                file_info.status = status
                
                # 根据状态更新相应的额外信息
                if status == 'completed' and output_path:
                    file_info.output_path = output_path
                elif status == 'skipped' and skip_reason:
                    file_info.skip_reason = skip_reason
                elif status == 'failed' and error_message:
                    file_info.error_message = error_message
                
                logger.info(f"文件状态已更新: {file_path} -> {status}")
                break
        
        # 重新渲染列表以更新显示
        self._show_file_list()
    
    def keep_first_file_only(self):
        """
        只保留第一个文件（用于批量模式切换到单文件模式）
        """
        logger.debug("只保留第一个文件")
        
        if len(self.files) > 1:
            first_file = self.files[0]
            self.files = [first_file]
            
            # 确保第一个文件被选中
            first_file.is_selected = True
            self.selected_file = first_file
            
            self._show_file_list()
            logger.info(f"已保留第一个文件: {first_file.file_path}")
        elif len(self.files) == 0:
            logger.debug("列表为空，无需操作")
        else:
            logger.debug("列表只有一个文件，无需操作")
    
    def add_file_with_mode(self, file_path: str, mode: str = "batch", auto_select: bool = True) -> Tuple[bool, str]:
        """
        根据模式添加文件到列表
        
        参数:
            file_path: 文件路径
            mode: 模式 ("single" 或 "batch")
            auto_select: 是否自动选中文件
            
        返回:
            Tuple[bool, str]: (是否成功添加, 错误消息)
        """
        logger.debug(f"添加文件到列表 (模式: {mode}): {file_path}")
        
        # 单文件模式：清空列表，添加新文件（替换）
        if mode == "single":
            self.files.clear()
            self.current_category = None
            self.selected_file = None
            logger.debug("单文件模式：已清空列表")
        
        # 调用 add_file 方法，设置自动选中
        return self.add_file(file_path, auto_select)
    
    def _setup_drag_drop(self):
        """设置拖拽支持"""
        logger.debug("设置文件列表拖拽支持")
        
        if not TKINTERDND2_AVAILABLE:
            logger.warning("tkinterdnd2未安装，文件列表拖拽功能不可用")
            self.drag_enabled = False
            return
        
        try:
            # 将Canvas注册为拖拽目标
            self.canvas.drop_target_register(DND_FILES)
            self.canvas.dnd_bind('<<DragEnter>>', self._on_drag_enter)
            self.canvas.dnd_bind('<<DragLeave>>', self._on_drag_leave)
            self.canvas.dnd_bind('<<Drop>>', self._on_drop)
            logger.info("文件列表拖拽支持设置成功")
        except Exception as e:
            logger.error(f"设置文件列表拖拽支持失败: {str(e)}")
            self.drag_enabled = False
    
    def _on_drag_enter(self, event):
        """处理拖拽进入事件"""
        if not self.drag_enabled:
            return
        
        logger.debug("文件拖拽进入文件列表区域")
        self.is_dragging = True
        
        return "copy"
    
    def _on_drag_leave(self, event):
        """处理拖拽离开事件"""
        if not self.drag_enabled:
            return
        
        logger.debug("文件拖拽离开文件列表区域")
        self.is_dragging = False
    
    def _on_drop(self, event):
        """处理文件放置事件 - 使用智能分类和自动切换选项卡"""
        if not self.drag_enabled:
            return
        
        logger.debug("文件放置到文件列表事件触发")
        self.is_dragging = False
        
        try:
            file_data = event.data.strip()
            logger.debug(f"文件列表接收到拖拽数据: {file_data}")
            
            # 解析拖拽的文件
            files = self._parse_dropped_files(file_data)
            
            if len(files) > 0:
                logger.info(f"文件列表拖拽文件: {len(files)} 个文件")
                
                # 处理文件（支持文件夹递归）
                processed_files = self._process_batch_files(files)
                logger.info(f"处理后得到 {len(processed_files)} 个支持文件")
                
                if processed_files:
                    # 获取 TabbedFileSelector 的引用
                    tabbed_list = self._get_tabbed_batch_file_list()
                    
                    if tabbed_list:
                        # 使用 TabbedFileSelector 的 add_files 方法
                        # 这样可以实现智能分类和自动切换选项卡
                        added_files, failed_files = tabbed_list.add_files(processed_files)
                        
                        if added_files:
                            logger.info(f"成功添加 {len(added_files)} 个文件，已自动分类和切换选项卡")
                        
                        if failed_files:
                            logger.warning(f"有 {len(failed_files)} 个文件添加失败")
                            for file_path, error_msg in failed_files:
                                logger.debug(f"失败文件: {file_path}, 原因: {error_msg}")
                    else:
                        # 如果无法获取 TabbedFileSelector，回退到本地处理
                        logger.warning("无法获取 TabbedFileSelector，使用本地处理")
                        added_files, failed_files = self.add_files(processed_files, auto_select_first=True)
                        
                        if added_files:
                            logger.info(f"成功添加 {len(added_files)} 个文件到当前列表")
                else:
                    logger.warning("未找到任何支持的文件")
            else:
                logger.warning("未检测到有效文件")
        
        except Exception as e:
            logger.error(f"处理文件列表拖拽失败: {str(e)}")
    
    def _get_tabbed_batch_file_list(self):
        """获取 TabbedFileSelector 的引用"""
        try:
            # 向上查找主窗口
            current = self.master
            while current and not hasattr(current, '_main_window'):
                current = current.master if hasattr(current, 'master') else None
            
            if current and hasattr(current, '_main_window'):
                main_window = current._main_window
                if hasattr(main_window, 'tabbed_batch_file_list'):
                    return main_window.tabbed_batch_file_list
            
            # 如果上面的方法失败，尝试直接从父组件获取
            # FileSelector 的父组件应该是 TabbedFileSelector 中的一个 tab
            current = self.master
            while current:
                # 检查父组件的父组件是否是 TabbedFileSelector
                if hasattr(current, 'master') and current.master:
                    parent = current.master
                    if hasattr(parent, 'master') and parent.master:
                        grandparent = parent.master
                        if grandparent.__class__.__name__ == 'TabbedFileSelector':
                            return grandparent
                current = current.master if hasattr(current, 'master') else None
                
        except Exception as e:
            logger.error(f"获取 TabbedFileSelector 失败: {e}")
        
        return None
    
    def _parse_dropped_files(self, file_data: str) -> List[str]:
        """
        解析拖拽的文件数据 - 改进版，支持混合格式
        
        能够正确处理以下格式：
        - {file with spaces.jpg} file2.jpg
        - {file1} {file2}
        - file1 file2
        - 单个文件
        """
        import re
        
        files = []
        
        # 策略1: 使用正则表达式提取所有花括号包裹的路径
        # 匹配模式: {路径内容}
        brace_pattern = r'\{([^}]+)\}'
        brace_matches = re.findall(brace_pattern, file_data)
        
        if brace_matches:
            # 找到了花括号包裹的路径
            files.extend(brace_matches)
            
            # 移除已提取的花括号部分，处理剩余内容
            remaining = re.sub(brace_pattern, '', file_data).strip()
            
            if remaining:
                # 剩余部分可能包含没有花括号的路径（空格分隔）
                potential_paths = remaining.split()
                for path in potential_paths:
                    path = path.strip()
                    if path and os.path.exists(path):
                        files.append(path)
        else:
            # 没有花括号，使用传统的空格分隔方式
            if ' ' in file_data:
                # 先尝试整个字符串是否是有效路径
                if os.path.exists(file_data):
                    files = [file_data]
                else:
                    # 按空格分割并验证每个路径
                    potential_files = file_data.split(' ')
                    files = [f for f in potential_files if f.strip() and os.path.exists(f.strip())]
                    if not files:
                        # 如果都无效，保留原始数据
                        files = [file_data]
            else:
                # 单个文件
                files = [file_data]
        
        # 清理并去重
        cleaned_files = []
        seen = set()
        for f in files:
            f = f.strip()
            if f and f not in seen:
                cleaned_files.append(f)
                seen.add(f)
        
        return cleaned_files
    
    def _process_batch_files(self, files: List[str]) -> List[str]:
        """
        处理批量模式下的拖拽文件，支持文件夹递归
        
        复用FileDropArea的处理逻辑
        """
        processed_files = []
        
        for file_path in files:
            if os.path.isdir(file_path):
                # 如果是文件夹，递归收集所有支持的文件
                folder_files = collect_files_from_folder(file_path)
                logger.info(f"从文件夹 {file_path} 收集到 {len(folder_files)} 个支持文件")
                processed_files.extend(folder_files)
            else:
                # 如果是文件，检查是否支持
                if is_supported_file(file_path):
                    processed_files.append(file_path)
                else:
                    logger.debug(f"跳过不支持的文件: {file_path}")
        
        return processed_files


# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 创建测试窗口
    root = tb.Window(title="批量文件列表组件测试", themename="morph")
    root.geometry("400x500")
    
    def on_file_removed(file_path):
        logger.info(f"文件被移除: {file_path}")
    
    def on_file_opened(file_path):
        logger.info(f"文件被打开: {file_path}")
    
    def on_file_selected(file_info):
        logger.info(f"文件被选中: {file_info.file_path}")
    
    # 创建批量文件列表组件
    file_list = FileSelector(
        root,
        on_file_removed=on_file_removed,
        on_file_opened=on_file_opened,
        on_file_selected=on_file_selected
    )
    file_list.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
    
    # 添加测试按钮框架
    test_frame = tb.Frame(root, bootstyle="default")
    test_frame.pack(fill=tk.X, padx=10, pady=10)
    
    # 测试按钮回调函数
    def test_add_file():
        file_list.add_file("/fake/path/test1.docx", auto_select=True)
    
    def test_add_files():
        files = [
            "/fake/path/test2.md",
            "/fake/path/test3.xlsx", 
            "/fake/path/test4.docx"
        ]
        file_list.add_files(files, auto_select_first=True)
    
    def test_clear():
        file_list.clear_all_files()
    
    def test_get_selected():
        selected = file_list.get_selected_file()
        if selected:
            logger.info(f"当前选中文件: {selected.file_path}")
        else:
            logger.info("没有选中文件")
    
    # 创建测试按钮
    buttons = [
        ("添加文件", test_add_file, "primary"),
        ("批量添加", test_add_files, "success"), 
        ("清空", test_clear, "danger"),
        ("获取选中", test_get_selected, "info")
    ]
    
    for text, command, style in buttons:
        btn = tb.Button(
            test_frame,
            text=text,
            command=command,
            bootstyle=style,
            width=10
        )
        btn.pack(side=tk.LEFT, padx=5)
    
    # 运行测试
    logger.info("启动批量文件列表组件测试")
    root.mainloop()
