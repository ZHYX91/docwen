"""
文件选择器核心类模块

提供批量文件列表的核心功能（不含拖拽功能）：
- 文件列表显示（序号、文件名、路径、大小）
- 文件操作按钮（移至顶部、打开位置、移除）
- 文件选中机制（单选，支持回调）
- 智能文件分类（根据实际格式，非扩展名）
- 格式验证和警告提示
- 完成状态显示（转换完成后显示图标）

拖拽功能由 DragDropMixin 提供，通过多继承组合。
"""

import logging
import tkinter as tk
from collections.abc import Callable
from pathlib import Path

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.i18n import t
from docwen.utils.file_type_utils import (
    get_actual_file_category,
    get_category_name,
)
from docwen.utils.font_utils import get_default_font, get_micro_font, get_small_font
from docwen.utils.gui_utils import bind_single_line_ellipsis
from docwen.utils.icon_utils import load_image_icon

from .dialogs import open_file_location, show_error_dialog, show_skip_dialog
from .file_info import FileInfo

logger = logging.getLogger(__name__)


class FileSelectorCore:
    """
    批量文件列表核心类

    批量模式下的核心文件管理界面，提供：
    - 滚动文件列表显示
    - 文件添加和移除
    - 文件选中和回调
    - 同类文件验证
    - 状态管理和更新
    - 空状态提示

    此类不包含拖拽功能，拖拽功能由 DragDropMixin 提供。

    属性:
        master: 父组件对象
        files: 文件信息列表
        current_category: 当前文件类别
        selected_file: 当前选中的文件
        canvas: 滚动画布
        inner_frame: 内部框架（用于放置文件项）

    回调函数:
        on_file_removed: 文件移除时的回调
        on_file_opened: 文件打开时的回调
        on_list_cleared: 列表清空时的回调
        on_file_selected: 文件选中时的回调
    """

    def __init__(
        self,
        master,
        on_file_removed: Callable | None = None,
        on_file_opened: Callable | None = None,
        on_list_cleared: Callable | None = None,
        on_file_selected: Callable | None = None,
        **kwargs,
    ):
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
        self.files: list[FileInfo] = []
        self.current_category: str | None = None  # 当前文件类别：text/document/spreadsheet/image/layout
        self.selected_file: FileInfo | None = None  # 当前选中的文件

        # UI引用管理：保存每个文件项的frame引用，用于增量更新
        self.file_item_frames: dict[str, tk.Frame] = {}

        # 拖拽状态（由 DragDropMixin 使用）
        self.is_dragging = False
        self.drag_enabled = True

        # 获取字体配置
        self.default_font, self.default_size = get_default_font()
        self.small_font, self.small_size = get_small_font()
        self.micro_font, self.micro_size = get_micro_font()

        # 创建界面元素
        self._create_widgets()

        # 设置拖拽支持（由 DragDropMixin 提供）
        self._setup_drag_drop()

        logger.info("批量文件列表组件初始化完成")

    def _setup_drag_drop(self):
        """
        设置拖拽支持（占位方法）

        此方法由 DragDropMixin 覆盖，提供实际的拖拽功能。
        如果未混入 DragDropMixin，此方法将被调用但不执行任何操作。
        """
        logger.debug("拖拽功能未启用（需要通过 DragDropMixin 混入）")

    def _create_widgets(self):
        """
        创建界面元素

        构建文件列表的基本结构：
        - Canvas + Scrollbar（滚动区域）
        - 内部框架（用于放置文件项）
        - 绑定滚动事件

        初始状态显示空状态提示。
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
        """
        logger.debug("创建文件列表区域")

        # 创建Canvas和Scrollbar（直接在 master 下）
        self.canvas = tb.Canvas(self.master, highlightthickness=0, takefocus=1)
        logger.info(f"Canvas创建完成，takefocus={self.canvas.cget('takefocus')}")

        # 自定义滚动命令（带边界检查，防止滚动超出内容区域）
        def bounded_yview(*args):
            """带边界检查的滚动命令"""
            if args[0] == "scroll":
                # 点击滚动条按钮时的滚动操作：检查是否已到边界
                top, bottom = self.canvas.yview()
                direction = int(args[1])
                if top <= 0 and direction < 0:  # 已在顶部，不能再向上滚动
                    return
                if bottom >= 1 and direction > 0:  # 已在底部，不能再向下滚动
                    return
            self.canvas.yview(*args)

        self.scrollbar = tb.Scrollbar(self.master, orient=tk.VERTICAL, command=bounded_yview, bootstyle="round-warning")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # 创建内部框架（用于放置文件项）
        self.inner_frame = tb.Frame(self.canvas, bootstyle="default")

        # 将内部框架添加到Canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

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
        self.files[current_index], self.files[current_index - 1] = (
            self.files[current_index - 1],
            self.files[current_index],
        )

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
        self.files[current_index], self.files[current_index + 1] = (
            self.files[current_index + 1],
            self.files[current_index],
        )

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
            from docwen.utils.dpi_utils import scale

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
                logger.debug("向上滚动以显示选中文件")
            elif item_bottom > current_view_bottom:
                # 文件项在可见区域下方，滚动到底部
                scroll_position = max(0, (item_bottom - canvas_height) / total_height)
                self.canvas.yview_moveto(scroll_position)
                logger.debug("向下滚动以显示选中文件")

        except Exception as e:
            logger.debug(f"滚动到选中文件失败: {e}")

    def _update_file_item_selection(self, file_info: FileInfo, is_selected: bool):
        """
        增量更新单个文件项的选中状态（只改变边框）

        用于同一选项卡内点击选择文件时，避免完整刷新导致的闪烁。

        参数:
            file_info: 要更新的文件信息对象
            is_selected: 是否选中
        """
        if file_info.file_path not in self.file_item_frames:
            logger.warning(f"找不到文件项UI引用: {file_info.file_path}")
            return

        frame = self.file_item_frames[file_info.file_path]

        try:
            from docwen.utils.dpi_utils import scale

            border_width = scale(2)

            if is_selected:
                # 选中状态：添加橙色边框
                style = tb.Style.get_instance()
                warning_color = style.colors.warning
                parent_bg = style.colors.bg

                frame.configure(
                    bg=parent_bg,
                    highlightbackground=warning_color,
                    highlightcolor=warning_color,
                    highlightthickness=border_width,
                )
            else:
                # 未选中状态：移除边框
                frame.configure(highlightthickness=0)

            logger.debug(f"增量更新文件项选中状态成功: {file_info.file_path}, 选中={is_selected}")

        except Exception as e:
            logger.error(f"增量更新文件项选中状态失败: {e}")
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
        from docwen.utils.dpi_utils import scale

        empty_padding = scale(20)
        empty_label = tb.Label(
            self.inner_frame,
            text=t("components.file_selector.empty_hint"),
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            justify=tk.CENTER,
            anchor=tk.CENTER,
            padding=empty_padding,
        )
        empty_label.pack(expand=True, fill=tk.BOTH)

    def _show_file_list(self):
        """
        显示文件列表（完整刷新）

        清除现有内容，重新渲染所有文件项。
        """
        logger.debug("显示文件列表（完整刷新）")

        # 清空UI引用字典
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
        from docwen.utils.dpi_utils import scale

        icon_size = scale(16)
        border_width = scale(2)
        item_padx = scale(2)
        item_pady = scale(5)
        index_padx_left = scale(5)
        index_padx_right = scale(10)
        size_padx = scale(2)

        # 根据选中状态设置框架样式
        # 统一使用 tk.Frame，以支持 highlightthickness 参数
        item_frame = tk.Frame(self.inner_frame)

        try:
            # 获取当前主题的颜色
            style = tb.Style.get_instance()
            warning_color = style.colors.warning
            parent_bg = style.colors.bg

            if file_info.is_selected:
                # 选中状态：显示橙色边框
                item_frame.configure(
                    bg=parent_bg,
                    highlightbackground=warning_color,
                    highlightcolor=warning_color,
                    highlightthickness=border_width,
                )
            else:
                # 未选中状态：无边框
                item_frame.configure(bg=parent_bg, highlightthickness=0)
        except Exception as e:
            logger.debug(f"设置边框颜色失败，使用默认: {e}")
            try:
                parent_bg = self.inner_frame.cget("background")
            except Exception:
                parent_bg = None
            if file_info.is_selected:
                if parent_bg:
                    item_frame.configure(bg=parent_bg, highlightthickness=border_width)
                else:
                    item_frame.configure(highlightthickness=border_width)
            else:
                if parent_bg:
                    item_frame.configure(bg=parent_bg, highlightthickness=0)
                else:
                    item_frame.configure(highlightthickness=0)

        item_frame.pack(fill=tk.X, padx=item_padx, pady=item_pady)

        # 保存文件项的UI引用
        self.file_item_frames[file_info.file_path] = item_frame

        # 配置网格布局
        item_frame.grid_columnconfigure(0, weight=0)  # 序号/选中图标列
        item_frame.grid_columnconfigure(1, weight=1)  # 文件信息
        item_frame.grid_columnconfigure(2, weight=0)  # 文件大小
        item_frame.grid_columnconfigure(3, weight=0)  # 打开按钮
        item_frame.grid_columnconfigure(4, weight=0)  # 移除按钮

        # 序号/完成图标/跳过图标/失败图标列
        if file_info.status == "completed":
            # 显示完成图标
            complete_icon = load_image_icon("complete_icon.png", master=item_frame, size=(icon_size, icon_size))
            if complete_icon:
                index_label = tb.Label(
                    item_frame, image=complete_icon, bootstyle="default", anchor=tk.CENTER, cursor="hand2"
                )
                index_label.image = complete_icon
                target_path = file_info.output_path if file_info.output_path else file_info.file_path
                index_label.bind("<Button-1>", lambda e, path=target_path: self._open_file(path))
                logger.debug(f"文件 {index + 1} 已完成，点击可打开: {target_path}")
            else:
                index_label = tb.Label(
                    item_frame,
                    text=str(index + 1),
                    font=(self.small_font, self.small_size),
                    bootstyle="default",
                    anchor=tk.CENTER,
                    cursor="hand2",
                )
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._select_file(fi))
        elif file_info.status == "skipped":
            # 显示跳过图标
            skip_icon = load_image_icon("skip_icon.png", master=item_frame, size=(icon_size, icon_size))
            if skip_icon:
                index_label = tb.Label(
                    item_frame, image=skip_icon, bootstyle="default", anchor=tk.CENTER, cursor="hand2"
                )
                index_label.image = skip_icon
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._show_skip_dialog(fi))
                logger.debug(f"文件 {index + 1} 已跳过，点击可查看原因")
            else:
                index_label = tb.Label(
                    item_frame,
                    text=str(index + 1),
                    font=(self.small_font, self.small_size),
                    bootstyle="warning",
                    anchor=tk.CENTER,
                    cursor="hand2",
                )
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._show_skip_dialog(fi))
        elif file_info.status == "failed":
            # 显示失败图标
            fail_icon = load_image_icon("fail_icon.png", master=item_frame, size=(icon_size, icon_size))
            if fail_icon:
                index_label = tb.Label(
                    item_frame, image=fail_icon, bootstyle="default", anchor=tk.CENTER, cursor="hand2"
                )
                index_label.image = fail_icon
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._show_error_dialog(fi))
                logger.debug(f"文件 {index + 1} 处理失败，点击可查看详情")
            else:
                index_label = tb.Label(
                    item_frame,
                    text=str(index + 1),
                    font=(self.small_font, self.small_size),
                    bootstyle="danger",
                    anchor=tk.CENTER,
                    cursor="hand2",
                )
                index_label.bind("<Button-1>", lambda e, fi=file_info: self._show_error_dialog(fi))
        else:
            # 显示序号
            index_label = tb.Label(
                item_frame,
                text=str(index + 1),
                font=(self.small_font, self.small_size),
                bootstyle="default",
                anchor=tk.CENTER,
                cursor="hand2",
            )
            index_label.bind("<Button-1>", lambda e, fi=file_info: self._select_file(fi))

        index_label.grid(row=0, column=0, sticky="ew", padx=(index_padx_left, index_padx_right))

        # 文件信息框架
        info_frame = tb.Frame(item_frame, bootstyle="default")
        info_frame.grid(row=0, column=1, sticky="ew")

        # 文件名标签
        full_name_text = file_info.file_name
        name_label = tb.Label(
            info_frame,
            text=full_name_text,
            font=(self.small_font, self.small_size),
            bootstyle="warning",
            anchor=tk.W,
            cursor="hand2",
        )
        name_label.pack(fill=tk.X)
        bind_single_line_ellipsis(name_label, full_name_text, padding=4, mode="end")
        name_label.bind("<Button-1>", lambda e, fi=file_info: self._select_file(fi))

        # 文件路径
        full_path_text = str(Path(file_info.file_path).parent)
        path_label = tb.Label(
            info_frame, text=full_path_text, font=(self.micro_font, self.micro_size), bootstyle="secondary", anchor=tk.W
        )
        path_label.pack(fill=tk.X)
        bind_single_line_ellipsis(path_label, full_path_text, padding=4, mode="middle")

        # 格式警告
        if file_info.warning_message:
            full_warning_text = file_info.warning_message
            warning_label = tb.Label(
                info_frame,
                text=full_warning_text,
                font=(self.micro_font, self.micro_size),
                bootstyle="warning",
                anchor=tk.W,
            )
            warning_label.pack(fill=tk.X)
            bind_single_line_ellipsis(warning_label, full_warning_text, padding=4, mode="end")
            logger.debug(f"文件格式不匹配: {file_info.file_path} - {file_info.warning_message}")

        # 文件大小
        size_label = tb.Label(
            item_frame,
            text=file_info.file_size,
            font=(self.micro_font, self.micro_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
        )
        size_label.grid(row=0, column=2, sticky="ew", padx=size_padx)

        # 打开按钮
        open_icon = load_image_icon("location_icon.png", master=item_frame, size=(icon_size, icon_size))
        open_button = tb.Button(
            item_frame,
            image=open_icon,
            bootstyle="secondary-link",
            command=lambda fi=file_info: self._open_file(fi.file_path),
        )
        if open_icon:
            open_button.image = open_icon
        open_button.grid(row=0, column=3, sticky="e", padx=0)

        # 移除按钮
        remove_icon = load_image_icon("remove_icon.png", master=item_frame, size=(icon_size, icon_size))
        remove_button = tb.Button(
            item_frame, image=remove_icon, bootstyle="danger-link", command=lambda fi=file_info: self._remove_file(fi)
        )
        if remove_icon:
            remove_button.image = remove_icon
        remove_button.grid(row=0, column=4, sticky="e", padx=0)

    def _select_file(self, file_info: FileInfo):
        """
        选中指定文件

        取消当前选中的文件，选中新文件，触发回调。

        参数:
            file_info: 要选中的文件信息对象
        """
        logger.debug(f"选中文件: {file_info.file_path}")

        previous_selected = self.selected_file

        # 如果是已选中的文件，也要触发回调以确保UI同步
        if previous_selected and previous_selected.file_path == file_info.file_path:
            logger.debug("重新选中已选中的文件，触发回调以同步UI")
            if self.on_file_selected:
                self.on_file_selected(file_info)
            return

        # 更新数据层状态
        for file in self.files:
            file.is_selected = False
        file_info.is_selected = True
        self.selected_file = file_info

        # 使用增量更新UI
        if previous_selected:
            self._update_file_item_selection(previous_selected, False)
        self._update_file_item_selection(file_info, True)

        # 让Canvas获取焦点
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
        open_file_location(file_path)

        if self.on_file_opened:
            self.on_file_opened(file_path)

    def _remove_file(self, file_info: FileInfo):
        """
        从列表中移除文件

        如果移除的是选中文件，自动选中另一个文件。

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
                new_index = max(0, current_index - 1) if current_index > 0 else 0
                new_selected = self.files[new_index]
                new_selected.is_selected = True
                self.selected_file = new_selected
                logger.info(f"自动选中文件: {new_selected.file_path} (位置: {new_index + 1})")
            else:
                self.selected_file = None

        # 更新显示
        if self.files:
            self._show_file_list()
            if was_selected and self.selected_file and self.on_file_selected:
                self.on_file_selected(self.selected_file)
                logger.debug(f"触发文件选中回调: {self.selected_file.file_path}")
        else:
            self._show_empty_state()
            self.current_category = None

            logger.info("批量列表已清空，触发UI重置回调")
            if self.on_list_cleared:
                self.on_list_cleared()

        if self.on_file_removed:
            self.on_file_removed(file_info.file_path)

    def _show_skip_dialog(self, file_info: FileInfo):
        """显示文件跳过原因对话框"""
        show_skip_dialog(self.master, file_info)

    def _show_error_dialog(self, file_info: FileInfo):
        """显示文件处理失败详情对话框"""
        show_error_dialog(self.master, file_info)

    def _can_add_file(self, file_path: str) -> tuple[bool, str]:
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
        file_category = get_actual_file_category(file_path)
        return self._can_add_category(file_category)

    def _can_add_category(self, file_category: str | None) -> tuple[bool, str]:
        if not file_category:
            file_category = "unknown"

        if file_category == "unknown":
            return False, t("messages.invalid_file_type")

        if not self.files:
            if not self.current_category:
                self.current_category = file_category
            logger.debug(f"设置当前类别为: {file_category}")
            return True, ""

        if not self.current_category:
            self.current_category = file_category

        if file_category != self.current_category:
            current_name = get_category_name(self.current_category, t_func=t)
            new_name = get_category_name(file_category, t_func=t)
            return False, t("messages.category_mismatch", current_name=current_name, new_name=new_name)

        return True, ""

    def add_file(self, file_path: str, auto_select: bool = False) -> tuple[bool, str]:
        """
        添加单个文件到列表

        智能处理文件已存在的情况：
        - 文件不存在：正常添加到列表
        - 文件已存在且未选中：选中该文件
        - 文件已存在且已选中：静默处理

        参数:
            file_path: 要添加的文件路径
            auto_select: 是否自动选中该文件

        返回:
            Tuple[bool, str]: (是否成功, 消息)
        """
        logger.debug(f"添加文件到批量列表: {file_path}")

        # 检查文件是否已存在
        existing_file = next((f for f in self.files if f.file_path == file_path), None)
        if existing_file:
            logger.debug(f"文件已存在于列表中: {file_path}")
            logger.info(f"文件已存在，重新选中以同步UI: {file_path}")
            self._select_file(existing_file)
            return True, ""

        file_info = FileInfo(file_path)
        can_add, error_msg = self._can_add_category(file_info.actual_category)
        if not can_add:
            logger.warning(f"文件类别不匹配: {error_msg}")
            return False, error_msg

        # 如果是第一个文件或者自动选中
        should_select = auto_select or len(self.files) == 0
        if should_select:
            for f in self.files:
                f.is_selected = False
            file_info.is_selected = True
            self.selected_file = file_info
            logger.debug(f"自动选中文件: {file_path}")

        self.files.append(file_info)
        logger.info(f"文件已添加到批量列表: {file_path}")

        # 更新显示
        self._show_file_list()

        # 触发选中回调
        if should_select and self.on_file_selected:
            logger.debug(f"触发文件选中回调: {file_path}")
            self.on_file_selected(file_info)

        return True, ""

    def _determine_category_by_majority(self, file_paths: list[str]) -> tuple[str, list[str], list[str]]:
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
        category_count = {"text": 0, "document": 0, "spreadsheet": 0, "layout": 0, "image": 0}
        categorized_files = {"text": [], "document": [], "spreadsheet": [], "layout": [], "image": []}

        for file_path in file_paths:
            category = get_actual_file_category(file_path)
            if category in category_count:
                category_count[category] += 1
                categorized_files[category].append(file_path)

        non_zero_categories = {cat: count for cat, count in category_count.items() if count > 0}
        if not non_zero_categories:
            selected_category = "text"
        else:
            selected_category = max(non_zero_categories, key=lambda k: non_zero_categories[k])

        matching_files = categorized_files[selected_category]
        non_matching_files = []
        for category, files in categorized_files.items():
            if category != selected_category and files:
                non_matching_files.extend(files)

        logger.info(
            f"根据多数原则选定类别: {selected_category}, 匹配文件: {len(matching_files)}, 不匹配文件: {len(non_matching_files)}"
        )
        return selected_category, matching_files, non_matching_files

    def add_files(
        self, file_paths: list[str], auto_select_first: bool = False, known_category: str | None = None
    ) -> tuple[list[str], list[tuple[str, str]]]:
        """
        批量添加多个文件到列表

        智能处理逻辑：
        - 列表为空：使用多数决策算法确定类别
        - 列表不为空：只添加匹配当前类别的文件

        参数:
            file_paths: 要添加的文件路径列表
            auto_select_first: 是否自动选中第一个成功添加的文件
            known_category: 已知的文件类别（仅列表为空时生效，跳过多数决策）

        返回:
            Tuple[List[str], List[Tuple[str, str]]]:
                - 成功添加的文件路径列表
                - 失败的文件列表
        """
        logger.debug(f"批量添加文件到列表: {len(file_paths)} 个文件")

        added_files = []
        failed_files = []

        if not self.files:
            if known_category:
                self.current_category = known_category
                logger.info(f"列表为空，使用已知类别: {known_category}")
                matching_files, non_matching_files = file_paths, []
                selected_category = known_category
            else:
                logger.info("列表为空，使用智能类别决策")
                selected_category, matching_files, non_matching_files = self._determine_category_by_majority(file_paths)
                self.current_category = selected_category
                logger.info(f"设置当前类别为: {selected_category}")

            for i, file_path in enumerate(matching_files):
                auto_select = auto_select_first and i == 0
                success, error_msg = self.add_file(file_path, auto_select)
                if success:
                    added_files.append(file_path)
                else:
                    failed_files.append((file_path, error_msg))

            for file_path in non_matching_files:
                category_names = {
                    "text": "文本类",
                    "document": "文档类",
                    "spreadsheet": "表格类",
                    "layout": "版式类",
                    "image": "图片类",
                }
                selected_name = category_names.get(selected_category, selected_category)
                file_category = get_actual_file_category(file_path)
                file_category_name = category_names.get(file_category, file_category)
                failed_files.append(
                    (
                        file_path,
                        t(
                            "messages.category_mismatch_detail",
                            selected_name=selected_name,
                            file_category_name=file_category_name,
                        ),
                    )
                )

        else:
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
        self.current_category = None
        self.selected_file = None
        self._show_empty_state()

        logger.info("批量文件列表已清空")

    def get_files(self) -> list[str]:
        """获取所有文件路径"""
        return [f.file_path for f in self.files]

    def get_file_objects(self) -> list[FileInfo]:
        """获取所有文件对象"""
        return self.files.copy()

    def get_selected_file(self) -> FileInfo | None:
        """获取当前选中的文件"""
        return self.selected_file

    def get_selected_file_path(self) -> str | None:
        """获取当前选中文件的路径"""
        return self.selected_file.file_path if self.selected_file else None

    def get_file_count(self) -> int:
        """获取文件数量"""
        return len(self.files)

    def get_file_by_path(self, file_path: str) -> FileInfo | None:
        """
        根据文件路径获取FileInfo对象

        参数:
            file_path: 文件路径

        返回:
            Optional[FileInfo]: FileInfo对象或None
        """
        for file_info in self.files:
            if file_info.file_path == file_path:
                return file_info
        return None

    def has_files(self) -> bool:
        """检查是否有文件"""
        return len(self.files) > 0

    def get_current_category(self) -> str | None:
        """获取当前文件类别"""
        return self.current_category

    def update_file_status(
        self,
        file_path: str,
        status: str,
        output_path: str | None = None,
        skip_reason: str | None = None,
        error_message: str | None = None,
    ):
        """
        更新文件的处理状态

        参数:
            file_path: 原始文件路径
            status: 新的处理状态
            output_path: 输出文件路径（status为completed时提供）
            skip_reason: 跳过原因（status为skipped时提供）
            error_message: 错误信息（status为failed时提供）
        """
        logger.debug(f"更新文件状态: {file_path} -> {status}")

        for file_info in self.files:
            if file_info.file_path == file_path:
                file_info.status = status

                if status == "completed" and output_path:
                    file_info.output_path = output_path
                elif status == "skipped" and skip_reason:
                    file_info.skip_reason = skip_reason
                elif status == "failed" and error_message:
                    file_info.error_message = error_message

                logger.info(f"文件状态已更新: {file_path} -> {status}")
                break

        self._show_file_list()

    def keep_first_file_only(self):
        """只保留第一个文件（用于模式切换）"""
        logger.debug("只保留第一个文件")

        if len(self.files) > 1:
            first_file = self.files[0]
            self.files = [first_file]
            first_file.is_selected = True
            self.selected_file = first_file
            self._show_file_list()
            logger.info(f"已保留第一个文件: {first_file.file_path}")
        elif len(self.files) == 0:
            logger.debug("列表为空，无需操作")
        else:
            logger.debug("列表只有一个文件，无需操作")

    def add_file_with_mode(self, file_path: str, mode: str = "batch", auto_select: bool = True) -> tuple[bool, str]:
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

        if mode == "single":
            self.files.clear()
            self.current_category = None
            self.selected_file = None
            logger.debug("单文件模式：已清空列表")

        return self.add_file(file_path, auto_select)
