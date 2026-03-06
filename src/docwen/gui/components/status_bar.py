"""
状态栏组件模块
实现可滚动的消息历史记录显示，支持时间戳、颜色编码和定位按钮
"""

import contextlib
import logging
import threading
import time
import tkinter as tk
from collections.abc import Callable
from datetime import datetime
from pathlib import Path

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.i18n import t
from docwen.utils.dpi_utils import scale
from docwen.utils.font_utils import get_default_font, get_small_font
from docwen.utils.gui_utils import bind_label_wraplength_to_container
from docwen.utils.icon_utils import load_image_icon

logger = logging.getLogger(__name__)

TRANSIENT_PRIORITY = {"error": 0, "progress": 1, "processing": 2}


class StatusBar(tb.Frame):
    """
    可滚动的状态栏组件

    功能特性：
    - 支持最多100条消息历史记录（超出上限时按策略移除旧消息）
    - 可滚动查看历史消息（始终显示滚动条）
    - 自动占满剩余垂直空间
    - 支持时间戳、颜色编码和定位按钮
    - 新消息添加时自动滚动到底部

    技术实现：
    - 使用 list 存储消息历史，并在追加时手动裁剪到上限
    - 使用 Canvas + Scrollbar 实现可滚动容器
    - 追加消息时增量创建控件，超限时移除最旧行控件
    """

    def __init__(self, master, on_location_clicked: Callable | None = None, **kwargs):
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

        # 初始化消息历史队列（使用 list + 手动限制长度）
        # 每个元素是一个字典: {"message": str, "type": str, "show_location": bool, "file_path": str}
        self.message_history = []
        logger.debug(f"消息队列初始化完成，最大容量: {self.max_messages}")

        self._history_row_widgets = []

        self.transient_messages = {}
        self._transient_current_key = None
        self._transient_label = None
        self._transient_refresh_scheduled = False
        self._transient_refresh_after_id = None
        self._scroll_to_bottom_scheduled = False
        self._scroll_to_bottom_after_id = None
        self._scroll_to_bottom_requested = False
        self._scrollregion_refresh_scheduled = False
        self._scrollregion_refresh_after_id = None
        self._last_canvas_width = None
        self._is_destroyed = False
        self._last_message_signature = None
        self._last_message_time = 0.0

        # 创建滚动容器和界面元素
        self._create_scrollable_container()

        self.bind("<Destroy>", self._on_destroy, add="+")

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
        self.grid_rowconfigure(1, weight=0)  # 临时状态行固定高度
        self.grid_columnconfigure(0, weight=1)  # Canvas列可扩展
        self.grid_columnconfigure(1, weight=0)  # Scrollbar列固定宽度

        # 创建Canvas（用于滚动）
        self.canvas = tb.Canvas(self, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        logger.debug("Canvas创建完成")

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

        # 创建垂直滚动条（始终显示，使用带边界检查的命令）
        self.scrollbar = tb.Scrollbar(self, orient="vertical", command=bounded_yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        logger.debug("滚动条创建完成（始终显示）")

        # 创建滚动框架（实际的消息容器）
        self.scrollable_frame = tb.Frame(self.canvas)
        self.scrollable_frame.grid_columnconfigure(0, weight=1)  # 消息列可扩展
        logger.debug("滚动框架创建完成")

        # 将滚动框架放入Canvas
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        # 配置Canvas的滚动命令
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # 绑定事件：当滚动框架大小改变时更新滚动区域
        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)

        # 绑定事件：当Canvas大小改变时调整滚动框架宽度
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # 绑定鼠标滚轮事件
        self._bind_mousewheel_events()

        self.transient_frame = tb.Frame(self)
        self.transient_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.transient_frame.grid_columnconfigure(0, weight=1)
        self._transient_label = tb.Label(
            self.transient_frame,
            text="",
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            anchor="w",
            justify="left",
            wraplength=scale(360),
            padding=(scale(5), scale(2)),
        )
        self._transient_label.grid(row=0, column=0, sticky="ew")
        self.transient_frame.grid_remove()
        bind_label_wraplength_to_container(
            self._transient_label, self.transient_frame, min_wraplength=scale(160), padding=scale(20)
        )

        logger.debug("可滚动容器创建完成")

    def _on_destroy(self, event):
        if event.widget is not self:
            return

        self._is_destroyed = True

        if self._transient_refresh_after_id:
            with contextlib.suppress(Exception):
                self.after_cancel(self._transient_refresh_after_id)
            self._transient_refresh_after_id = None
        self._transient_refresh_scheduled = False

        if self._scroll_to_bottom_after_id:
            with contextlib.suppress(Exception):
                self.after_cancel(self._scroll_to_bottom_after_id)
            self._scroll_to_bottom_after_id = None
        self._scroll_to_bottom_scheduled = False

        if self._scrollregion_refresh_after_id:
            with contextlib.suppress(Exception):
                self.after_cancel(self._scrollregion_refresh_after_id)
            self._scrollregion_refresh_after_id = None
        self._scrollregion_refresh_scheduled = False

        with contextlib.suppress(Exception):
            self.clear_all_transient_messages()

    def set_transient_message(self, key: str, message: str, message_type: str = "secondary", ttl_ms: int | None = None):
        if threading.current_thread() is not threading.main_thread():
            logger.error("非主线程调用 set_transient_message，已忽略")
            return
        priority = TRANSIENT_PRIORITY.get(key, 99)
        now = time.monotonic()

        previous = self.transient_messages.get(key)
        version = (previous.get("version", 0) + 1) if previous else 1
        if previous and previous.get("after_id"):
            with contextlib.suppress(Exception):
                self.after_cancel(previous["after_id"])

        after_id = None
        if ttl_ms is not None:
            ttl_value = None
            try:
                ttl_value = int(ttl_ms)
            except Exception:
                ttl_value = None

            if ttl_value is not None:
                try:
                    if ttl_value <= 0:
                        after_id = self.after_idle(lambda k=key, v=version: self._expire_transient_message(k, v))
                    else:
                        after_id = self.after(ttl_value, lambda k=key, v=version: self._expire_transient_message(k, v))
                except tk.TclError:
                    after_id = None

        self.transient_messages[key] = {
            "key": key,
            "message": message,
            "type": message_type,
            "priority": priority,
            "updated_at": now,
            "version": version,
            "after_id": after_id,
        }
        self._schedule_transient_refresh()

    def clear_transient_message(self, key: str):
        entry = self.transient_messages.get(key)
        if not entry:
            return
        if entry.get("after_id"):
            with contextlib.suppress(Exception):
                self.after_cancel(entry["after_id"])
        self.transient_messages.pop(key, None)
        self._schedule_transient_refresh()

    def clear_all_transient_messages(self):
        for entry in list(self.transient_messages.values()):
            if entry.get("after_id"):
                with contextlib.suppress(Exception):
                    self.after_cancel(entry["after_id"])
        self.transient_messages.clear()
        self._schedule_transient_refresh()

    def _expire_transient_message(self, key: str, version: int):
        entry = self.transient_messages.get(key)
        if not entry or entry.get("version") != version:
            return
        self.transient_messages.pop(key, None)
        self._schedule_transient_refresh()

    def _schedule_transient_refresh(self):
        if self._is_destroyed:
            return
        if self._transient_refresh_scheduled:
            return
        self._transient_refresh_scheduled = True
        try:
            self._transient_refresh_after_id = self.after_idle(self._run_transient_refresh)
        except tk.TclError:
            self._transient_refresh_after_id = None
            self._transient_refresh_scheduled = False

    def _run_transient_refresh(self):
        self._transient_refresh_after_id = None
        self._transient_refresh_scheduled = False
        if self._is_destroyed:
            return
        self._refresh_transient_ui()

    def _schedule_scroll_to_bottom(self):
        if self._is_destroyed:
            return
        if self._scroll_to_bottom_scheduled:
            return
        self._scroll_to_bottom_scheduled = True
        try:
            self._scroll_to_bottom_after_id = self.after_idle(self._scroll_to_bottom)
        except tk.TclError:
            self._scroll_to_bottom_after_id = None
            self._scroll_to_bottom_scheduled = False

    def _schedule_scrollregion_refresh(self):
        if self._is_destroyed:
            return
        if self._scrollregion_refresh_scheduled:
            return
        self._scrollregion_refresh_scheduled = True
        try:
            self._scrollregion_refresh_after_id = self.after_idle(self._refresh_scrollregion)
        except tk.TclError:
            self._scrollregion_refresh_after_id = None
            self._scrollregion_refresh_scheduled = False

    def _scroll_to_bottom(self):
        self._scroll_to_bottom_after_id = None
        self._scroll_to_bottom_scheduled = False
        if self._is_destroyed:
            return
        self._scroll_to_bottom_requested = True
        try:
            if not self.canvas or not self.canvas.winfo_exists():
                return
            self.update_idletasks()
            self._refresh_scrollregion()
        except (tk.TclError, Exception):
            pass

    def _refresh_transient_ui(self):
        if self._is_destroyed:
            return
        if not self._transient_label or not self._transient_label.winfo_exists():
            return

        if not self.transient_messages:
            self._transient_current_key = None
            self._transient_label.configure(text="")
            self.transient_frame.grid_remove()
            return

        active = sorted(self.transient_messages.values(), key=lambda m: (m["priority"], -m["updated_at"]))[0]

        if self._transient_current_key != active["key"]:
            self._transient_current_key = active["key"]

        self._transient_label.configure(text=active["message"], bootstyle=active["type"])
        self.transient_frame.grid()

    def _on_frame_configure(self, event=None):
        """
        滚动框架配置改变事件处理
        更新Canvas的滚动区域以匹配框架大小
        """
        self._schedule_scrollregion_refresh()

    def _on_canvas_configure(self, event):
        """
        Canvas配置改变事件处理
        调整滚动框架宽度以匹配Canvas宽度
        """
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_frame, width=canvas_width)
        if self._last_canvas_width == canvas_width:
            return
        self._last_canvas_width = canvas_width
        self._update_message_wraplengths(canvas_width)

    def _get_message_wraplength(self, canvas_width: int, has_location: bool) -> int:
        min_wrap = scale(160)
        reserved = scale(20)
        if has_location:
            reserved += scale(40)
        return max(min_wrap, canvas_width - reserved)

    def _update_message_wraplengths(self, canvas_width: int):
        try:
            if self._transient_label and self._transient_label.winfo_exists():
                self._transient_label.configure(wraplength=max(scale(160), canvas_width - scale(20)))
        except tk.TclError:
            pass

        for row in self._history_row_widgets:
            widgets = row.get("widgets", [])
            if not widgets:
                continue
            message_label = widgets[0]
            try:
                if not message_label.winfo_exists():
                    continue
                has_location = len(widgets) > 1
                message_label.configure(wraplength=self._get_message_wraplength(canvas_width, has_location))
            except tk.TclError:
                continue

    def _bind_mousewheel_events(self):
        """
        绑定鼠标滚轮事件到所有相关组件
        支持Windows、Linux和macOS系统
        """

        def on_mousewheel(event):
            yview = self.canvas.yview()
            try:
                windowing_system = self.tk.call("tk", "windowingsystem")
            except Exception:
                windowing_system = ""

            delta = int(-1 * event.delta) if windowing_system == "aqua" else int(-1 * (event.delta / 120))

            if delta == 0:
                return

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

    def add_message(
        self,
        message: str,
        message_type: str = "secondary",
        show_location: bool = False,
        file_path: str | None = None,
    ):
        """
        添加状态消息到历史记录

        新消息会添加到列表底部，并自动滚动到最新消息。
        如果消息数量超过上限，会按裁剪策略移除旧消息。

        参数:
            message: 状态消息文本
            message_type: 消息类型/颜色 ('success', 'danger', 'warning', 'secondary', 'info')
            show_location: 是否显示定位按钮
            file_path: 关联的文件路径（用于定位按钮）
        """
        if threading.current_thread() is not threading.main_thread():
            logger.error("非主线程调用 add_message，已忽略")
            return
        logger.debug(f"添加状态消息: {message} (类型: {message_type}, 定位: {show_location})")

        signature = (message, message_type, bool(show_location), file_path or "")
        now = time.monotonic()
        if signature == self._last_message_signature and (now - self._last_message_time) < 0.25:
            return
        self._last_message_signature = signature
        self._last_message_time = now

        try:
            # 获取当前时间并格式化为"HH:MM:SS"
            current_time = datetime.now().strftime("%H:%M:%S")
            timestamped_message = f"{current_time} {message}"

            # 添加到消息历史队列（达到上限时会触发裁剪）
            message_data = {
                "message": timestamped_message,
                "type": message_type,
                "show_location": show_location,
                "file_path": file_path,
            }
            self.message_history.append(message_data)
            logger.debug(f"消息已添加到历史队列，当前队列长度: {len(self.message_history)}")
            self._append_message_row(message_data)
            self._enforce_message_limit()
            self._scroll_to_bottom_requested = True
            self._schedule_scrollregion_refresh()
            self._schedule_scroll_to_bottom()

        except Exception as e:
            logger.error(f"添加状态消息失败: {e!s}", exc_info=True)

    def _enforce_message_limit(self):
        important_types = {"success", "danger", "warning"}

        while len(self.message_history) > self.max_messages:
            removal_index = None
            for i, msg in enumerate(self.message_history):
                if msg.get("type") not in important_types:
                    removal_index = i
                    break

            if removal_index is None:
                removal_index = 0

            self._remove_history_row_at(removal_index)

    def _remove_history_row_at(self, index: int):
        if index < 0 or index >= len(self.message_history):
            return
        self.message_history.pop(index)

        if index < len(self._history_row_widgets):
            removed = self._history_row_widgets.pop(index)
            for widget in removed.get("widgets", []):
                with contextlib.suppress(Exception):
                    widget.destroy()

        for row_index in range(index, len(self._history_row_widgets)):
            row = self._history_row_widgets[row_index]
            for widget in row.get("widgets", []):
                with contextlib.suppress(Exception):
                    widget.grid_configure(row=row_index)

        with contextlib.suppress(tk.TclError):
            self.after_idle(self._refresh_scrollregion)

    def _refresh_scrollregion(self):
        self._scrollregion_refresh_after_id = None
        self._scrollregion_refresh_scheduled = False
        if self._is_destroyed:
            return
        try:
            if not self.canvas or not self.canvas.winfo_exists():
                return
            bbox = self.canvas.bbox("all")
            if bbox:
                self.canvas.configure(scrollregion=bbox)
                if self._scroll_to_bottom_requested:
                    self.canvas.yview_moveto(1.0)
                    self._scroll_to_bottom_requested = False
                    logger.debug("已自动滚动到最新消息")
        except tk.TclError:
            pass

    def _render_all_messages(self):
        """
        重新渲染所有消息

        清空当前显示，然后根据消息历史队列重新创建所有Label和按钮
        """
        logger.debug("开始重新渲染所有消息")

        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self._history_row_widgets.clear()

        for msg_data in self.message_history:
            self._append_message_row(msg_data)

        logger.debug(f"消息渲染完成，共 {len(self.message_history)} 条")

    def _append_message_row(self, msg_data: dict):
        row_index = len(self._history_row_widgets)
        widgets = []

        canvas_width = 0
        try:
            if self.canvas and self.canvas.winfo_exists():
                canvas_width = int(self.canvas.winfo_width())
        except Exception:
            canvas_width = 0

        wraplength = (
            self._get_message_wraplength(canvas_width, bool(msg_data["show_location"]))
            if canvas_width > 0
            else scale(360)
        )
        message_label = tb.Label(
            self.scrollable_frame,
            text=msg_data["message"],
            font=(self.small_font, self.small_size),
            bootstyle=msg_data["type"],
            anchor="w",
            justify="left",
            wraplength=wraplength,
            padding=(scale(5), scale(2)),
        )
        message_label.grid(row=row_index, column=0, sticky="ew", padx=(0, scale(5)))
        widgets.append(message_label)

        location_icon = None
        if msg_data["show_location"]:
            icon_size = scale(16)
            location_icon = load_image_icon(
                "location_icon.png", master=self.scrollable_frame, size=(icon_size, icon_size)
            )

            location_button = tb.Button(
                self.scrollable_frame,
                image=location_icon,
                bootstyle="secondary-link",
                command=lambda fp=msg_data["file_path"]: self._on_location_button_clicked(fp),
            )
            if location_icon:
                location_button.image = location_icon
            location_button.grid(row=row_index, column=1, sticky="e", padx=(0, scale(5)))
            location_button.bind(
                "<Button-3>", lambda event, fp=msg_data["file_path"]: self._show_location_context_menu(event, fp)
            )
            widgets.append(location_button)
            logger.debug(f"第 {row_index} 行已添加定位按钮")

        self._history_row_widgets.append({"widgets": widgets, "icon": location_icon})

    def _show_location_context_menu(self, event, file_path: str | None):
        menu = tk.Menu(self, tearoff=0)
        if file_path:
            menu.add_command(label=t("status_bar.copy_path"), command=lambda: self._copy_to_clipboard(file_path))
        else:
            menu.add_command(label=t("status_bar.copy_path"), state="disabled")
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            with contextlib.suppress(Exception):
                menu.grab_release()

    def _copy_to_clipboard(self, text: str):
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.set_transient_message("progress", t("status_bar.path_copied"), "info", ttl_ms=3000)
        except Exception:
            self.set_transient_message("error", t("status_bar.copy_failed"), "warning", ttl_ms=3000)

    def _create_message_row(self, row_index: int, msg_data: dict):
        """
        创建单条消息的显示行

        参数:
            row_index: 行索引（用于grid布局）
            msg_data: 消息数据字典
        """
        self._append_message_row(msg_data)

    def clear_all(self):
        """
        清空所有消息历史记录
        """
        logger.debug("清空所有消息历史")

        try:
            # 清空消息队列
            self.message_history.clear()

            self.clear_all_transient_messages()

            # 清空显示
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self._history_row_widgets.clear()

            logger.info("状态栏已清空")

        except Exception as e:
            logger.error(f"清空状态栏失败: {e!s}", exc_info=True)

    def _on_location_button_clicked(self, file_path: str | None):
        """
        处理定位按钮点击事件

        参数:
            file_path: 要定位的文件路径
        """
        logger.debug(f"定位按钮被点击，文件路径: {file_path}")

        if not file_path:
            self.set_transient_message("error", t("status_bar.cannot_locate_no_path"), "warning", ttl_ms=8000)
            return

        if not Path(file_path).exists():
            self.set_transient_message("error", t("status_bar.file_missing", path=file_path), "warning", ttl_ms=8000)
            return

        if self.on_location_clicked:
            logger.info(f"执行定位操作: {file_path}")
            self.on_location_clicked(file_path)
        else:
            self.set_transient_message("error", t("status_bar.cannot_locate_no_callback"), "warning", ttl_ms=8000)

    def refresh_style(self):
        """
        刷新组件样式以匹配当前主题
        主题切换时调用此方法
        """
        logger.debug("刷新状态栏组件样式")

        # 重新渲染所有消息以应用新主题
        self._render_all_messages()

        logger.debug("状态栏样式刷新完成")
