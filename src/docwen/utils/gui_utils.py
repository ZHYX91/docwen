"""
GUI相关的工具函数
"""

import contextlib
import logging
import tkinter as tk
import tkinter.font as tkfont
from typing import Any

from docwen.utils.dpi_utils import scale
from docwen.utils.font_utils import get_small_font

logger = logging.getLogger(__name__)


def show_info_dialog(title: str, message: str, alert: bool = False, parent=None, localize: bool = True):
    """
    通用对话框显示函数

    参数:
        title: 对话框标题
        message: 消息内容
        alert: 是否使用警告样式 (True for warning, False for info)
        parent: 父窗口（可选），用于定位对话框。如果提供父窗口，对话框将显示在父窗口中央
    """
    try:
        import ttkbootstrap as tb
        from ttkbootstrap.dialogs import MessageDialog

        # 如果提供了父窗口，使用父窗口；否则创建临时窗口
        temp_root = None
        if parent is not None:
            dialog_parent = parent
        else:
            # 创建临时窗口，但尝试获取活动窗口的位置
            temp_root = tb.Window()
            temp_root.withdraw()
            dialog_parent = temp_root

            # 尝试将临时窗口放置在屏幕中央
            try:
                screen_width = temp_root.winfo_screenwidth()
                screen_height = temp_root.winfo_screenheight()
                temp_root.geometry(f"+{screen_width // 2}+{screen_height // 2}")
            except Exception:
                pass

        dialog = MessageDialog(parent=dialog_parent, title=title, message=message, alert=alert, localize=localize)
        dialog.show()

        # 只在创建了临时窗口时销毁它
        if temp_root is not None:
            temp_root.destroy()

    except Exception:
        try:
            import tkinter as tk
            from tkinter import messagebox

            # 使用parent或创建临时root
            if parent is not None:
                if alert:
                    messagebox.showwarning(title, message, parent=parent)
                else:
                    messagebox.showinfo(title, message, parent=parent)
            else:
                root = tk.Tk()
                root.withdraw()
                if alert:
                    messagebox.showwarning(title, message)
                else:
                    messagebox.showinfo(title, message)
                root.destroy()
        except Exception:
            logger.warning(
                "无法显示GUI对话框，已降级为日志输出: [%s] %s: %s",
                "警告" if alert else "信息",
                title,
                message,
            )


def show_error_dialog(title: str, message: str):
    """
    显示错误对话框（现在是 show_info_dialog 的一个包装）
    """
    show_info_dialog(title, message, alert=True)


class ToolTip:
    """
    工具提示类，用于在鼠标悬停时显示提示信息
    """

    def __init__(self, widget, text: str = "", delay: int = 500):
        """
        初始化工具提示

        参数:
            widget: 要绑定工具提示的控件
            text: 提示文本
            delay: 显示延迟（毫秒）
        """
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window = None
        self.id = None
        self.x = self.y = 0

        # 绑定事件
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<Motion>", self.motion)

    def enter(self, event=None):
        """鼠标进入控件时调度显示"""
        self.schedule()

    def leave(self, event=None):
        """鼠标离开控件时取消显示"""
        self.unschedule()
        self.hidetip()

    def motion(self, event=None):
        """鼠标移动时更新位置"""
        if event is None:
            return
        self.x = event.x
        self.y = event.y
        self.unschedule()
        self.schedule()

    def schedule(self):
        """调度显示工具提示"""
        self.unschedule()
        self.id = self.widget.after(self.delay, self.showtip)

    def unschedule(self):
        """取消调度"""
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def showtip(self):
        """显示工具提示"""
        if self.tip_window or not self.text:
            return

        # 获取小字体配置
        small_font_name, small_font_size = get_small_font()

        # 计算提示窗口位置
        x = self.widget.winfo_rootx() + self.x + scale(20)
        y = self.widget.winfo_rooty() + self.y + scale(10)

        # 创建提示窗口
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        try:
            # 尝试使用ttkbootstrap样式
            import ttkbootstrap as tb

            style = tb.Style.get_instance()
            colors = getattr(style, "colors", None)
            bg = getattr(colors, "bg", "white")
            fg = getattr(colors, "fg", "black")
            label = tb.Label(
                tw,
                text=self.text,
                justify="left",
                background=bg,
                foreground=fg,
                relief="solid",
                borderwidth=1,
                font=(small_font_name, small_font_size),
                padding=(scale(8), scale(6)),
            )
        except ImportError:
            # 回退到标准tkinter
            bg = None
            fg = None
            try:
                bg = self.widget.cget("background")
            except Exception:
                bg = None
            try:
                fg = self.widget.cget("foreground")
            except Exception:
                fg = None
            label = tk.Label(
                tw,
                text=self.text,
                justify="left",
                background=bg or "white",
                foreground=fg or "black",
                relief="solid",
                borderwidth=1,
                font=(small_font_name, small_font_size),
                padx=scale(8),
                pady=scale(6),
            )

        label.pack(ipadx=1)

    def hidetip(self):
        """隐藏工具提示"""
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class ConditionalToolTip(ToolTip):
    """
    条件显示的工具提示类

    只在满足特定条件时才显示提示信息，用于动态场景
    """

    def __init__(self, widget, text: str = "", condition_func=None, delay: int = 500):
        """
        初始化条件工具提示

        参数:
            widget: 要绑定工具提示的控件
            text: 提示文本
            condition_func: 条件函数，返回True时显示tooltip，返回False时不显示
            delay: 显示延迟（毫秒）
        """
        self.condition_func = condition_func if condition_func else lambda: True
        super().__init__(widget, text, delay)

    def showtip(self):
        """只在满足条件时显示工具提示"""
        # 检查条件函数
        if not self.condition_func():
            return

        # 调用父类的showtip方法
        super().showtip()


def ellipsize_end_text(text: str, font: tkfont.Font, max_width: int) -> str:
    if max_width <= 0:
        return ""
    if font.measure(text) <= max_width:
        return text
    ellipsis = "…"
    ellipsis_width = font.measure(ellipsis)
    if ellipsis_width >= max_width:
        return ellipsis

    low = 0
    high = len(text)
    while low < high:
        mid = (low + high) // 2
        candidate = text[:mid] + ellipsis
        if font.measure(candidate) <= max_width:
            low = mid + 1
        else:
            high = mid
    keep = max(0, low - 1)
    return text[:keep] + ellipsis


def ellipsize_middle_text(text: str, font: tkfont.Font, max_width: int) -> str:
    if max_width <= 0:
        return ""
    if font.measure(text) <= max_width:
        return text
    ellipsis = "…"
    ellipsis_width = font.measure(ellipsis)
    if ellipsis_width >= max_width:
        return ellipsis

    max_side = len(text) // 2
    low = 0
    high = max_side
    while low < high:
        mid = (low + high + 1) // 2
        candidate = text[:mid] + ellipsis + text[-mid:]
        if font.measure(candidate) <= max_width:
            low = mid
        else:
            high = mid - 1
    if low <= 0:
        return ellipsis
    return text[:low] + ellipsis + text[-low:]


def bind_single_line_ellipsis(label: tk.Widget, full_text: str, padding: int = 4, mode: str = "end") -> None:
    label_any: Any = label
    label_any._is_truncated = False
    tooltip = ConditionalToolTip(label, full_text, condition_func=lambda: bool(getattr(label, "_is_truncated", False)))
    label_any._tooltip = tooltip
    mode = str(mode).lower()
    state = {"retries": 0, "after_id": None}

    def update(_event=None):
        state["after_id"] = None
        try:
            if not label.winfo_exists():
                return
            width = int(label.winfo_width())
        except Exception:
            return
        if width <= 1:
            try:
                if state["retries"] >= 10:
                    return
                if state["after_id"] is None:
                    state["retries"] += 1
                    state["after_id"] = label.after(50, update)
            except Exception:
                pass
            return
        state["retries"] = 0
        max_width = max(10, width - padding)
        font = tkfont.Font(font=label.cget("font"))
        if mode == "middle":
            truncated = ellipsize_middle_text(full_text, font, max_width)
        else:
            truncated = ellipsize_end_text(full_text, font, max_width)
        label_any._is_truncated = truncated != full_text
        try:
            if label.cget("text") != truncated:
                label_any.configure(text=truncated)
        except Exception:
            pass

    label.bind("<Configure>", update, add="+")
    with contextlib.suppress(Exception):
        label.after_idle(update)


def bind_label_wraplength_to_container(
    label: tk.Widget, container: tk.Widget, min_wraplength: int = 160, padding: int = 20
) -> None:
    state = {"retries": 0, "after_id": None}

    def update(_event=None):
        state["after_id"] = None
        try:
            if not container.winfo_exists() or not label.winfo_exists():
                return
            width = int(container.winfo_width())
        except Exception:
            return
        if width <= 1:
            try:
                if state["retries"] >= 10:
                    return
                if state["after_id"] is None:
                    state["retries"] += 1
                    state["after_id"] = label.after(50, update)
            except Exception:
                pass
            return
        state["retries"] = 0
        wraplength = max(int(min_wraplength), width - int(padding))
        try:
            label_any: Any = label
            label_any.configure(wraplength=wraplength)
        except Exception:
            pass

    container.bind("<Configure>", update, add="+")
    with contextlib.suppress(Exception):
        label.after_idle(update)


def create_info_icon(parent, tooltip_text: str, bootstyle: str = "secondary") -> Any:
    """
    创建信息图标并绑定工具提示

    参数:
        parent: 父控件
        tooltip_text: 工具提示文本
        bootstyle: ttkbootstrap样式

    返回:
        tk.Label: 信息图标标签
    """
    # 获取小字体配置
    small_font_name, small_font_size = get_small_font()

    try:
        import ttkbootstrap as tb

        info_label = tb.Label(
            parent,
            text="ⓘ",  # 使用信息符号
            bootstyle=bootstyle,
            cursor="question_arrow",
            font=(small_font_name, small_font_size, "bold"),
        )
    except ImportError:
        info_label = tk.Label(
            parent, text="ⓘ", cursor="question_arrow", font=(small_font_name, small_font_size, "bold")
        )

    # 绑定工具提示
    ToolTip(info_label, tooltip_text)

    return info_label
