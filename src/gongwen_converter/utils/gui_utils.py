"""
GUI相关的工具函数
"""

import tkinter as tk

from gongwen_converter.utils.font_utils import get_small_font


def show_info_dialog(title: str, message: str, alert: bool = False, parent=None):
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
                temp_root.geometry(f"+{screen_width//2}+{screen_height//2}")
            except Exception:
                pass
        
        dialog = MessageDialog(
            parent=dialog_parent,
            title=title,
            message=message,
            alert=alert,
            localize=False # 确保按钮文本是标准英文
        )
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
            print(f"[{'警告' if alert else '信息'}] {title}: {message}")


def show_error_dialog(title: str, message: str):
    """
    显示错误对话框（现在是 show_info_dialog 的一个包装）
    """
    show_info_dialog(title, message, alert=True)


class ToolTip:
    """
    工具提示类，用于在鼠标悬停时显示提示信息
    """
    
    def __init__(self, widget, text: str = '', delay: int = 500):
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
        x = self.widget.winfo_rootx() + self.x + 20
        y = self.widget.winfo_rooty() + self.y + 10
        
        # 创建提示窗口
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        
        try:
            # 尝试使用ttkbootstrap样式
            import ttkbootstrap as tb
            label = tb.Label(
                tw, 
                text=self.text, 
                justify='left',
                background="white",
                relief='solid', 
                borderwidth=1,
                font=(small_font_name, small_font_size),
                padding=(8, 6)  # 增加内边距
            )
        except ImportError:
            # 回退到标准tkinter
            label = tk.Label(
                tw, 
                text=self.text, 
                justify='left',
                background="white",
                relief='solid', 
                borderwidth=1,
                font=(small_font_name, small_font_size),
                padx=8,  # 水平内边距
                pady=6   # 垂直内边距
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
    
    def __init__(self, widget, text: str = '', condition_func=None, delay: int = 500):
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


def create_info_icon(parent, tooltip_text: str, bootstyle: str = "secondary") -> tk.Label:
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
            font=(small_font_name, small_font_size, "bold")
        )
    except ImportError:
        info_label = tk.Label(
            parent,
            text="ⓘ",
            cursor="question_arrow",
            font=(small_font_name, small_font_size, "bold")
        )
    
    # 绑定工具提示
    ToolTip(info_label, tooltip_text)
    
    return info_label
