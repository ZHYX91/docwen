"""
对话框基类 - 提供统一的对话框基础功能

提供所有自定义对话框的基类，包括：
- 自动继承应用程序图标
- 模态对话框支持
- 居中显示功能
- DPI缩放支持

同时提供MessageBox包装类，统一消息框接口。
"""

import logging
import ttkbootstrap as tb
from tkinter import messagebox

from docwen.utils.dpi_utils import ScalableMixin

logger = logging.getLogger(__name__)


class BaseDialog(tb.Toplevel, ScalableMixin):
    """
    对话框基类
    
    为所有自定义对话框提供统一的基础功能：
    - 自动继承父窗口的应用程序图标
    - 支持模态和非模态对话框
    - 提供居中显示方法
    - 继承DPI缩放功能
    
    所有自定义对话框都应该继承此基类以保持一致性。
    """
    
    def __init__(self, parent, title="", modal=True, **kwargs):
        """
        初始化对话框
        
        参数:
            parent: 父窗口对象
            title: 对话框标题，默认为空字符串
            modal: 是否为模态对话框，True时会阻塞父窗口
            **kwargs: 传递给Toplevel的其他参数
        """
        super().__init__(parent, **kwargs)
        
        self.parent = parent
        self.title(title)
        
        # 图标会自动从应用级默认图标继承，不需要手动设置
        
        # 设置为模态对话框（如果需要）
        if modal:
            self.transient(parent)
            self.grab_set()
        
        logger.debug(f"创建对话框: {title}")
    
    def center_on_parent(self):
        """
        在父窗口中心位置显示对话框
        
        计算父窗口的位置和大小,将对话框定位到父窗口的中心位置。
        通常在对话框创建完成后调用此方法。
        """
        # 确保窗口已经映射到屏幕
        self.update_idletasks()
        
        # 使用 winfo_rootx/rooty 获取相对于屏幕的绝对坐标（更可靠）
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        # 再次更新确保获取到正确的对话框尺寸
        self.update_idletasks()
        dialog_width = self.winfo_width()
        dialog_height = self.winfo_height()
        
        # 计算居中位置
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        # 确保对话框不会超出屏幕范围
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # 调整位置确保在屏幕内
        x = max(0, min(x, screen_width - dialog_width))
        y = max(0, min(y, screen_height - dialog_height))
        
        # 设置对话框位置
        self.geometry(f"+{x}+{y}")
        
        logger.debug(f"父窗口位置: ({parent_x}, {parent_y}), 尺寸: {parent_width}x{parent_height}")
        logger.debug(f"对话框尺寸: {dialog_width}x{dialog_height}")
        logger.debug(f"对话框居中位置: ({x}, {y})")


class MessageBox:
    """
    消息框包装类
    
    为tkinter.messagebox提供统一的静态方法接口：
    - showinfo: 显示信息对话框
    - showerror: 显示错误对话框
    - showwarning: 显示警告对话框
    - askyesno: 显示是/否确认对话框
    - askokcancel: 显示确定/取消对话框
    - askretrycancel: 显示重试/取消对话框
    
    所有消息框会自动继承应用程序的默认图标。
    """
    
    @staticmethod
    def showinfo(title, message, parent=None, **kwargs):
        """
        显示信息对话框
        
        参数:
            title: 对话框标题
            message: 要显示的信息内容
            parent: 父窗口对象（可选，用于定位对话框）
            **kwargs: 传递给messagebox.showinfo的其他参数
            
        返回:
            str: 用户点击的按钮名称
        """
        logger.debug(f"显示信息对话框: {title}")
        return messagebox.showinfo(title, message, parent=parent, **kwargs)
    
    @staticmethod
    def showerror(title, message, parent=None, **kwargs):
        """
        显示错误对话框
        
        参数:
            title: 对话框标题
            message: 要显示的错误消息内容
            parent: 父窗口对象（可选，用于定位对话框）
            **kwargs: 传递给messagebox.showerror的其他参数
            
        返回:
            str: 用户点击的按钮名称
        """
        logger.debug(f"显示错误对话框: {title}")
        return messagebox.showerror(title, message, parent=parent, **kwargs)
    
    @staticmethod
    def showwarning(title, message, parent=None, **kwargs):
        """
        显示警告对话框
        
        参数:
            title: 对话框标题
            message: 要显示的警告消息内容
            parent: 父窗口对象（可选，用于定位对话框）
            **kwargs: 传递给messagebox.showwarning的其他参数
            
        返回:
            str: 用户点击的按钮名称
        """
        logger.debug(f"显示警告对话框: {title}")
        return messagebox.showwarning(title, message, parent=parent, **kwargs)
    
    @staticmethod
    def askyesno(title, message, parent=None, **kwargs):
        """
        显示是/否确认对话框
        
        参数:
            title: 对话框标题
            message: 要确认的消息内容
            parent: 父窗口对象（可选，用于定位对话框）
            **kwargs: 传递给messagebox.askyesno的其他参数
            
        返回:
            bool: 用户点击"是"返回True，点击"否"返回False
        """
        logger.debug(f"显示确认对话框: {title}")
        return messagebox.askyesno(title, message, parent=parent, **kwargs)
    
    @staticmethod
    def askokcancel(title, message, parent=None, **kwargs):
        """
        显示确定/取消对话框
        
        参数:
            title: 对话框标题
            message: 要显示的消息内容
            parent: 父窗口对象（可选，用于定位对话框）
            **kwargs: 传递给messagebox.askokcancel的其他参数
            
        返回:
            bool: 用户点击"确定"返回True，点击"取消"返回False
        """
        logger.debug(f"显示确定/取消对话框: {title}")
        return messagebox.askokcancel(title, message, parent=parent, **kwargs)
    
    @staticmethod
    def askretrycancel(title, message, parent=None, **kwargs):
        """
        显示重试/取消对话框
        
        参数:
            title: 对话框标题
            message: 要显示的消息内容
            parent: 父窗口对象（可选，用于定位对话框）
            **kwargs: 传递给messagebox.askretrycancel的其他参数
            
        返回:
            bool: 用户点击"重试"返回True，点击"取消"返回False
        """
        logger.debug(f"显示重试/取消对话框: {title}")
        return messagebox.askretrycancel(title, message, parent=parent, **kwargs)
