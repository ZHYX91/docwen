"""
设置选项卡基类模块

本模块提供所有设置选项卡的抽象基类，实现了以下功能：
1. 统一的初始化流程（模板方法模式）
2. 可滚动容器的创建和管理
3. 丰富的UI组件辅助方法
4. 标准的接口规范

设计模式：
- 模板方法模式：定义标准初始化流程
- 抽象基类：强制子类实现必要方法

使用示例：
    class MyTab(BaseSettingsTab):
        def _create_interface(self):
            frame = self.create_section_frame(
                self.scrollable_frame,
                "我的设置",
                SectionStyle.PRIMARY
            )
            # 创建其他控件...
        
        def get_settings(self) -> Dict[str, Any]:
            return {"key": "value"}
        
        def apply_settings(self) -> bool:
            return True
"""

import logging
import tkinter as tk
from abc import ABC, abstractmethod
from typing import Dict, Any, Callable, List, Optional

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.utils.font_utils import get_small_font
from gongwen_converter.utils.dpi_utils import scale
from .config import LAYOUT_CONFIG, SectionStyle

logger = logging.getLogger(__name__)


class BaseSettingsTab(tb.Frame, ABC):
    """
    设置选项卡抽象基类
    
    提供统一的初始化流程和UI构建辅助方法。
    所有设置选项卡必须继承此类并实现抽象方法。
    
    属性：
        config_manager: 配置管理器实例
        on_change: 设置变更回调函数
        layout_config: 布局配置实例
        small_font: 小字体名称
        small_size: 小字体大小
        scrollable_frame: 可滚动框架引用
        canvas: 画布引用
        scrollbar: 滚动条引用
    
    子类必须实现的抽象方法：
        _create_interface(): 创建选项卡界面
        get_settings(): 获取当前设置
        apply_settings(): 应用设置到配置文件
    """
    
    def __init__(
        self, 
        parent: tk.Widget, 
        config_manager: Any, 
        on_change: Callable[[str, Any], None]
    ):
        """
        初始化设置选项卡基类
        
        采用模板方法模式，定义标准的初始化流程：
        1. 初始化父类
        2. 保存核心引用
        3. 加载配置
        4. 创建滚动容器
        5. 调用子类界面创建方法
        6. 执行后置处理
        
        参数：
            parent: 父组件（通常是Notebook）
            config_manager: 配置管理器实例，用于读写配置
            on_change: 设置变更回调函数，签名为 (key: str, value: Any) -> None
        """
        super().__init__(parent)
        logger.debug(f"开始初始化 {self.__class__.__name__}")
        
        # 保存核心引用
        self.config_manager = config_manager
        self.on_change = on_change
        logger.debug(f"{self.__class__.__name__}: 已保存配置管理器和回调函数引用")
        
        # 加载配置
        self.layout_config = LAYOUT_CONFIG
        self.small_font, self.small_size = get_small_font()
        logger.debug(f"{self.__class__.__name__}: 已加载布局配置和字体设置")
        
        # 初始化UI组件引用（稍后创建）
        self.scrollable_frame: Optional[tb.Frame] = None
        self.canvas: Optional[tb.Canvas] = None
        self.scrollbar: Optional[tb.Scrollbar] = None
        
        # 执行初始化流程（模板方法）
        self._initialize()
        
        logger.info(f"{self.__class__.__name__} 初始化完成")
    
    def _initialize(self) -> None:
        """
        初始化流程（模板方法）
        
        定义了标准的初始化顺序，子类不应重写此方法。
        此方法按照固定顺序执行以下步骤：
        1. 创建滚动容器
        2. 调用子类的界面创建方法
        3. 执行后置处理（可选）
        """
        logger.debug(f"{self.__class__.__name__}: 开始执行初始化流程")
        
        # 步骤1: 创建滚动容器
        self._create_scrollable_container()
        logger.debug(f"{self.__class__.__name__}: 滚动容器创建完成")
        
        # 步骤2: 调用子类的界面创建方法
        self._create_interface()
        logger.debug(f"{self.__class__.__name__}: 界面创建完成")
        
        # 步骤3: 执行后置处理（子类可选择实现）
        self._post_initialize()
        logger.debug(f"{self.__class__.__name__}: 后置处理完成")
    
    def _create_scrollable_container(self) -> None:
        """
        创建可滚动容器
        
        创建包含画布、滚动条和滚动框架的标准布局。
        这个容器允许选项卡内容超过可视区域时进行滚动。
        
        组件结构：
            Frame (self)
            └── Frame (main_frame)
                ├── Canvas (canvas)
                │   └── Frame (scrollable_frame) - 实际的内容容器
                └── Scrollbar (scrollbar)
        """
        logger.debug(f"{self.__class__.__name__}: 创建滚动容器")
        
        # 创建主框架
        main_frame = tb.Frame(self)
        main_frame.pack(
            fill="both", 
            expand=True, 
            padx=self.layout_config.canvas_padding, 
            pady=self.layout_config.canvas_padding
        )
        
        # 创建画布和滚动条
        self.canvas = tb.Canvas(main_frame)
        self.scrollbar = tb.Scrollbar(
            main_frame, 
            orient="vertical", 
            command=self.canvas.yview
        )
        # 创建滚动框架，添加右侧内边距以与滚动条保持间距
        self.scrollable_frame = tb.Frame(
            self.canvas,
            padding=(0, 0, self.layout_config.scrollbar_spacing, 0)
        )
        
        # 配置滚动区域自动更新
        def update_scrollregion(event):
            """当内容大小改变时更新滚动区域"""
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        self.scrollable_frame.bind("<Configure>", update_scrollregion)
        
        # 将滚动框架添加到画布
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 布局画布和滚动条
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # 确保滚动框架宽度随画布调整
        def adjust_frame_width(event):
            """调整滚动框架宽度以匹配画布"""
            canvas_width = event.width
            self.canvas.itemconfig(
                self.canvas.find_all()[0],
                width=canvas_width
            )
        
        self.canvas.bind("<Configure>", adjust_frame_width)
        
        # 绑定鼠标滚轮事件以支持滚动
        def on_mousewheel(event):
            """处理鼠标滚轮事件（跨平台支持，带边界检查）"""
            # 获取当前滚动位置
            yview = self.canvas.yview()
            
            # 计算滚动方向
            delta = int(-1 * (event.delta / 120))
            
            # 边界检查
            if delta < 0 and yview[0] <= 0:
                # 向上滚动但已在顶部
                return
            elif delta > 0 and yview[1] >= 1:
                # 向下滚动但已在底部
                return
            
            # 执行滚动
            self.canvas.yview_scroll(delta, "units")
        
        def on_linux_scroll_up(event):
            """处理 Linux 系统向上滚动（带边界检查）"""
            # 检查是否已在顶部
            if self.canvas.yview()[0] <= 0:
                return
            self.canvas.yview_scroll(-1, "units")
        
        def on_linux_scroll_down(event):
            """处理 Linux 系统向下滚动（带边界检查）"""
            # 检查是否已在底部
            if self.canvas.yview()[1] >= 1:
                return
            self.canvas.yview_scroll(1, "units")
        
        # 将滚轮事件处理函数保存为实例属性，供递归绑定使用
        self._mousewheel_handler = on_mousewheel
        self._linux_scroll_up_handler = on_linux_scroll_up
        self._linux_scroll_down_handler = on_linux_scroll_down
        
        # 绑定到主要组件
        self.bind("<MouseWheel>", on_mousewheel)
        self.bind("<Button-4>", on_linux_scroll_up)
        self.bind("<Button-5>", on_linux_scroll_down)
        
        # 绑定到 canvas
        self.canvas.bind("<MouseWheel>", on_mousewheel)
        self.canvas.bind("<Button-4>", on_linux_scroll_up)
        self.canvas.bind("<Button-5>", on_linux_scroll_down)
        
        # 绑定到 scrollable_frame
        self.scrollable_frame.bind("<MouseWheel>", on_mousewheel)
        self.scrollable_frame.bind("<Button-4>", on_linux_scroll_up)
        self.scrollable_frame.bind("<Button-5>", on_linux_scroll_down)
        
        logger.debug(f"{self.__class__.__name__}: 滚动容器创建完成（已启用鼠标滚轮支持）")
    
    def _bind_mousewheel_recursive(self, widget: tk.Widget) -> None:
        """
        递归绑定鼠标滚轮事件到所有子组件
        
        此方法遍历给定组件的所有子组件，为每个子组件绑定鼠标滚轮事件。
        这确保了无论鼠标悬停在哪个控件上，滚轮都能正常工作。
        
        特殊处理：
        - Combobox：允许页面滚动，但阻止改变下拉框的值
        
        参数：
            widget: 要绑定的根组件
        """
        try:
            # 特殊处理 Combobox：允许页面滚动，但阻止改变值
            if isinstance(widget, tb.Combobox):
                # 绑定处理函数，并用"break"阻止默认行为（改变值）
                widget.bind("<MouseWheel>", 
                           lambda e: self._mousewheel_handler(e) or "break", 
                           add="+")
                widget.bind("<Button-4>", 
                           lambda e: self._linux_scroll_up_handler(e) or "break", 
                           add="+")
                widget.bind("<Button-5>", 
                           lambda e: self._linux_scroll_down_handler(e) or "break", 
                           add="+")
                logger.debug(f"已为 Combobox 绑定特殊滚轮处理（滚动页面但不改变值）")
                return  # 不递归子组件（下拉列表）
            
            # 其他组件正常绑定
            widget.bind("<MouseWheel>", self._mousewheel_handler, add="+")
            widget.bind("<Button-4>", self._linux_scroll_up_handler, add="+")
            widget.bind("<Button-5>", self._linux_scroll_down_handler, add="+")
            
            # 递归绑定所有子组件
            for child in widget.winfo_children():
                self._bind_mousewheel_recursive(child)
        except tk.TclError:
            # 某些组件可能不支持绑定，忽略错误
            pass
    
    # ========== 抽象方法（子类必须实现） ==========
    
    @abstractmethod
    def _create_interface(self) -> None:
        """
        创建选项卡界面（抽象方法）
        
        子类必须实现此方法来创建具体的设置界面。
        应在 self.scrollable_frame 中创建所有UI组件。
        
        实现示例：
            def _create_interface(self):
                frame = self.create_section_frame(
                    self.scrollable_frame,
                    "我的设置",
                    SectionStyle.PRIMARY
                )
                # 创建更多控件...
        """
        pass
    
    @abstractmethod
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置（抽象方法）
        
        子类必须实现此方法，返回当前所有设置项的值。
        
        返回：
            Dict[str, Any]: 设置字典，键为设置项名称，值为对应的值
            
        实现示例：
            def get_settings(self) -> Dict[str, Any]:
                return {
                    "option1": self.option1_var.get(),
                    "option2": self.option2_var.get()
                }
        """
        pass
    
    @abstractmethod
    def apply_settings(self) -> bool:
        """
        应用当前设置到配置文件（抽象方法）
        
        子类必须实现此方法，将当前设置保存到配置文件。
        
        返回：
            bool: 应用是否成功，True表示成功，False表示失败
            
        实现示例：
            def apply_settings(self) -> bool:
                try:
                    settings = self.get_settings()
                    for key, value in settings.items():
                        self.config_manager.update_config_value(..., key, value)
                    return True
                except Exception as e:
                    logger.error(f"应用设置失败: {e}")
                    return False
        """
        pass
    
    # ========== 可选的钩子方法 ==========
    
    def _post_initialize(self) -> None:
        """
        初始化后处理（可选钩子方法）
        
        在所有UI组件创建完成后，递归绑定鼠标滚轮事件到所有子组件。
        这确保了无论鼠标悬停在界面的任何位置，滚轮都能正常工作。
        
        子类可以重写此方法来执行额外的后置处理，但应该调用 super()._post_initialize()
        以确保滚轮绑定正常工作。
        """
        # 递归绑定所有子组件的滚轮事件
        self._bind_mousewheel_recursive(self.scrollable_frame)
        logger.debug(f"{self.__class__.__name__}: 已递归绑定所有子组件的鼠标滚轮事件")
    
    # ========== UI辅助方法 ==========
    
    def create_section_frame(
        self,
        parent: tk.Widget,
        title: str,
        style: SectionStyle = SectionStyle.PRIMARY,
        spacing_top: Optional[int] = None,
        spacing_bottom: Optional[int] = None
    ) -> tb.Labelframe:
        """
        创建设置区域框架
        
        创建一个带标题的Labelframe，用于组织相关的设置项。
        自动应用标准的间距和样式。
        
        参数：
            parent: 父组件，通常是 self.scrollable_frame
            title: 区域标题，显示在框架边框上
            style: 样式枚举，决定框架的颜色主题
            spacing_top: 顶部间距（像素），None使用默认值
            spacing_bottom: 底部间距（像素），None使用默认值
            
        返回：
            tb.Labelframe: 创建的区域框架
            
        使用示例：
            frame = self.create_section_frame(
                self.scrollable_frame,
                "基本设置",
                SectionStyle.PRIMARY
            )
            # 在frame中创建控件...
        """
        logger.debug(f"创建区域框架: {title} (样式: {style.value})")
        
        # 使用配置或自定义间距
        top = spacing_top if spacing_top is not None else self.layout_config.section_spacing_top
        bottom = spacing_bottom if spacing_bottom is not None else self.layout_config.section_spacing_bottom
        
        # 创建框架
        frame = tb.Labelframe(
            parent,
            text=title,
            padding=self.layout_config.section_padding,
            bootstyle=style.value
        )
        frame.pack(fill="x", pady=(scale(top), scale(bottom)))
        
        logger.debug(f"区域框架创建完成: {title}")
        return frame
    
    def create_checkbox_with_info(
        self,
        parent: tk.Widget,
        text: str,
        variable: tk.BooleanVar,
        tooltip: str,
        command: Optional[Callable] = None
    ) -> tb.Frame:
        """
        创建带信息图标的复选框
        
        创建一个复选框和它旁边的信息图标（ⓘ），
        鼠标悬停在图标上会显示工具提示。
        
        参数：
            parent: 父组件
            text: 复选框文本
            variable: 绑定的BooleanVar变量
            tooltip: 工具提示文本，悬停时显示
            command: 复选框状态改变时的回调函数
            
        返回：
            tb.Frame: 包含复选框和信息图标的容器框架
            
        使用示例：
            var = tk.BooleanVar(value=True)
            self.create_checkbox_with_info(
                parent_frame,
                "启用此功能",
                var,
                "这个功能用于...",
                lambda: self.on_setting_changed("feature", var.get())
            )
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        logger.debug(f"创建复选框+信息图标: {text}")
        
        # 创建容器
        frame = tb.Frame(parent)
        frame.pack(fill="x", pady=(0, self.layout_config.widget_spacing))
        
        # 创建复选框
        checkbox = tb.Checkbutton(
            frame,
            text=text,
            variable=variable,
            bootstyle="primary",
            command=command
        )
        checkbox.pack(side="left")
        
        # 创建信息图标
        info = create_info_icon(frame, tooltip, "info")
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        logger.debug(f"复选框创建完成: {text}")
        return frame
    
    def create_label_entry_pair(
        self,
        parent: tk.Widget,
        label_text: str,
        variable: tk.StringVar,
        tooltip: Optional[str] = None
    ) -> tb.Frame:
        """
        创建标签-输入框对
        
        创建一个标签和它下方的输入框，可选地在标签旁边显示信息图标。
        
        参数：
            parent: 父组件
            label_text: 标签文本
            variable: 绑定的StringVar变量
            tooltip: 可选的工具提示文本
            
        返回：
            tb.Frame: 包含标签和输入框的容器框架
            
        使用示例：
            var = tk.StringVar(value="默认值")
            self.create_label_entry_pair(
                parent_frame,
                "配置名称:",
                var,
                "输入配置的名称"
            )
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        logger.debug(f"创建标签-输入框对: {label_text}")
        
        # 创建容器
        container = tb.Frame(parent)
        container.pack(fill="x", pady=(0, self.layout_config.widget_spacing))
        
        # 创建标签行
        label_frame = tb.Frame(container)
        label_frame.pack(fill="x", pady=(0, self.layout_config.label_spacing))
        
        label = tb.Label(label_frame, text=label_text, bootstyle="secondary")
        label.pack(side="left")
        
        # 可选的信息图标
        if tooltip:
            info = create_info_icon(label_frame, tooltip, "info")
            info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 创建输入框
        entry = tb.Entry(container, textvariable=variable, bootstyle="secondary")
        entry.pack(fill="x")
        
        logger.debug(f"标签-输入框对创建完成: {label_text}")
        return container
    
    def create_combobox_with_info(
        self,
        parent: tk.Widget,
        label_text: str,
        variable: tk.StringVar,
        values: List[str],
        tooltip: str,
        command: Optional[Callable] = None
    ) -> tb.Frame:
        """
        创建带标签和信息图标的下拉框
        
        创建一个带标题和信息图标的下拉选择框。
        
        参数：
            parent: 父组件
            label_text: 标签文本
            variable: 绑定的StringVar变量
            values: 下拉选项列表
            tooltip: 工具提示文本
            command: 选择改变时的回调函数
            
        返回：
            tb.Frame: 包含标签和下拉框的容器框架
            
        使用示例：
            var = tk.StringVar(value="选项1")
            self.create_combobox_with_info(
                parent_frame,
                "选择模式:",
                var,
                ["选项1", "选项2", "选项3"],
                "选择运行模式",
                lambda e: self.on_mode_changed()
            )
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        logger.debug(f"创建下拉框+信息图标: {label_text}")
        
        # 创建容器
        container = tb.Frame(parent)
        container.pack(fill="x", pady=(0, self.layout_config.widget_spacing))
        
        # 创建标签行
        label_frame = tb.Frame(container)
        label_frame.pack(fill="x", pady=(0, self.layout_config.label_spacing))
        
        label = tb.Label(label_frame, text=label_text, bootstyle="secondary")
        label.pack(side="left")
        
        info = create_info_icon(label_frame, tooltip, "info")
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 创建下拉框
        combobox = tb.Combobox(
            container,
            textvariable=variable,
            values=values,
            state="readonly",
            bootstyle="secondary"
        )
        combobox.pack(fill="x")
        
        # 绑定事件
        if command:
            combobox.bind("<<ComboboxSelected>>", command)
        
        logger.debug(f"下拉框创建完成: {label_text}")
        return container
    
    def create_label_with_info(self, parent: tk.Widget, text: str, tooltip_text: str) -> tb.Frame:
        """
        创建带信息图标的标签
        
        创建一个标签和它旁边的信息图标。
        此方法保持向后兼容性。
        
        参数：
            parent: 父组件
            text: 标签文本
            tooltip_text: 工具提示文本
            
        返回：
            tb.Frame: 包含标签和信息图标的框架
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        logger.debug(f"创建标签+信息图标: {text}")
        
        label_frame = tb.Frame(parent)
        label_frame.pack(fill="x", pady=(0, self.layout_config.label_spacing))
        
        label = tb.Label(
            label_frame,
            text=text,
            bootstyle="secondary"
        )
        label.pack(side="left")
        
        info = create_info_icon(
            label_frame,
            tooltip_text,
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        logger.debug(f"标签+信息图标创建完成: {text}")
        return label_frame
