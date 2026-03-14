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

from __future__ import annotations

import contextlib
import logging
import tkinter as tk
from abc import ABC, abstractmethod
from collections.abc import Callable
from typing import Any, cast

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.utils.dpi_utils import scale
from docwen.utils.font_utils import get_small_font

from .config import LAYOUT_CONFIG, SectionStyle

logger = logging.getLogger(__name__)


# 软件颜色映射（使用 ttkbootstrap 语义色）
# 用于软件优先级卡片的左侧色条
SOFTWARE_COLORS = {
    # 文档处理软件
    "wps_writer": "danger",  # 红色 - WPS
    "msoffice_word": "info",  # 蓝绿色 - Microsoft Office
    "libreoffice": "success",  # 绿色 - LibreOffice
    # 表格处理软件
    "wps_spreadsheets": "danger",
    "msoffice_excel": "info",
}


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

    def __init__(self, parent: tk.Widget, config_manager: Any, on_change: Callable[[str, Any], None]):
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
        self.scrollable_frame: tb.Frame = cast(tb.Frame, None)
        self.canvas: tb.Canvas = cast(tb.Canvas, None)
        self.scrollbar: tb.Scrollbar = cast(tb.Scrollbar, None)

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

        # 创建主框架（padding需要DPI缩放）
        main_frame = tb.Frame(self)
        main_frame.pack(
            fill="both",
            expand=True,
            padx=scale(self.layout_config.canvas_padding),
            pady=scale(self.layout_config.canvas_padding),
        )

        # 创建画布和滚动条
        self.canvas = tb.Canvas(main_frame)

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

        self.scrollbar = tb.Scrollbar(main_frame, orient="vertical", command=bounded_yview)
        # 创建滚动框架，添加右侧内边距以与滚动条保持间距（需要DPI缩放）
        self.scrollable_frame = tb.Frame(self.canvas, padding=(0, 0, scale(self.layout_config.scrollbar_spacing), 0))

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
            self.canvas.itemconfig(self.canvas.find_all()[0], width=canvas_width)

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

    def _bind_mousewheel_recursive(self, widget: tk.Misc) -> None:
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
                widget.bind("<MouseWheel>", lambda e: self._mousewheel_handler(e) or "break", add="+")
                widget.bind("<Button-4>", lambda e: self._linux_scroll_up_handler(e) or "break", add="+")
                widget.bind("<Button-5>", lambda e: self._linux_scroll_down_handler(e) or "break", add="+")
                logger.debug("已为 Combobox 绑定特殊滚轮处理（滚动页面但不改变值）")
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

    @abstractmethod
    def get_settings(self) -> dict[str, Any]:
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

    def bind_label_wraplength(
        self, label: tb.Label, container: tk.Widget | None = None, min_wraplength: int = 0, padding: int = 0
    ) -> None:
        target = container or label.master
        if min_wraplength <= 0:
            min_wraplength = scale(240)
        if padding <= 0:
            padding = scale(40)
        state = {"retries": 0, "after_id": None}

        def update(_event=None):
            try:
                if not label.winfo_exists() or not target.winfo_exists():
                    return
                width = int(target.winfo_width())
            except Exception:
                return
            if width <= 1:
                try:
                    if state["retries"] >= 10:
                        return
                    if state["after_id"] is None:
                        state["retries"] += 1
                        state["after_id"] = label.after(50, scheduled_update)
                except Exception:
                    pass
                return

            if state["after_id"] is not None:
                with contextlib.suppress(Exception):
                    label.after_cancel(state["after_id"])
                state["after_id"] = None
            state["retries"] = 0
            wraplength = max(min_wraplength, width - padding)
            with contextlib.suppress(Exception):
                label.configure(wraplength=wraplength)

        def scheduled_update():
            state["after_id"] = None
            update()

        try:
            target.bind("<Configure>", update, add="+")
            label.after_idle(update)
        except Exception:
            pass

    def create_section_frame(
        self,
        parent: tk.Widget,
        title: str,
        style: SectionStyle = SectionStyle.PRIMARY,
        spacing_top: int | None = None,
        spacing_bottom: int | None = None,
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

        # 创建框架（padding需要DPI缩放）
        frame = tb.Labelframe(
            parent, text=title, padding=scale(self.layout_config.section_padding), bootstyle=style.value
        )
        frame.pack(fill="x", pady=(scale(top), scale(bottom)))

        logger.debug(f"区域框架创建完成: {title}")
        return frame

    def create_checkbox_with_info(
        self, parent: tk.Widget, text: str, variable: tk.BooleanVar, tooltip: str, command: Callable | None = None
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
        from docwen.utils.gui_utils import create_info_icon

        logger.debug(f"创建复选框+信息图标: {text}")

        # 创建容器
        frame = tb.Frame(parent)
        frame.pack(fill="x", pady=(0, self.layout_config.widget_spacing))

        # 创建复选框
        checkbox = tb.Checkbutton(frame, text=text, variable=variable, bootstyle="primary", command=command)
        checkbox.pack(side="left")

        # 创建信息图标
        info = create_info_icon(frame, tooltip, "info")
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))

        logger.debug(f"复选框创建完成: {text}")
        return frame

    def create_label_entry_pair(
        self, parent: tk.Widget, label_text: str, variable: tk.StringVar, tooltip: str | None = None
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
        from docwen.utils.gui_utils import create_info_icon

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
        values: list[str],
        tooltip: str,
        command: Callable | None = None,
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
        from docwen.utils.gui_utils import create_info_icon

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
        combobox = tb.Combobox(container, textvariable=variable, values=values, state="readonly", bootstyle="secondary")
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
        from docwen.utils.gui_utils import create_info_icon

        logger.debug(f"创建标签+信息图标: {text}")

        label_frame = tb.Frame(parent)
        label_frame.pack(fill="x", pady=(0, self.layout_config.label_spacing))

        label = tb.Label(label_frame, text=text, bootstyle="secondary")
        label.pack(side="left")

        info = create_info_icon(label_frame, tooltip_text, "info")
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))

        logger.debug(f"标签+信息图标创建完成: {text}")
        return label_frame

    def _create_extraction_mode_selectors(
        self, parent: tk.Widget, initial_ext_mode: str, initial_ocr_place: str, prefix: str
    ):
        """
        创建提取方式与OCR位置选择器
        """
        from docwen.gui.components.config_combobox import ConfigCombobox
        from docwen.i18n import t
        from docwen.utils.gui_utils import create_info_icon

        toplevel = parent.winfo_toplevel()
        var_name = "docwen_ocr_show_blockquote_title"
        try:
            toplevel.getvar(var_name)
            ocr_title_enabled_var = tk.BooleanVar(master=toplevel, name=var_name)
        except tk.TclError:
            ocr_title_enabled_var = tk.BooleanVar(
                master=toplevel,
                name=var_name,
                value=bool(self.config_manager.get_ocr_blockquote_title_enabled()),
            )
        self._ocr_blockquote_title_enabled_var = ocr_title_enabled_var

        ocr_combo = None
        prev_ocr_mode_before_base64 = None
        _refresh_title_toggle_state = None

        def refresh_ocr_combo_state(mode: str | None = None):
            nonlocal prev_ocr_mode_before_base64

            if ocr_combo is None:
                return

            current_mode = mode if mode is not None else ext_combo.get_config_value()
            if current_mode == "base64":
                current_ocr_mode = ocr_combo.get_config_value()
                if current_ocr_mode != "main_md":
                    prev_ocr_mode_before_base64 = current_ocr_mode
                    ocr_combo.set_config_value("main_md")
                    self.on_change("to_md_ocr_placement_mode", "main_md")
                ocr_combo.configure(state="disabled")
            else:
                ocr_combo.configure(state="readonly")
                if prev_ocr_mode_before_base64 is not None:
                    ocr_combo.set_config_value(prev_ocr_mode_before_base64)
                    self.on_change("to_md_ocr_placement_mode", prev_ocr_mode_before_base64)
                    prev_ocr_mode_before_base64 = None

            if _refresh_title_toggle_state is not None:
                _refresh_title_toggle_state()

        def ext_mode_changed(v: str):
            self.on_change("to_md_image_extraction_mode", v)
            refresh_ocr_combo_state(v)

        # 提取方式
        container1 = tb.Frame(parent)
        container1.pack(fill="x", pady=(scale(10), 0))

        label_frame1 = tb.Frame(container1)
        label_frame1.pack(fill="x", pady=(0, self.layout_config.label_spacing))

        label1 = tb.Label(
            label_frame1, text=t("settings.extraction.image_extraction_mode_label"), bootstyle="secondary"
        )
        label1.pack(side="left")

        info1 = create_info_icon(label_frame1, t("settings.extraction.image_extraction_mode_tooltip"), "info")
        info1.pack(side="left", padx=(self.layout_config.widget_spacing, 0))

        ext_combo = ConfigCombobox(
            container1,
            config_values=["file", "base64"],
            translate_keys={
                "file": "settings.extraction.image_extraction_mode_file",
                "base64": "settings.extraction.image_extraction_mode_base64",
            },
            initial_value=initial_ext_mode,
            on_change=ext_mode_changed,
        )
        ext_combo.pack(fill="x")
        setattr(self, f"{prefix}_image_ext_mode_combo", ext_combo)

        # OCR位置
        container2 = tb.Frame(parent)
        container2.pack(fill="x", pady=(scale(10), 0))

        label_frame2 = tb.Frame(container2)
        label_frame2.pack(fill="x", pady=(0, self.layout_config.label_spacing))

        label2 = tb.Label(label_frame2, text=t("settings.extraction.ocr_placement_mode_label"), bootstyle="secondary")
        label2.pack(side="left")

        info2 = create_info_icon(label_frame2, t("settings.extraction.ocr_placement_mode_tooltip"), "info")
        info2.pack(side="left", padx=(self.layout_config.widget_spacing, 0))

        ocr_combo = ConfigCombobox(
            container2,
            config_values=["image_md", "main_md"],
            translate_keys={
                "image_md": "settings.extraction.ocr_placement_mode_image_md",
                "main_md": "settings.extraction.ocr_placement_mode_main_md",
            },
            initial_value=initial_ocr_place,
            on_change=lambda v: self.on_change("to_md_ocr_placement_mode", v),
        )
        ocr_combo.pack(fill="x")
        setattr(self, f"{prefix}_ocr_placement_combo", ocr_combo)

        title_frame = tb.Frame(parent)
        title_frame.pack(fill="x", pady=(scale(10), 0))
        title_toggle = tb.Checkbutton(
            title_frame,
            text=t("settings.extraction.ocr_blockquote_title_enabled_label"),
            variable=ocr_title_enabled_var,
            bootstyle="round-toggle",
        )
        title_toggle.pack(side="left")
        info3 = create_info_icon(title_frame, t("settings.extraction.ocr_blockquote_title_enabled_tooltip"), "info")
        info3.pack(side="left", padx=(self.layout_config.widget_spacing, 0))

        def refresh_title_toggle_state(value: str | None = None):
            mode = value if value is not None else ocr_combo.get_config_value()
            title_toggle.configure(state=("normal" if mode == "main_md" else "disabled"))

        _refresh_title_toggle_state = refresh_title_toggle_state
        ocr_combo.bind("<<ComboboxSelected>>", lambda _e: refresh_title_toggle_state(), add="+")
        refresh_ocr_combo_state(initial_ext_mode)
        refresh_title_toggle_state()

    def _get_ocr_blockquote_title_enabled_setting(self) -> bool:
        var = getattr(self, "_ocr_blockquote_title_enabled_var", None)
        if var is not None:
            return bool(var.get())
        try:
            return bool(self.config_manager.get_ocr_blockquote_title_enabled())
        except Exception:
            return True

    def create_software_card(
        self,
        parent: tk.Widget,
        software_id: str,
        display_name: str,
        is_selected: bool,
        index: int,
        list_len: int,
        on_select: Callable,
        on_move: Callable[[str], None],
    ) -> tk.Frame:
        """
        创建带左侧色条的软件优先级卡片

        根据软件ID显示不同颜色的左侧色条，用于视觉区分不同软件。

        参数：
            parent: 父组件
            software_id: 软件标识符（用于颜色映射）
            display_name: 软件显示名称
            is_selected: 是否选中状态
            index: 当前索引位置
            list_len: 列表总长度
            on_select: 点击选中时的回调函数
            on_move: 移动方向回调函数，参数为 'left' 或 'right'

        返回：
            tk.Frame: 创建的卡片框架
        """
        logger.debug(f"创建软件卡片: {display_name} (选中: {is_selected})")

        # 获取主题颜色
        colors = None
        try:
            style = tb.Style.get_instance() or tb.Style()
            colors = style.colors
            parent_bg = colors.bg

            # 获取软件对应的语义色作为边框颜色
            bootstyle = SOFTWARE_COLORS.get(software_id, "secondary")
            fallback_color = getattr(colors, "secondary", None) or getattr(colors, "primary", None) or "#6C757D"
            software_color = getattr(colors, bootstyle, fallback_color)

            # 边框颜色使用软件对应的颜色，选中时加粗
            border_color = software_color
            border_width = scale(2) if is_selected else scale(1)
        except Exception:
            parent_bg = None
            try:
                style = tb.Style.get_instance() or tb.Style()
                colors = style.colors
                parent_bg = colors.bg
            except Exception:
                pass

            if not parent_bg:
                try:
                    parent_bg = parent.cget("background")
                except Exception:
                    parent_bg = None

            if not parent_bg:
                try:
                    parent_bg = self.cget("background")
                except Exception:
                    parent_bg = None
            # 后备颜色
            fallback_colors = {
                "wps_writer": "#DC3545",
                "wps_spreadsheets": "#DC3545",
                "msoffice_word": "#0D6EFD",
                "msoffice_excel": "#0D6EFD",
                "libreoffice": "#198754",
            }
            border_color = fallback_colors.get(software_id, "#6C757D")
            border_width = scale(2 if is_selected else 1)

        parent_bg = parent_bg or "#FFFFFF"

        # 外层卡片容器（边框颜色区分软件）
        card = tk.Frame(
            parent,
            bg=parent_bg,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=border_width,
        )
        card.pack(side="left", padx=scale(3))

        # 内容区
        content_frame = tk.Frame(card, bg=parent_bg)
        content_frame.pack(fill="both", expand=True)

        if is_selected:
            # 选中状态：显示左右移动按钮

            # 左移按钮
            if index > 0:
                left_btn = tk.Button(content_frame, text="◀", width=1, command=lambda: on_move("left"))
                try:
                    if colors is not None:
                        primary_color = getattr(colors, "primary", None)
                        if not primary_color:
                            raise ValueError("primary color not found")
                        left_btn.configure(
                            bg=primary_color,
                            fg="white",
                            activebackground=primary_color,
                            activeforeground="white",
                            relief=tk.FLAT,
                            borderwidth=0,
                            padx=scale(0),
                            pady=scale(6),
                            cursor="hand2",
                            font=(self.small_font, self.small_size),
                        )
                except Exception:
                    pass
                left_btn.pack(side="left")

            # 名称标签
            name_label = tb.Label(
                content_frame,
                text=display_name,
                font=(self.small_font, self.small_size),
                bootstyle="default",
                anchor=tk.CENTER,
                cursor="hand2",
                width=12,
            )
            name_label.pack(side="left", pady=(scale(6), scale(6)))
            name_label.bind("<Button-1>", lambda e: on_select())

            # 右移按钮
            if index < list_len - 1:
                right_btn = tk.Button(content_frame, text="▶", width=1, command=lambda: on_move("right"))
                try:
                    if colors is not None:
                        primary_color = getattr(colors, "primary", None)
                        if not primary_color:
                            raise ValueError("primary color not found")
                        right_btn.configure(
                            bg=primary_color,
                            fg="white",
                            activebackground=primary_color,
                            activeforeground="white",
                            relief=tk.FLAT,
                            borderwidth=0,
                            padx=scale(0),
                            pady=scale(6),
                            cursor="hand2",
                            font=(self.small_font, self.small_size),
                        )
                except Exception:
                    pass
                right_btn.pack(side="left")
        else:
            # 未选中状态：仅显示名称
            name_label = tb.Label(
                content_frame,
                text=display_name,
                font=(self.small_font, self.small_size),
                bootstyle="default",
                anchor=tk.CENTER,
                cursor="hand2",
                width=12,
            )
            name_label.pack(pady=(scale(8), scale(8)))

            # 绑定点击事件
            card.bind("<Button-1>", lambda e: on_select())
            content_frame.bind("<Button-1>", lambda e: on_select())
            name_label.bind("<Button-1>", lambda e: on_select())

        logger.debug(f"软件卡片创建完成: {display_name}")
        return card
