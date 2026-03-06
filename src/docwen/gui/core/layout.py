"""
主窗口布局构建模块
负责创建主窗口的所有UI组件和布局
"""

import logging
from collections.abc import Callable
from tkinter import ttk
from typing import Any, cast

# config_manager将通过依赖注入传入，因此这里不导入全局实例
# 导入ttkbootstrap用于界面美化和样式管理
import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.i18n import t  # 导入翻译函数
from docwen.utils.dpi_utils import ScalableMixin
from docwen.utils.font_utils import get_default_font, get_micro_font, get_small_font, get_title_font
from docwen.utils.icon_utils import load_image_icon

# 配置日志记录器
logger = logging.getLogger()


class MainWindowLayoutBuilder(ScalableMixin):
    """
    主窗口布局构建器类
    负责创建和配置主窗口的所有UI组件
    使用grid布局管理器，与业务逻辑完全分离
    """

    def __init__(
        self,
        root: tb.Window,
        main_container: tb.Frame,
        on_file_dropped: Callable[..., Any] | None = None,
        on_template_selected: Callable[..., Any] | None = None,
        on_action: Callable[..., Any] | None = None,
        on_cancel: Callable[..., Any] | None = None,
        on_clear_clicked: Callable[..., Any] | None = None,
        on_settings_clicked: Callable[..., Any] | None = None,
        on_about_clicked: Callable[..., Any] | None = None,
        config_manager: Any = None,
        window_width: int = 400,
        file_drop_height: int = 120,
    ):
        """
        初始化布局构建器

        参数:
            window_width: 窗口宽度（原始值，会自动缩放）
            file_drop_height: 文件拖拽区域高度（原始值，会自动缩放）
        """
        logger.info("初始化主窗口布局构建器")

        self.root = root
        self.main_container = main_container
        self.config_manager = config_manager

        # 应用缩放（使用继承自 ScalableMixin 的 scale 方法）
        self.window_width = self.scale(window_width)
        self.file_drop_height = self.scale(file_drop_height)

        self.on_file_dropped = on_file_dropped or (lambda *args, **kwargs: None)
        self.on_template_selected = on_template_selected or (lambda *args, **kwargs: None)
        self.on_action = on_action or (lambda *args, **kwargs: None)
        self.on_cancel = on_cancel or (lambda *args, **kwargs: None)
        self.on_clear_clicked = on_clear_clicked or (lambda *args, **kwargs: None)
        self.on_settings_clicked = on_settings_clicked or (lambda *args, **kwargs: None)
        self.on_about_clicked = on_about_clicked or (lambda *args, **kwargs: None)

        self.components = {}
        self._initialize_fonts()

        logger.info("布局构建器初始化完成")

    def _initialize_fonts(self):
        """初始化字体配置"""
        logger.debug("初始化字体配置")
        self.default_font, self.default_size = get_default_font()
        self.title_font, self.title_size = get_title_font()
        self.small_font, self.small_size = get_small_font()
        self.micro_font, self.micro_size = get_micro_font()
        logger.debug("字体配置初始化完成")

    def build_complete_layout(self) -> dict[str, Any]:
        """
        构建完整的主窗口布局
        """
        logger.info("开始构建完整的主窗口布局")

        self._configure_main_container_grid()

        center_frame = self._create_center_panel()
        template_frame = self._create_template_panel()
        batch_frame = self._create_batch_panel()

        batch_frame.grid(row=0, column=1, sticky="nsew", padx=(self.scale(10), self.scale(10)))
        logger.debug("批量面板框架已放置到网格: row=0, column=1, sticky=nsew, padx=(10, 10)")

        center_frame.grid(row=0, column=2, sticky="nsew", padx=(self.scale(10), self.scale(10)))
        logger.debug("中栏框架已放置到网格: row=0, column=2, sticky=nsew, padx=(10, 10)")

        template_frame.grid(row=0, column=3, sticky="nsew", padx=(self.scale(10), self.scale(10)))
        logger.debug("模板栏框架已放置到网格: row=0, column=3, sticky=nsew, padx=(10, 10)")

        template_frame.grid_remove()
        logger.debug("模板栏框架已隐藏 (grid_remove)")

        batch_frame.grid_remove()
        logger.debug("批量面板框架已隐藏 (grid_remove)")

        self.components["center_frame"] = center_frame
        self.components["template_frame"] = template_frame
        self.components["batch_frame"] = batch_frame

        logger.info("完整的主窗口布局构建完成")
        logger.debug(f"布局组件: {list(self.components.keys())}")
        return self.components

    def _configure_main_container_grid(self):
        """
        配置主容器的网格权重为五列，使用固定边距布局
        列0: 左空白列 - 容错/居中
        列1: 左栏 (批量面板)
        列2: 中栏 (主要内容区域)
        列3: 右栏 (模板面板)
        列4: 右空白列 - 容错/居中
        """
        logger.debug("配置主容器的网格权重为五列")

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)  # 左空白列
        self.main_container.grid_columnconfigure(1, weight=0)  # 批量面板列
        self.main_container.grid_columnconfigure(2, weight=4)  # 中栏列
        self.main_container.grid_columnconfigure(3, weight=0)  # 模板面板列
        self.main_container.grid_columnconfigure(4, weight=1)  # 右空白列

        logger.debug("主容器五列网格权重配置完成")

    def _create_center_panel(self) -> tb.Frame:
        """
        创建中栏框架（主要功能区域）
        """
        logger.debug("创建中栏框架")

        # 中栏宽度现在直接由self.window_width控制
        center_panel_width = self.window_width - self.scale(20)

        center_frame = tb.Frame(self.main_container, width=center_panel_width, bootstyle="default")
        center_frame.grid_propagate(False)

        # 配置行权重：让状态栏占满剩余垂直空间
        center_frame.grid_rowconfigure(0, weight=0)  # content_frame (文件输入+操作面板) - 固定高度
        center_frame.grid_rowconfigure(1, weight=0)  # 空白区域 - 固定高度（不再扩展）
        center_frame.grid_rowconfigure(2, weight=1)  # status_frame (状态栏) - 占满剩余空间
        center_frame.grid_rowconfigure(3, weight=0)  # separator (分隔线) - 固定高度
        center_frame.grid_rowconfigure(4, weight=0)  # button_frame (底部按钮) - 固定高度
        center_frame.grid_rowconfigure(5, weight=0)  # 保留行 - 固定高度
        center_frame.grid_columnconfigure(0, weight=1)

        content_frame = self._create_content_frame(center_frame)
        content_frame.grid(row=0, column=0, sticky="nsew")

        status_frame, status_bar = self._create_status_bar(center_frame)
        # 使用nsew让status_frame在垂直方向上也能扩展，占满剩余空间
        status_frame.grid(row=2, column=0, sticky="nsew", pady=(self.scale(10), self.scale(5)))

        separator = ttk.Separator(center_frame, orient="horizontal")
        separator.grid(row=3, column=0, sticky="ew", pady=(0, self.scale(5)))

        button_frame = self._create_button_footer(center_frame)
        button_frame.grid(row=4, column=0, sticky="ew", padx=self.scale(10), pady=(self.scale(5), 0))

        self.components["content_frame"] = content_frame
        self.components["status_frame"] = status_frame
        self.components["status_bar"] = status_bar
        self.components["button_frame"] = button_frame

        logger.debug(f"中栏框架创建完成 - 宽度{center_panel_width}像素")
        return center_frame

    def _create_content_frame(self, parent: tb.Frame) -> tb.Frame:
        """
        创建内容框架
        """
        logger.debug("创建内容框架")

        content_frame = tb.Frame(parent, bootstyle="default")

        content_frame.grid_rowconfigure(0, weight=0)
        content_frame.grid_rowconfigure(1, weight=0)
        content_frame.grid_rowconfigure(2, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)

        top_frame, file_drop_area = self._create_top_section(content_frame)
        top_frame.grid(row=0, column=0, sticky="ew", pady=(0, self.scale(20)))

        bottom_frame, action_panel = self._create_bottom_section(content_frame)

        self.components["top_frame"] = top_frame
        self.components["file_drop_area"] = file_drop_area
        self.components["bottom_frame"] = bottom_frame
        self.components["action_panel"] = action_panel

        logger.debug("内容框架创建完成")
        return content_frame

    def _create_top_section(self, parent: tb.Frame) -> tuple[tb.Labelframe, Any]:
        """
        创建上栏（文件拖拽区域）
        """
        logger.debug("创建上栏区域")

        top_frame = tb.Labelframe(
            parent, text=t("components.file_drop.section_title"), padding=self.scale(10), bootstyle="info"
        )

        top_frame.grid_rowconfigure(0, weight=1)
        top_frame.grid_columnconfigure(0, weight=1)

        try:
            from ..components.file_drop import FileDropArea

            # 从配置管理器读取默认模式
            default_mode = "single"  # 默认值
            if self.config_manager:
                try:
                    default_mode = self.config_manager.get_default_mode()
                    logger.info(f"从配置读取到默认模式: {default_mode}")
                except Exception as e:
                    logger.warning(f"读取默认模式配置失败，使用默认值 'single': {e}")

            file_drop_area = FileDropArea(
                top_frame,
                on_file_dropped=self.on_file_dropped,
                on_show_batch_panel=self._get_show_batch_panel_callback(),
                on_hide_batch_panel=self._get_hide_batch_panel_callback(),
                height=self.file_drop_height,
                default_mode=default_mode,
            )
            file_drop_area.set_clear_callback(self.on_clear_clicked)
            file_drop_area.grid(row=0, column=0, sticky="nsew")
            logger.debug(f"文件拖拽区域创建成功，默认模式: {default_mode}")
        except ImportError as e:
            logger.error(f"导入文件拖拽组件失败: {e!s}")
            file_drop_area = tb.Label(
                top_frame,
                text=t("components.file_drop.unavailable"),
                font=(self.default_font, self.default_size),
                bootstyle="danger",
            )
            file_drop_area.grid(row=0, column=0, sticky="nsew", pady=self.scale(20))

        logger.debug("上栏区域创建完成")
        return top_frame, file_drop_area

    def _get_show_batch_panel_callback(self):
        """获取显示批量面板的回调函数"""
        # 直接通过主容器获取主窗口引用
        main_window = getattr(self.main_container, "_main_window", None)
        if main_window:
            logger.debug("找到主窗口引用，返回show_batch_panel方法")
            return main_window.show_batch_panel
        logger.warning("无法找到主窗口引用")
        return None

    def _get_hide_batch_panel_callback(self):
        """获取隐藏批量面板的回调函数"""
        # 直接通过主容器获取主窗口引用
        main_window = getattr(self.main_container, "_main_window", None)
        if main_window:
            logger.debug("找到主窗口引用，返回hide_batch_panel方法")
            return main_window.hide_batch_panel
        logger.warning("无法找到主窗口引用")
        return None

    def _get_batch_list_cleared_callback(self):
        """获取批量列表清空的回调函数"""
        # 直接通过主容器获取主窗口引用
        main_window = getattr(self.main_container, "_main_window", None)
        if main_window:
            logger.debug("找到主窗口引用，返回on_batch_list_cleared方法")
            return getattr(main_window, "on_batch_list_cleared", None)
        else:
            logger.warning("无法找到主窗口引用")
            return None

    def _create_bottom_section(self, parent: tb.Frame) -> tuple[tb.Labelframe, Any]:
        """
        创建下栏（操作面板）
        """
        logger.debug("创建下栏区域")

        bottom_frame = tb.Labelframe(
            parent, text=t("components.action_panel.section_title"), padding=self.scale(10), bootstyle="primary"
        )

        bottom_frame.grid_rowconfigure(0, weight=1)
        bottom_frame.grid_columnconfigure(0, weight=1)

        try:
            from ..components.action_panel import ActionPanel

            action_panel = ActionPanel(
                bottom_frame, self.config_manager, on_action=self.on_action, on_cancel=self.on_cancel
            )
            action_panel.grid(row=0, column=0, sticky="nsew")
            logger.debug("操作面板创建成功")
        except ImportError as e:
            logger.error(f"导入操作面板失败: {e!s}")
            action_panel = tb.Label(bottom_frame, text=t("components.action_panel.unavailable"), bootstyle="danger")
            action_panel.grid(row=0, column=0, sticky="nsew", pady=self.scale(20))

        logger.debug("下栏区域创建完成")
        return bottom_frame, action_panel

    def _create_status_bar(self, parent: tb.Frame) -> tuple[tb.Frame, Any]:
        """
        创建状态栏（可滚动，占满剩余垂直空间）
        """
        logger.debug("创建可滚动状态栏")

        # 状态栏容器框架，配置为占满垂直空间
        status_frame = tb.Frame(parent, bootstyle="default")

        # 让状态栏内部可以垂直扩展
        status_frame.grid_rowconfigure(0, weight=1)
        status_frame.grid_columnconfigure(0, weight=1)

        try:
            from ..components.status_bar import StatusBar

            # 创建可滚动的StatusBar组件（不再需要line_count参数）
            status_bar = StatusBar(status_frame, on_location_clicked=None)
            # 使用nsew让StatusBar占满整个status_frame（包括垂直方向）
            status_bar.grid(row=0, column=0, sticky="nsew")
            logger.debug("可滚动StatusBar组件创建成功")
        except ImportError as e:
            logger.error(f"导入StatusBar组件失败: {e!s}")
            status_bar = tb.Label(status_frame, text=t("components.status_bar.unavailable"), bootstyle="danger")
            status_bar.grid(row=0, column=0, sticky="ew", pady=self.scale(20))

        self.components["status_bar"] = status_bar

        logger.debug("状态栏创建完成")

        return status_frame, status_bar

    def _create_button_footer(self, parent: tb.Frame) -> tb.Frame:
        """
        创建底部按钮框架
        """
        logger.debug("创建底部按钮框架")

        button_frame = tb.Frame(parent, bootstyle="default")

        button_frame.grid_rowconfigure(0, weight=1)
        button_frame.grid_columnconfigure(0, weight=0)  # 关于按钮 - 固定宽度
        button_frame.grid_columnconfigure(1, weight=1)  # 声明文本 - 可扩展
        button_frame.grid_columnconfigure(2, weight=0)  # 设置按钮 - 固定宽度

        icon_size = (self.scale(20), self.scale(20))

        about_icon = load_image_icon("about_icon.png", master=button_frame, size=icon_size)
        settings_icon = load_image_icon("settings_icon.png", master=button_frame, size=icon_size)

        about_button = tb.Button(
            button_frame, image=about_icon, command=self.on_about_clicked, bootstyle="secondary-link"
        )
        if about_icon:
            cast(Any, about_button).image = about_icon
        about_button.grid(row=0, column=0, sticky="w", padx=(0, self.scale(5)))

        from docwen import __version__

        version_label = tb.Label(
            button_frame,
            text=f"v{__version__}",
            font=(self.micro_font, self.micro_size),
            bootstyle="secondary",
            anchor="center",
        )
        version_label.grid(row=0, column=1, sticky="nsew")

        settings_button = tb.Button(
            button_frame, image=settings_icon, command=self.on_settings_clicked, bootstyle="secondary-link"
        )
        if settings_icon:
            cast(Any, settings_button).image = settings_icon
        settings_button.grid(row=0, column=2, sticky="e", padx=(self.scale(5), 0))

        self.components["about_button"] = about_button
        self.components["settings_button"] = settings_button
        self.components["version_label"] = version_label

        logger.debug("底部按钮框架创建完成")
        return button_frame

    def _create_template_panel(self) -> tb.Frame:
        """
        创建模板栏框架（模板选择区域和格式转换面板）
        """
        logger.debug("创建模板栏框架")

        # 获取模板面板宽度配置并应用缩放
        template_panel_width = self.config_manager.get_template_panel_width()
        scaled_width = self.scale(template_panel_width)
        logger.debug(f"模板面板宽度配置: {template_panel_width} -> 缩放后: {scaled_width}")

        # 创建固定宽度的模板面板框架
        template_frame = tb.Frame(self.main_container, width=scaled_width, bootstyle="default")
        template_frame.grid_propagate(False)  # 禁用网格传播，确保固定宽度生效

        template_frame.grid_rowconfigure(0, weight=1)
        template_frame.grid_columnconfigure(0, weight=1)

        right_panel_container, template_selector, conversion_panel = self._create_template_section(template_frame)
        right_panel_container.grid(row=0, column=0, sticky="nsew")

        self.components["right_panel_container"] = right_panel_container
        self.components["template_selector"] = template_selector
        self.components["conversion_panel"] = conversion_panel

        logger.debug(f"模板栏框架创建完成 - 固定宽度: {scaled_width}像素")
        return template_frame

    def _create_template_section(self, parent: tb.Frame) -> tuple[tb.Frame, Any, Any]:
        """
        创建模板选择区域和格式转换面板
        两者共享同一个右栏空间，根据文件类型切换显示

        返回:
            Tuple[tb.Frame, Any, Any]: (容器框架, 模板选择器, 转换面板)
        """
        logger.debug("创建模板选择区域和格式转换面板")

        # right_panel_container 就是传入的 parent（外层容器，无padding）
        right_panel_container = parent

        right_panel_container.grid_rowconfigure(0, weight=1)
        right_panel_container.grid_columnconfigure(0, weight=1)

        # 创建有padding的template_frame，用于模板选择器
        template_frame = tb.Frame(right_panel_container, padding=self.scale(10))
        template_frame.grid(row=0, column=0, sticky="nsew")

        template_frame.grid_rowconfigure(0, weight=1)
        template_frame.grid_columnconfigure(0, weight=1)

        # 模板选择器放在template_frame中（有padding美化）
        try:
            from ..components.template_selector_tabbed import TabbedTemplateSelector

            tabbed_template_selector = TabbedTemplateSelector(
                template_frame, on_template_selected=self.on_template_selected
            )
            tabbed_template_selector.grid(row=0, column=0, sticky="nsew")
            logger.debug("选项卡式模板选择器创建成功")
        except ImportError as e:
            logger.error(f"导入选项卡式模板选择器失败: {e!s}")
            tabbed_template_selector = tb.Label(
                template_frame, text=t("components.template_selector.unavailable"), bootstyle="danger"
            )
            tabbed_template_selector.grid(row=0, column=0, sticky="nsew", pady=self.scale(20))

        # 格式转换面板直接放在right_panel_container中（无padding，与template_frame同级）
        try:
            from ..components.conversion_panel import ConversionPanel

            conversion_panel = ConversionPanel(
                right_panel_container,
                config_manager=self.config_manager,
                on_action=self.on_action,
                height=self.file_drop_height,  # 传入文件拖拽区域高度，用于匹配格式转换section高度
            )
            conversion_panel.grid(row=0, column=0, sticky="nsew")
            logger.debug("格式转换面板创建成功")
        except ImportError as e:
            logger.error(f"导入格式转换面板失败: {e!s}")
            conversion_panel = tb.Label(
                right_panel_container, text=t("components.conversion_panel.unavailable"), bootstyle="danger"
            )
            conversion_panel.grid(row=0, column=0, sticky="nsew", pady=self.scale(20))

        # 默认隐藏转换面板（显示模板选择器）
        conversion_panel.grid_remove()

        logger.debug("模板选择区域和格式转换面板创建完成")

        # 返回外层容器、模板选择器、转换面板
        return right_panel_container, tabbed_template_selector, conversion_panel

    def _create_batch_panel(self) -> tb.Frame:
        """
        创建批量面板框架（批量文件列表区域）
        """
        logger.debug("创建批量面板框架")

        # 获取批量面板宽度配置并应用缩放
        batch_panel_width = self.config_manager.get_batch_panel_width()
        scaled_width = self.scale(batch_panel_width)
        logger.debug(f"批量面板宽度配置: {batch_panel_width} -> 缩放后: {scaled_width}")

        # 创建固定宽度的批量面板框架
        batch_frame = tb.Frame(self.main_container, width=scaled_width, bootstyle="default")
        batch_frame.grid_propagate(False)  # 禁用网格传播，确保固定宽度生效

        batch_frame.grid_rowconfigure(0, weight=1)
        batch_frame.grid_columnconfigure(0, weight=1)

        batch_section_frame, batch_file_list = self._create_batch_section(batch_frame)
        batch_section_frame.grid(row=0, column=0, sticky="nsew")

        self.components["batch_section_frame"] = batch_section_frame
        self.components["batch_file_list"] = batch_file_list

        logger.debug(f"批量面板框架创建完成 - 固定宽度: {scaled_width}像素")
        return batch_frame

    def _create_batch_section(self, parent: tb.Frame) -> tuple[tb.Frame, Any]:
        """
        创建批量文件列表区域（使用选项卡式布局）
        """
        logger.debug("创建选项卡式批量文件列表区域")

        batch_frame = tb.Frame(parent, padding=self.scale(10))

        batch_frame.grid_rowconfigure(0, weight=1)
        batch_frame.grid_columnconfigure(0, weight=1)

        try:
            from ..components.file_selector import TabbedFileSelector

            tabbed_batch_file_list = TabbedFileSelector(
                batch_frame,
                on_tab_changed=None,  # 选项卡切换回调将在主窗口中设置
                on_file_removed=None,  # 这些回调将在主窗口中设置
                on_file_opened=None,
                on_list_cleared=self._get_batch_list_cleared_callback(),
            )
            tabbed_batch_file_list.grid(row=0, column=0, sticky="nsew")
            logger.debug("选项卡式批量文件列表组件创建成功")
        except ImportError as e:
            logger.error(f"导入选项卡式批量文件列表组件失败: {e!s}")
            tabbed_batch_file_list = tb.Label(
                batch_frame, text=t("components.file_selector.unavailable"), bootstyle="danger"
            )
            tabbed_batch_file_list.grid(row=0, column=0, sticky="nsew", pady=self.scale(20))

        logger.debug("选项卡式批量文件列表区域创建完成")
        return batch_frame, tabbed_batch_file_list
