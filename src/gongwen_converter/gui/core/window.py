"""
公文转换器图形用户界面主窗口模块
"""
import os
import logging
import tkinter as tk
from typing import Optional, Any, Tuple
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import MessageDialog

from .layout import MainWindowLayoutBuilder
from .event_handler import MainWindowEventHandler
from .theme_manager import get_theme_manager
from gongwen_converter.utils.dpi_utils import initialize_dpi_manager, ScalableMixin
from gongwen_converter.gui.components.selector_manager import SelectorManager

logger = logging.getLogger()

class MainWindow(ScalableMixin):
    """
    公文转换器主窗口类
    """
    
    def __init__(self, root: tb.Window, config_manager: Any, initial_file_path: Optional[str] = None):
        logger.info("初始化主窗口")
        self.root = root
        self.config_manager = config_manager

        # 初始化DPI管理器
        initialize_dpi_manager(root)
        
        # 加载并应用缩放
        center_width = self.config_manager.get_center_panel_width()
        left_width = self.config_manager.get_batch_panel_width()
        right_width = self.config_manager.get_template_panel_width()
        height = self.config_manager.get_window_height()
        
        # 计算各种布局的宽度（使用固定边距10像素）
        self.center_width = self.scale(center_width)
        self.left_width = self.scale(left_width)
        self.right_width = self.scale(right_width)
        self.margin = self.scale(10)  # 固定边距
        self.original_height = self.scale(height)
        
        # 新布局方案：所有边距由各栏的 padx=(10,10) 管理
        # 单栏模式：10 + 中栏 + 10
        self.single_width = self.margin + self.center_width + self.margin
        # 双栏模式（左+中）：10 + 左栏 + 20 + 中栏 + 10
        self.left_center_width = self.margin + self.left_width + self.margin * 2 + self.center_width + self.margin
        # 双栏模式（中+右）：10 + 中栏 + 20 + 右栏 + 10
        self.center_right_width = self.margin + self.center_width + self.margin * 2 + self.right_width + self.margin
        # 三栏模式：10 + 左栏 + 20 + 中栏 + 20 + 右栏 + 10
        self.triple_width = self.margin + self.left_width + self.margin * 2 + self.center_width + self.margin * 2 + self.right_width + self.margin

        # 布局状态
        self.batch_panel_visible = False
        self.template_panel_visible = False
        
        # 获取窗口位置
        self.window_x, self.window_y = self.config_manager.get_window_position()
        
        # 位置跟踪变量（用于实时跟踪窗口位置）
        self.center_panel_screen_x = None  # 跟踪中栏在屏幕上的X坐标
        self.current_y = self.window_y      # 跟踪窗口的Y坐标
        
        self.default_theme = self.config_manager.get_default_theme()
        self.transparency_enabled = self.config_manager.is_transparency_enabled()
        self.transparency_value = self.config_manager.get_transparency_value()
        self.remember_gui_state = self.config_manager.should_remember_gui_state()
        self.auto_center = self.config_manager.should_auto_center()

        self._setup_window_properties()
        
        # 初始化主题管理器
        theme_manager = get_theme_manager()
        theme_manager.initialize(self.root, self.default_theme)
        
        self.current_file_path: Optional[str] = None
        self.selected_template: Optional[Tuple[str, str]] = None
        self.final_output_path: Optional[str] = None
        self.location_button: Optional[tb.Button] = None
        
        # 初始化应用级默认图标（一次设置，全局生效）
        from gongwen_converter.utils.icon_utils import IconManager
        IconManager.initialize(self.root)
        self.icon_path = IconManager.get_icon_path()
        
        self.main_container = tb.Frame(self.root, bootstyle="default")
        self.main_container.grid(row=0, column=0, sticky="nsew", padx=0, pady=self.scale(10))
        
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        self.event_handler = MainWindowEventHandler(self, self.config_manager)
        
        self._create_layout_with_builder()
        
        # 设置主窗口引用，供子组件访问
        self.main_container._main_window = self
        
        # 不再使用分隔线管理器，使用固定边距布局
        logger.debug("使用固定边距布局，不再使用分隔线")
        
        # 设置文件拖拽区域的回调函数
        if hasattr(self, 'file_drop_area') and self.file_drop_area:
            logger.debug("设置文件拖拽区域回调函数")
            self.file_drop_area.on_show_batch_panel = self.show_batch_panel
            self.file_drop_area.on_hide_batch_panel = self.hide_batch_panel
        
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        if initial_file_path:
            logger.info(f"处理命令行传入的文件: {initial_file_path}")
            self.root.after(100, lambda: self._load_initial_file(initial_file_path))
        
        # 延迟初始化位置跟踪（等待窗口完全显示）
        self.root.after(150, self._initialize_position_tracking)
        
        # 绑定窗口配置事件
        self.root.bind('<Configure>', self._on_window_configure)
        
        # 根据默认模式显示批量面板
        if self.file_drop_area and self.file_drop_area.get_mode() == "batch":
            self.root.after(100, self.show_batch_panel)
            logger.debug("默认模式为批量，将显示批量面板")

        logger.info("主窗口初始化完成")
    
    def _load_initial_file(self, file_path: str):
        if self.file_drop_area:
            self.file_drop_area.file_path = file_path
            self.file_drop_area._switch_to_file_state()
            self.event_handler.on_file_dropped(file_path)
            logger.info(f"初始文件加载完成: {file_path}")
        else:
            logger.error("文件拖拽区未初始化，无法加载初始文件")
    
    def _initialize_position_tracking(self):
        """
        初始化位置跟踪
        计算并记录中栏的初始屏幕坐标
        """
        try:
            # 获取窗口当前位置
            window_x = self.root.winfo_x()
            window_y = self.root.winfo_y()
            
            # 根据当前面板状态计算中栏X坐标
            if self.batch_panel_visible:
                self.center_panel_screen_x = window_x + self.left_width + self.margin * 2
            else:
                self.center_panel_screen_x = window_x + self.margin
            
            # 初始化Y坐标跟踪
            self.current_y = window_y
            
            logger.debug(f"位置跟踪已初始化: 中栏X={self.center_panel_screen_x}, 窗口Y={self.current_y}")
        except Exception as e:
            logger.error(f"初始化位置跟踪失败: {e}")
    
    def _on_window_configure(self, event):
        """
        窗口配置变化事件处理
        自动跟踪窗口位置变化，实时更新中栏X坐标和窗口Y坐标
        """
        # 只处理窗口本身的事件，忽略子组件
        if event.widget != self.root:
            return
        
        # 如果位置跟踪尚未初始化，跳过
        if self.center_panel_screen_x is None:
            return
        
        try:
            # 获取窗口当前位置
            window_x = self.root.winfo_x()
            window_y = self.root.winfo_y()
            
            # 根据窗口X和当前面板状态，计算中栏X坐标
            if self.batch_panel_visible:
                new_center_x = window_x + self.left_width + self.margin * 2
            else:
                new_center_x = window_x + self.margin
            
            # 更新中栏X坐标（设置阈值避免噪音）
            if abs(new_center_x - self.center_panel_screen_x) > 2:
                self.center_panel_screen_x = new_center_x
                logger.debug(f"中栏X坐标已更新: {new_center_x}")
            
            # 更新窗口Y坐标
            if abs(window_y - self.current_y) > 2:
                self.current_y = window_y
                logger.debug(f"窗口Y坐标已更新: {window_y}")
                
        except Exception as e:
            logger.error(f"窗口配置事件处理失败: {e}")
    
    def _setup_window_properties(self):
        self.root.withdraw()
        # 使用单栏宽度作为初始宽度
        self.root.geometry(f"{self.single_width}x{self.original_height}")
        min_h = self.config_manager.get_min_height()
        min_height = self.scale(min_h)
        self.root.minsize(self.single_width, min_height)
        self.root.resizable(False, True)
        self.root.title("公文转换器（完全离线版）")
        self._set_window_position()
        self._set_window_theme()
        self._set_window_transparency()
        self.root.deiconify()

    def _set_window_position(self):
        if self.auto_center:
            self._center_window()
        else:
            # 使用新的基于中栏坐标的位置计算
            x, y = self._calculate_window_position()
            self.root.geometry(f"+{x}+{y}")
    
    def _calculate_window_position(self) -> Tuple[int, int]:
        """
        基于中栏坐标计算窗口位置，考虑边界检测
        使用跟踪的位置值，而不是配置值
        """
        # 使用跟踪的中栏X坐标（如果未初始化则从配置读取）
        center_panel_x = self.center_panel_screen_x
        if center_panel_x is None:
            center_panel_x = self.config_manager.get_center_panel_screen_x()
            self.center_panel_screen_x = center_panel_x  # 初始化跟踪变量
        
        # 使用跟踪的Y坐标
        y = self.current_y if self.current_y is not None else self.window_y
        
        # 计算窗口宽度
        window_width = self._calculate_total_window_width()
        
        # 计算窗口X坐标（保持中栏位置固定）
        if self.batch_panel_visible:
            # 显示批量面板：窗口向左移动
            window_x = center_panel_x - self.left_width - self.margin * 3
        else:
            # 不显示批量面板：窗口X = 中栏X - 中栏左侧边距
            window_x = center_panel_x - self.margin
        
        # 边界检测
        window_x, y = self._adjust_position_for_screen_boundaries(window_x, y, window_width)
        
        # 根据调整后的窗口X，反向更新中栏X（边界检测可能调整了位置）
        if self.batch_panel_visible:
            self.center_panel_screen_x = window_x + self.left_width + self.margin * 2
        else:
            self.center_panel_screen_x = window_x + self.margin
        
        # 更新Y坐标
        self.current_y = y
        
        return window_x, y
    
    def _calculate_total_window_width(self) -> int:
        """计算当前布局下的窗口总宽度"""
        if self.batch_panel_visible and self.template_panel_visible:
            return self.triple_width
        elif self.batch_panel_visible:
            return self.left_center_width
        elif self.template_panel_visible:
            return self.center_right_width
        else:
            return self.single_width
    
    def _adjust_position_for_screen_boundaries(self, x: int, y: int, width: int) -> Tuple[int, int]:
        """根据屏幕边界调整窗口位置"""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 检查左边界
        if x < 0:
            x = 0
        
        # 检查右边界
        if x + width > screen_width:
            x = screen_width - width
        
        # 检查上边界
        if y < 0:
            y = 0
        
        # 检查下边界
        if y + self.original_height > screen_height:
            y = screen_height - self.original_height
        
        return x, y
    
    def _center_window(self):
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - self.root.winfo_width()) // 2
        y = (screen_height - self.root.winfo_height()) // 2
        self.root.geometry(f"+{x}+{y}")
    
    def _set_window_theme(self):
        try:
            self.root.style.theme_use(self.default_theme)
        except Exception as e:
            logger.error(f"设置主题失败: {str(e)}")
            self.root.style.theme_use("morph")
    
    def _set_window_transparency(self):
        if not self.transparency_enabled: return
        try:
            self.root.attributes('-alpha', self.transparency_value)
        except Exception as e:
            logger.warning(f"设置窗口透明度失败: {str(e)}")
    
    def _create_layout_with_builder(self):
        window_width = self.config_manager.get_center_panel_width()
        file_drop_height = self.config_manager.get_file_drop_height()

        builder = MainWindowLayoutBuilder(
            self.root, self.main_container,
            on_file_dropped=self.event_handler.on_file_dropped,
            on_template_selected=self.event_handler.on_template_selected,
            on_action=self.event_handler.on_action,
            on_cancel=self.event_handler.on_cancel,
            on_clear_clicked=self.event_handler.on_clear_clicked,
            on_settings_clicked=self.event_handler.on_settings_clicked,
            on_about_clicked=self.event_handler.on_about_clicked,
            config_manager=self.config_manager,
            window_width=window_width,
            file_drop_height=file_drop_height
        )
        
        components = builder.build_complete_layout()
        
        self.center_frame = components.get("center_frame")
        self.template_frame = components.get("template_frame")
        self.batch_frame = components.get("batch_frame")
        self.file_drop_area = components.get("file_drop_area")
        self.action_panel = components.get("action_panel")
        self.bottom_frame = components.get("bottom_frame")
        self.template_selector = components.get("template_selector")
        self.conversion_panel = components.get("conversion_panel")  # 格式转换面板
        self.tabbed_batch_file_list = components.get("batch_file_list")  # 现在是选项卡式批量文件列表
        self.status_bar = components.get("status_bar")
        
        if self.status_bar:
            self.status_bar.on_location_clicked = self._on_status_bar_location_clicked
        
        # 手动设置选项卡式批量文件列表的回调（在主窗口引用设置后）
        if self.tabbed_batch_file_list:
            self.tabbed_batch_file_list.on_list_cleared = self.on_batch_list_cleared
            self.tabbed_batch_file_list.on_tab_changed = self.on_tab_changed
            logger.debug("已手动设置选项卡式批量文件列表的回调")
        
        # 初始化选项卡式文件管理器
        self._initialize_tabbed_file_manager()
        
        # 绑定文件添加完成回调（必须在 selector_manager 初始化后）
        if self.tabbed_batch_file_list and hasattr(self, 'tabbed_file_manager'):
            self.tabbed_batch_file_list.on_files_added = lambda added, failed: self._on_batch_files_added(added, failed)
            logger.debug("已绑定文件添加完成回调")
        
        # 配置5列统一布局：左空白 + 批量 + 中栏 + 模板 + 右空白
        # 左右空白列weight=1实现自动居中和容错，内容列weight=0保持固定宽度
        self._setup_grid_layout()
    
    def _setup_grid_layout(self):
        """配置5列统一网格布局，一次性设置，后续无需改变"""
        self.main_container.grid_columnconfigure(0, weight=1)  # 左空白列
        self.main_container.grid_columnconfigure(1, weight=0)  # 批量面板列
        self.main_container.grid_columnconfigure(2, weight=0)  # 中栏列
        self.main_container.grid_columnconfigure(3, weight=0)  # 模板面板列
        self.main_container.grid_columnconfigure(4, weight=1)  # 右空白列
        
        # 初始化时只显示中栏（在列2）
        self.center_frame.grid(row=0, column=2, sticky="nsew", padx=(self.margin, self.margin))
        logger.debug("5列统一布局配置完成")

    def _on_close(self):
        self._save_gui_state()
        self.event_handler.on_close()

    def show_template_selector(self):
        """显示模板选择区域并调整布局"""
        logger.info("=== 开始显示模板面板 ===")
        logger.debug(f"当前模板面板可见状态: {self.template_panel_visible}")
        
        # 互斥显示：无论面板是否可见，都要确保显示模板选择器、隐藏转换面板
        if self.conversion_panel:
            self.conversion_panel.hide()
        if self.template_selector:
            self.template_selector.show()
            # 智能刷新模板列表（仅在模板目录内容变化时更新）
            self.template_selector.refresh_templates()
        
        if not self.template_panel_visible:
            logger.info("模板面板当前不可见，准备显示")
            
            # 在显示前，先基于当前窗口位置更新中栏坐标
            current_x = self.root.winfo_x()
            if self.batch_panel_visible:
                # 当前显示批量+中栏：中栏X = 窗口X + 批量宽度 + 边距
                center_panel_screen_x = current_x + self.left_width + self.margin * 2
            else:
                # 当前只显示中栏：中栏X = 窗口X + 左边距
                center_panel_screen_x = current_x + self.margin
            self.config_manager.update_config_value("gui_config", "window", "center_panel_screen_x", int(center_panel_screen_x))
            logger.debug(f"基于当前位置更新中栏屏幕坐标: {center_panel_screen_x}")
            
            self.template_panel_visible = True
            
            # 在5列布局中，模板面板固定在列3
            self.template_frame.grid(row=0, column=3, sticky="nsew", padx=(self.margin, self.margin))
            
            # 更新窗口几何形状
            self._update_window_geometry()
            
            if not hasattr(self, '_template_selector_shown') or not self._template_selector_shown:
                self._set_default_template_tab()
                self._template_selector_shown = True
            
            logger.info("模板面板显示完成")
        else:
            logger.debug("模板面板已经可见，仅切换内容")
        logger.info("=== 结束显示模板面板 ===")
    
    def hide_template_selector(self):
        """隐藏模板选择区域"""
        logger.info("=== 开始隐藏模板面板 ===")
        logger.debug(f"当前模板面板可见状态: {self.template_panel_visible}")
        
        if self.template_panel_visible:
            logger.info("模板面板当前可见，准备隐藏")
            self.template_panel_visible = False
            
            # 隐藏模板面板（5列布局无需调整grid配置）
            self.template_frame.grid_remove()
            if self.template_selector: self.template_selector.hide()
            if self.conversion_panel: self.conversion_panel.hide()
            
            # 更新窗口几何形状
            self._update_window_geometry()
            
            logger.info("模板面板隐藏完成")
        else:
            logger.debug("模板面板已经隐藏，无需操作")
        logger.info("=== 结束隐藏模板面板 ===")
    
    def show_conversion_panel(self, category: str, current_format: str, file_path: str = None):
        """
        显示格式转换面板
        
        参数:
            category: 文件类别 ('document', 'spreadsheet', 'image', 'layout')
            current_format: 当前文件格式 (如 'docx', 'png' 等)
            file_path: 文件路径（用于执行转换操作）
        """
        logger.info(f"显示格式转换面板: category={category}, format={current_format}, file={file_path}")
        
        # 获取当前选项卡的文件列表（批量模式）
        file_list = []
        if hasattr(self, 'tabbed_batch_file_list') and self.tabbed_batch_file_list:
            file_list = self.tabbed_batch_file_list.get_current_files()
            logger.debug(f"批量模式：传递文件列表（{len(file_list)}个文件）")
        
        # 获取UI模式
        ui_mode = self.file_drop_area.get_mode() if self.file_drop_area else 'single'
        logger.debug(f"当前UI模式: {ui_mode}")
        
        # 互斥显示：无论面板是否可见，都要确保显示转换面板、隐藏模板选择器
        if self.template_selector:
            self.template_selector.hide()
        
        if self.conversion_panel:
            self.conversion_panel.show()
            # 传递文件列表和UI模式
            self.conversion_panel.set_file_info(category, current_format, file_path, file_list, ui_mode)
        
        if not self.template_panel_visible:
            self.template_panel_visible = True
            self.template_frame.grid(row=0, column=3, sticky="nsew", padx=(self.margin, self.margin))
            
            # 更新窗口几何形状
            self._update_window_geometry()
        
        logger.info("格式转换面板显示完成")
    
    def hide_conversion_panel(self):
        """隐藏格式转换面板"""
        logger.info("隐藏格式转换面板")
        
        if self.conversion_panel:
            self.conversion_panel.hide()
        
        logger.info("格式转换面板隐藏完成")
    
    def on_format_selected(self, target_format: str):
        """
        格式选择回调（从转换面板触发）
        
        参数:
            target_format: 目标格式 (如 'PDF', 'DOCX' 等)
        """
        logger.info(f"格式转换按钮被点击: {target_format}")
        
        # TODO: 实现格式转换逻辑
        # 这里暂时只是占位，后续需要：
        # 1. 根据当前文件和目标格式，确定转换策略
        # 2. 调用相应的转换处理逻辑
        # 3. 显示转换结果
        
        logger.info(f"格式转换功能待实现: 转换到 {target_format}")
        self.add_status_message(f"格式转换功能开发中：{target_format}", "info", False)
    
    def _update_window_geometry(self):
        """
        统一的窗口几何形状更新方法
        根据当前面板可见状态，自动计算并更新窗口位置、大小和可调整性
        """
        height = self.root.winfo_height()
        
        # 根据面板状态设置窗口可调整性
        # 只有同时显示批量面板和模板面板（三栏模式）时才允许横向调整
        can_resize_horizontally = self.batch_panel_visible and self.template_panel_visible
        self.root.resizable(can_resize_horizontally, True)
        
        # 重新计算窗口位置和宽度
        x, y = self._calculate_window_position()
        width = self._calculate_total_window_width()
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
        # 更新最小尺寸
        min_h = self.config_manager.get_min_height()
        min_height = self.scale(min_h)
        self.root.minsize(width, min_height)
        
        logger.debug(f"窗口几何已更新: 位置({x}, {y}), 尺寸({width}x{height}), "
                    f"可调整大小: {can_resize_horizontally}x纵向")
        
    def _set_default_template_tab(self):
        """根据配置设置默认模板选项卡"""
        try:
            default_template_type = self.config_manager.get_default_md_template_type()
            logger.info(f"配置的默认模板类型: {default_template_type}")

            if not self.template_selector:
                logger.warning("模板选择器未初始化")
                return

            tab_index = 1 if default_template_type == "xlsx" else 0
            self.template_selector.notebook.select(tab_index)
            
            # 手动触发选项卡更改后的逻辑
            self.template_selector._on_notebook_tab_changed(None)
            
            logger.info(f"已根据配置设置默认选项卡为: {default_template_type.upper()}")
        except Exception as e:
            logger.error(f"设置默认模板选项卡失败: {e}")

    def show_action_panel(self):
        """显示操作面板"""
        logger.debug("显示操作面板")
        self.bottom_frame.grid(row=1, column=0, sticky="ew", pady=(0, self.scale(10)))
        if self.action_panel:
            self.action_panel.show()
    
    def hide_action_panel(self):
        """隐藏操作面板"""
        logger.debug("隐藏操作面板")
        if self.action_panel:
            self.action_panel.hide()
        self.bottom_frame.grid_remove()

    def show_batch_panel(self):
        """显示批量面板（批量文件列表）"""
        logger.info("=== 开始显示批量面板 ===")
        logger.debug(f"当前批量面板可见状态: {self.batch_panel_visible}")
        
        if not self.batch_panel_visible:
            logger.info("批量面板当前不可见，准备显示")
            
            self.batch_panel_visible = True
            
            # 在5列布局中，批量面板固定在列1
            self.batch_frame.grid(row=0, column=1, sticky="nsew", padx=(self.margin, self.margin))
            
            # 更新窗口几何形状
            self._update_window_geometry()
            
            logger.info("批量面板显示完成")
        else:
            logger.debug("批量面板已经可见，无需操作")
        logger.info("=== 结束显示批量面板 ===")
            
    def hide_batch_panel(self):
        """隐藏批量面板（批量文件列表）"""
        logger.info("=== 开始隐藏批量面板 ===")
        logger.debug(f"当前批量面板可见状态: {self.batch_panel_visible}")
        
        if self.batch_panel_visible:
            logger.info("批量面板当前可见，准备隐藏")
            
            self.batch_panel_visible = False
            
            # 隐藏批量面板（5列布局无需调整grid配置）
            self.batch_frame.grid_remove()
            
            # 更新窗口几何形状
            self._update_window_geometry()
            
            logger.info("批量面板隐藏完成")
        else:
            logger.debug("批量面板已经隐藏，无需操作")
        logger.info("=== 结束隐藏批量面板 ===")

    def add_status_message(self, message: str, message_type: str = "secondary", show_location: bool = False, file_path: Optional[str] = None):
        if self.status_bar:
            self.status_bar.add_message(message, message_type, show_location, file_path)

    def _save_gui_state(self):
        """
        保存GUI状态
        使用跟踪的位置值，确保准确性
        """
        if not self.remember_gui_state: return
        try:
            # 使用跟踪的中栏X坐标
            center_panel_screen_x = self.center_panel_screen_x
            if center_panel_screen_x is None:
                # 回退：如果跟踪失败，重新计算
                window_x = self.root.winfo_x()
                if self.batch_panel_visible:
                    center_panel_screen_x = window_x + self.left_width + self.margin * 2
                else:
                    center_panel_screen_x = window_x + self.margin
            
            # 使用跟踪的窗口Y坐标
            window_y = self.current_y if self.current_y is not None else self.root.winfo_y()
            
            # 窗口高度
            current_height = self.root.winfo_height()
            
            # 保存前，将高度还原为100%缩放下的值
            scaling_factor = self.get_scaling_factor()
            height_at_100_percent = int(current_height / scaling_factor)

            # 保存配置
            self.config_manager.update_config_value("gui_config", "window", "center_panel_screen_x", int(center_panel_screen_x))
            self.config_manager.update_config_value("gui_config", "window", "window_y", int(window_y))
            self.config_manager.update_config_value("gui_config", "window", "default_height", height_at_100_percent)
            
            logger.info(f"GUI状态已保存: 中栏X({center_panel_screen_x}), 窗口Y({window_y}), 原始高度{height_at_100_percent}")
        except Exception as e:
            logger.error(f"保存GUI状态失败: {str(e)}")

    def _on_status_bar_location_clicked(self, file_path: str):
        if file_path and os.path.exists(file_path):
            self.event_handler.logic._open_and_select_file(file_path)
    
    def on_batch_list_cleared(self):
        """批量列表已清空，重置UI到初始状态"""
        logger.info("批量列表已清空，重置UI")
        # 清空文件拖拽区域，效果等同于点击清空按钮
        if self.file_drop_area:
            self.file_drop_area._on_clear_clicked()
        logger.debug("UI已重置到初始状态")

    def _on_batch_files_added(self, added_files, failed_files):
        """
        文件添加完成回调
        在批量添加文件完成后，触发SelectorManager刷新UI
        
        参数:
            added_files: 成功添加的文件列表
            failed_files: 添加失败的文件及错误消息列表
        """
        logger.info(f"文件添加完成回调触发: {len(added_files)} 成功, {len(failed_files)} 失败")
        
        # 使用SelectorManager处理文件添加后的UI更新
        if hasattr(self, 'tabbed_file_manager') and self.tabbed_file_manager:
            # 获取当前激活的选项卡类别
            current_category = self.tabbed_batch_file_list.get_current_category()
            logger.debug(f"当前激活选项卡: {current_category}")
            
            # 触发选项卡切换逻辑，以刷新UI
            self.tabbed_file_manager.on_tab_changed(current_category)
        else:
            logger.warning("TabbedFileManager未初始化，无法刷新UI")
    
    def _initialize_tabbed_file_manager(self):
        """
        初始化选项卡式文件管理器
        
        重要：传递 self (main_window) 引用给 TabbedFileManager
        这样它才能调用 show_template_selector() 和 show_action_panel() 等方法
        来显示/隐藏UI框架
        """
        logger.info("⚙️ 初始化选项卡式文件管理器")
        
        # 创建选项卡式文件管理器，传递 main_window 引用
        logger.debug("  创建 TabbedFileManager 对象，传递 main_window=self")
        self.tabbed_file_manager = SelectorManager(
            tabbed_file_list=self.tabbed_batch_file_list,
            action_panel=self.action_panel,
            template_selector=self.template_selector,
            file_drop_area=self.file_drop_area,
            status_bar=self.status_bar,
            main_window=self,  # 传递主窗口引用
            config_manager=self.config_manager  # 传递配置管理器
        )
        logger.info("  ✓ TabbedFileManager 已创建，main_window 引用已传递")
        
        # 更新选项卡式批量文件列表的回调，使用文件管理器
        if self.tabbed_batch_file_list:
            logger.debug("  设置选项卡切换回调")
            self.tabbed_batch_file_list.on_tab_changed = self.tabbed_file_manager.on_tab_changed
            logger.debug("  ✓ 选项卡切换回调已设置")
        
        logger.info("✓ 选项卡式文件管理器初始化完成\n")
    
    def on_tab_changed(self, category: str):
        """选项卡切换事件处理（使用TabbedFileManager）"""
        logger.info(f"选项卡切换: {category}")
        
        # 使用TabbedFileManager处理选项卡切换
        if hasattr(self, 'tabbed_file_manager') and self.tabbed_file_manager:
            self.tabbed_file_manager.on_tab_changed(category)
        else:
            # 备用处理：如果没有文件管理器，使用原有逻辑
            logger.warning("TabbedFileManager未初始化，使用备用处理")
            self._handle_tab_changed_backup(category)
    
    def _handle_tab_changed_backup(self, category: str):
        """选项卡切换备用处理"""
        # 检查当前激活选项卡是否有文件
        if self.tabbed_batch_file_list:
            has_files = self.tabbed_batch_file_list.has_files(category)
            current_files = self.tabbed_batch_file_list.get_current_files()
            
            logger.debug(f"选项卡 '{category}' 有文件: {has_files}, 文件数量: {len(current_files)}")
            
            if not has_files:
                # 无文件状态：隐藏操作面板和模板面板
                logger.info(f"选项卡 '{category}' 无文件，隐藏操作面板和模板面板")
                self.hide_action_panel()
                self.hide_template_selector()
                
                # 如果文件拖拽区域在批量模式，恢复无文件状态
                if self.file_drop_area and self.file_drop_area.get_mode() == "batch":
                    self.file_drop_area._switch_to_empty_state()
            else:
                # 有文件状态：根据文件类型显示相应面板
                logger.info(f"选项卡 '{category}' 有文件，显示相应面板")
                if current_files:
                    first_file = current_files[0]
                    _, ext = os.path.splitext(first_file)
                    ext = ext.lower()
                    
                    logger.debug(f"根据第一个文件设置UI: {first_file}, 扩展名: {ext}")
                    
                    if ext in ['.md', '.txt']:
                        # 文本类：显示模板选择器和操作面板
                        self.show_template_selector()
                        self.show_action_panel()
                        if self.action_panel:
                            self.action_panel.setup_for_md_to_document(first_file)
                    elif ext in ['.docx', '.doc', '.wps', '.rtf']:
                        # 文档类：隐藏模板选择器，显示操作面板
                        self.hide_template_selector()
                        self.show_action_panel()
                        if self.action_panel:
                            self.action_panel.setup_for_document_file(first_file)
                    elif ext in ['.xlsx', '.xls', '.et', '.csv']:
                        # 表格类：隐藏模板选择器，显示操作面板
                        self.hide_template_selector()
                        self.show_action_panel()
                        if self.action_panel:
                            self.action_panel.setup_for_spreadsheet_file(first_file, current_files)
                    elif ext in ['.pdf']:
                        # 版式类：隐藏模板选择器，显示操作面板
                        self.hide_template_selector()
                        self.show_action_panel()
                        if self.action_panel:
                            self.action_panel.setup_for_layout_file(first_file)
                    elif ext in ['.tif', '.tiff', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.heic', '.heif']:
                        # 图片类：隐藏模板选择器，显示操作面板
                        self.hide_template_selector()
                        self.show_action_panel()
                        if self.action_panel:
                            self.action_panel.setup_for_image_file(first_file)

    def _show_error(self, message: str, title: str = "发生错误"):
        """使用弹窗显示错误消息"""
        logger.error(f"向用户显示错误弹窗: {title} - {message}")
        try:
            from gongwen_converter.gui.components.base_dialog import MessageBox
            MessageBox.showerror(title, message, parent=self.root)
        except Exception as e:
            logger.error(f"显示错误弹窗失败: {e}, 回退到状态栏")
            if self.status_bar:
                self.status_bar.add_message(f"错误: {message}", "danger")

    
    def get_current_theme(self) -> str:
        """获取当前主题名称"""
        theme_manager = get_theme_manager()
        return theme_manager.get_current_theme() or self.default_theme

    def refresh_theme(self, theme_name: str, preview_only: bool = False):
        """
        刷新并应用新主题
        
        参数:
            theme_name: 主题名称
            preview_only: 是否仅为预览模式（不保存为当前主题）
        """
        logger.info(f"正在应用新主题: {theme_name} (预览模式: {preview_only})")
        
        try:
            # 使用 ThemeManager 自动应用主题到所有窗口
            theme_manager = get_theme_manager()
            success = theme_manager.apply_theme(theme_name, preview_only=preview_only)
            
            if success:
                if not preview_only:
                    self.default_theme = theme_name
                logger.info(f"主题成功切换为: {theme_name}")
                
                # 刷新装饰区域（斜线条纹颜色跟随主题）
                if hasattr(self, 'conversion_panel') and self.conversion_panel:
                    self.conversion_panel.refresh_decoration()
            else:
                # 只在非预览模式下显示错误
                if not preview_only:
                    available_themes = theme_manager.get_available_themes()
                    logger.warning(f"主题 '{theme_name}' 应用失败，可用主题: {available_themes}")
                    self._show_error(
                        f"主题 '{theme_name}' 不可用。\n可用主题: {', '.join(available_themes)}", 
                        "主题错误"
                    )
                
        except Exception as e:
            logger.error(f"应用主题 '{theme_name}' 失败: {e}")
            if not preview_only:
                self._show_error(f"无法应用主题 '{theme_name}'。\n错误: {e}", "主题错误")

    def is_transparency_enabled(self) -> bool:
        """检查透明效果是否已启用"""
        return self.transparency_enabled

    def set_transparency(self, value: float):
        """设置窗口透明度"""
        if self.transparency_enabled:
            logger.debug(f"设置窗口透明度为: {value:.2f}")
            try:
                # 确保透明度值在安全范围内
                safe_value = max(0.1, min(1.0, value))
                self.root.attributes('-alpha', safe_value)
                self.transparency_value = safe_value  # 更新当前透明度记录
            except Exception as e:
                logger.warning(f"设置窗口透明度失败: {e}")

    def setup_md_action_panel(self, file_path: str):
        """设置MD文件操作面板"""
        if self.action_panel:
            self.action_panel.setup_for_md_to_document(file_path)
    
    def setup_docx_action_panel(self, file_path: str):
        """设置文档文件操作面板"""
        if self.action_panel:
            self.action_panel.setup_for_document_file(file_path)
    
    def setup_excel_action_panel(self, file_path: str):
        """设置电子表格文件操作面板"""
        if self.action_panel:
            self.action_panel.setup_for_md_to_spreadsheet(file_path)
    
    def setup_table_action_panel(self, file_path: str, file_list: list = None):
        """设置电子表格文件操作面板"""
        if self.action_panel:
            self.action_panel.setup_for_spreadsheet_file(file_path, file_list)
    
    def setup_image_action_panel(self, file_path: str):
        """设置图片文件操作面板"""
        if self.action_panel:
            self.action_panel.setup_for_image_file(file_path)
    
    def setup_layout_action_panel(self, file_path: str):
        """设置版式文件操作面板"""
        if self.action_panel:
            self.action_panel.setup_for_layout_file(file_path)
    def set_button_state(self, button_type: str, state: str):
        """设置按钮状态
        
        参数:
            button_type: 按钮类型
            state: 状态 ('normal', 'disabled')
        """
        if self.action_panel:
            self.action_panel.set_button_state(button_type, state)
    
    def clear_status_bar(self):
        """清空状态栏"""
        if self.status_bar:
            self.status_bar.clear_all()
    
    def handle_ipc_command(self, command: dict):
        """
        处理来自 IPC 的命令
        
        参数:
            command: 命令字典，包含 action 和其他参数
        """
        action = command.get('action')
        logger.info(f"收到 IPC 命令: {action}")
        
        if action == 'add_file':
            file_path = command.get('file_path')
            mode = command.get('mode', 'single')
            self.add_file_from_ipc(file_path, mode)
        
        elif action == 'clear':
            self.clear_files_from_ipc()
        
        elif action == 'activate':
            self.activate_window()
        
        else:
            logger.warning(f"未知的 IPC 命令: {action}")
    
    def add_file_from_ipc(self, file_path: str, mode: str):
        """
        从 IPC 接收文件并添加到界面
        
        参数:
            file_path: 文件路径
            mode: 模式 ('single' 或 'batch')
        """
        if not file_path:
            if mode == 'single':
                # 单文件模式：无文件时清空
                self.clear_files_from_ipc()
            # 批量模式：无文件时不做任何操作
            return
        
        # 在主线程中执行 GUI 操作
        self.root.after(0, lambda: self._add_file_safe(file_path, mode))
    
    def _add_file_safe(self, file_path: str, mode: str):
        """
        线程安全的文件添加方法
        
        参数:
            file_path: 文件路径
            mode: 模式 ('single' 或 'batch')
        """
        try:
            if mode == 'single':
                # 单文件模式：清空后添加新文件
                logger.info(f"单文件模式：加载文件 {file_path}")
                
                # 清空当前状态
                if self.file_drop_area:
                    self.file_drop_area._on_clear_clicked()
                
                # 加载新文件
                if self.file_drop_area:
                    self.file_drop_area.file_path = file_path
                    self.file_drop_area._switch_to_file_state()
                    self.event_handler.on_file_dropped(file_path)
                
            else:
                # 批量模式：添加到列表
                logger.info(f"批量模式：添加文件 {file_path}")
                
                # 确保批量面板可见
                if not self.batch_panel_visible:
                    self.show_batch_panel()
                
                # 添加文件到选项卡式批量列表
                if self.tabbed_batch_file_list:
                    self.tabbed_batch_file_list.add_file(file_path)
            
            # 激活窗口
            self.activate_window()
            
            logger.info(f"文件已通过 IPC 添加: {file_path}")
            
        except Exception as e:
            logger.error(f"添加文件失败: {e}", exc_info=True)
    
    def clear_files_from_ipc(self):
        """从 IPC 清空文件列表"""
        logger.info("收到清空文件列表命令")
        self.root.after(0, lambda: self._clear_files_safe())
    
    def _clear_files_safe(self):
        """线程安全的文件清空方法"""
        try:
            if self.file_drop_area:
                self.file_drop_area._on_clear_clicked()
            logger.info("文件列表已清空")
        except Exception as e:
            logger.error(f"清空文件列表失败: {e}", exc_info=True)
    
    def activate_window(self):
        """激活并置顶窗口"""
        logger.info("激活窗口")
        self.root.after(0, lambda: self._activate_window_safe())
    
    def _activate_window_safe(self):
        """线程安全的窗口激活方法"""
        try:
            # 如果窗口最小化，恢复它
            if self.root.state() == 'iconic':
                self.root.deiconify()
            
            # 置顶窗口
            self.root.lift()
            
            # 强制获取焦点
            self.root.focus_force()
            
            logger.debug("窗口已激活")
        except Exception as e:
            logger.error(f"激活窗口失败: {e}", exc_info=True)
