"""
编辑器对话框基类模块

提供编辑器对话框的通用功能，包括：
- 窗口配置
- 状态管理（未保存修改跟踪）
- 列表管理（带复选框的项目列表）
- 项目操作（新增、复制、删除、排序）
- 对话框辅助方法
- 唯一ID生成
"""

import logging
import random
import string
import tkinter as tk
from abc import ABC, abstractmethod
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

import ttkbootstrap as tb

from gongwen_converter.utils.dpi_utils import ScalableMixin
from gongwen_converter.utils.font_utils import get_small_font

logger = logging.getLogger(__name__)


class BaseEditorDialog(tb.Toplevel, ScalableMixin, ABC):
    """
    编辑器对话框基类
    
    提供通用的窗口配置、状态管理、列表管理和辅助方法。
    子类需要实现抽象方法来定制具体功能。
    """
    
    # 子类需要覆盖的类属性
    WINDOW_TITLE: str = "编辑器"
    CONFIG_FILE_NAME: str = "config.toml"
    
    # 窗口尺寸配置（子类可覆盖）
    WINDOW_WIDTH: int = 900
    WINDOW_HEIGHT: int = 600
    MIN_WIDTH: int = 750
    MIN_HEIGHT: int = 500
    
    # 布局配置
    LEFT_PANEL_WIDTH: int = 260
    PADDING: int = 10
    SECTION_PADDING: int = 5
    
    # 左侧面板标题（子类可覆盖）
    LEFT_PANEL_TITLE: str = "列表"
    
    def __init__(
        self,
        parent: tk.Widget,
        config_manager: Any,
        on_save: Optional[Callable[[], None]] = None
    ):
        """
        初始化编辑器对话框
        
        参数:
            parent: 父窗口
            config_manager: 配置管理器
            on_save: 保存成功后的回调函数
        """
        super().__init__(parent)
        
        self.parent = parent
        self.config_manager = config_manager
        self.on_save_callback = on_save
        
        # 获取字体配置
        self.small_font, self.small_size = get_small_font()
        
        # 配置文件路径
        self.config_file_path = self._get_config_file_path()
        
        # 状态管理
        self.has_unsaved_changes: bool = False
        self._is_loading: bool = False
        
        # 项目管理（子类使用）
        self.items: Dict[str, Any] = {}  # 项目字典 {id: item_object}
        self.item_order: List[str] = []  # 项目顺序列表
        self.current_item_id: Optional[str] = None  # 当前选中的项目ID
        self.checkbox_vars: Dict[str, tk.BooleanVar] = {}  # 复选框变量
        
        # UI组件引用
        self.status_label: Optional[tb.Label] = None
        self.item_list_frame: Optional[tb.Frame] = None
        self.item_canvas: Optional[tk.Canvas] = None
        self.name_var: Optional[tk.StringVar] = None
        self.name_entry: Optional[tb.Entry] = None
        self.description_text: Optional[tb.Text] = None
    
    # ==================== 配置方法 ====================
    
    def _get_config_file_path(self) -> Path:
        """获取配置文件路径"""
        try:
            config_dir = Path(self.config_manager._config_dir)
            return config_dir / self.CONFIG_FILE_NAME
        except Exception as e:
            logger.error(f"获取配置文件路径失败: {e}")
            return Path(f"configs/{self.CONFIG_FILE_NAME}")
    
    def _configure_window(self):
        """配置窗口属性"""
        width = self.scale(self.WINDOW_WIDTH)
        height = self.scale(self.WINDOW_HEIGHT)
        min_width = self.scale(self.MIN_WIDTH)
        min_height = self.scale(self.MIN_HEIGHT)
        
        self.title(self.WINDOW_TITLE)
        self.geometry(f"{width}x{height}")
        self.minsize(min_width, min_height)
        
        # 设置窗口图标
        self._setup_icon()
        
        # 使对话框模态
        self.transient(self.parent)
        self.grab_set()
        
        # 绑定关闭事件
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # 绑定快捷键
        self.bind("<Control-s>", lambda e: self._on_save())
        self.bind("<Escape>", lambda e: self._on_close())
    
    def _setup_icon(self):
        """设置窗口图标"""
        try:
            from gongwen_converter.utils.icon_utils import IconManager
            IconManager.set_window_icon(self)
        except Exception as e:
            logger.debug(f"设置窗口图标失败: {e}")
    
    # ==================== UI 创建方法 ====================
    
    def _create_left_panel(self, parent: tb.Frame) -> tb.Labelframe:
        """
        创建左侧面板（项目列表）
        
        返回：左侧面板frame，子类可以在其中添加额外内容
        """
        left_frame = tb.Labelframe(
            parent,
            text=self.LEFT_PANEL_TITLE,
            padding=self.SECTION_PADDING
        )
        left_frame.pack(
            side="left",
            fill="y",
            padx=(0, self.PADDING)
        )
        left_frame.pack_propagate(False)
        left_frame.configure(width=self.scale(self.LEFT_PANEL_WIDTH))
        
        return left_frame
    
    def _create_item_list(self, parent: tb.Frame):
        """创建项目列表（带复选框和滚动）"""
        list_frame = tb.Frame(parent)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # 创建Canvas用于滚动
        canvas = tk.Canvas(list_frame, highlightthickness=0)
        
        # 自定义滚动命令（带边界检查，防止滚动超出内容区域）
        def bounded_yview(*args):
            """带边界检查的滚动命令"""
            if args[0] == 'scroll':
                # 点击滚动条按钮时的滚动操作：检查是否已到边界
                top, bottom = canvas.yview()
                direction = int(args[1])
                if top <= 0 and direction < 0:  # 已在顶部，不能再向上滚动
                    return
                if bottom >= 1 and direction > 0:  # 已在底部，不能再向下滚动
                    return
            canvas.yview(*args)
        
        scrollbar = tb.Scrollbar(list_frame, orient="vertical", command=bounded_yview)
        
        self.item_list_frame = tb.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        canvas_window = canvas.create_window((0, 0), window=self.item_list_frame, anchor="nw")
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        self.item_list_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # 绑定鼠标滚轮（带边界检查）
        def on_mousewheel(event):
            top, bottom = canvas.yview()
            if top <= 0 and event.delta > 0:
                return
            if bottom >= 1 and event.delta < 0:
                return
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        self.item_canvas = canvas
    
    def _refresh_item_list(self):
        """刷新项目列表（带复选框）"""
        if not self.item_list_frame:
            return
        
        # 清空现有内容
        for widget in self.item_list_frame.winfo_children():
            widget.destroy()
        
        self.checkbox_vars.clear()
        
        # 按 order 顺序创建复选框
        for item_id in self.item_order:
            item = self.items.get(item_id)
            if not item:
                continue
            
            row_frame = tb.Frame(self.item_list_frame)
            row_frame.pack(fill="x", pady=2)
            
            # 复选框（启用/禁用）
            var = tk.BooleanVar(value=self._get_item_enabled(item))
            self.checkbox_vars[item_id] = var
            
            cb = tb.Checkbutton(
                row_frame,
                variable=var,
                command=lambda iid=item_id: self._on_item_enabled_changed(iid),
                bootstyle="round-toggle"
            )
            cb.pack(side="left")
            
            # 项目名称标签（可点击选中）
            prefix = "🔒" if self._is_item_system(item) else "📝"
            label = tb.Label(
                row_frame,
                text=f"{prefix} {self._get_item_name(item)}",
                font=(self.small_font, self.small_size),
                cursor="hand2"
            )
            label.pack(side="left", fill="x", expand=True)
            label.bind("<Button-1>", lambda e, iid=item_id: self._on_item_clicked(iid))
            
            # 存储标签引用用于高亮
            row_frame.item_label = label
            row_frame.item_id = item_id
    
    def _update_list_highlight(self):
        """更新列表高亮"""
        if not self.item_list_frame:
            return
        
        for widget in self.item_list_frame.winfo_children():
            if hasattr(widget, 'item_id') and hasattr(widget, 'item_label'):
                if widget.item_id == self.current_item_id:
                    widget.configure(bootstyle="info")
                    widget.item_label.configure(bootstyle="inverse-info")
                else:
                    widget.configure(bootstyle="default")
                    widget.item_label.configure(bootstyle="default")
    
    def _create_action_buttons(self, parent: tb.Frame):
        """创建操作按钮（新增/复制/删除/排序）"""
        # 第一行：新增/复制/删除按钮（居中分散排列）
        action_frame = tb.Frame(parent)
        action_frame.pack(fill="x", pady=(0, 5))
        
        # 使用三列grid布局实现居中分散
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)
        action_frame.columnconfigure(2, weight=1)
        
        tb.Button(
            action_frame,
            text="➕ 新增",
            command=self._on_add_item,
            bootstyle="success",
            width=7
        ).grid(row=0, column=0, padx=2)
        
        tb.Button(
            action_frame,
            text="📋 复制",
            command=self._on_copy_item,
            bootstyle="info",
            width=7
        ).grid(row=0, column=1, padx=2)
        
        tb.Button(
            action_frame,
            text="➖ 删除",
            command=self._on_delete_item,
            bootstyle="danger",
            width=7
        ).grid(row=0, column=2, padx=2)
        
        # 第二行：排序按钮（居中分散排列）
        order_frame = tb.Frame(parent)
        order_frame.pack(fill="x")
        
        # 使用两列grid布局实现居中分散
        order_frame.columnconfigure(0, weight=1)
        order_frame.columnconfigure(1, weight=1)
        
        tb.Button(
            order_frame,
            text="⬆ 上移",
            command=self._on_move_up,
            bootstyle="primary-outline",
            width=12
        ).grid(row=0, column=0, padx=2)
        
        tb.Button(
            order_frame,
            text="⬇ 下移",
            command=self._on_move_down,
            bootstyle="primary-outline",
            width=12
        ).grid(row=0, column=1, padx=2)
    
    def _create_basic_info_section(self, parent: tb.Frame, title: str = "基本信息"):
        """创建基本信息编辑区（名称和说明）"""
        frame = tb.Labelframe(parent, text=title, padding=5)
        frame.pack(fill="x", pady=(0, 10))
        
        # 名称
        name_frame = tb.Frame(frame)
        name_frame.pack(fill="x", pady=(0, 5))
        
        tb.Label(
            name_frame,
            text="名称:",
            width=8,
            anchor="e"
        ).pack(side="left")
        
        self.name_var = tk.StringVar()
        self.name_entry = tb.Entry(
            name_frame,
            textvariable=self.name_var,
            width=40
        )
        self.name_entry.pack(side="left", padx=(5, 0), fill="x", expand=True)
        self.name_var.trace_add("write", lambda *_: self._on_content_changed())
        
        # 说明
        desc_frame = tb.Frame(frame)
        desc_frame.pack(fill="x")
        
        tb.Label(
            desc_frame,
            text="说明:",
            width=8,
            anchor="e"
        ).pack(side="left", anchor="n")
        
        self.description_text = tb.Text(
            desc_frame,
            width=40,
            height=2,
            font=(self.small_font, self.small_size)
        )
        self.description_text.pack(side="left", padx=(5, 0), fill="x", expand=True)
        self.description_text.bind("<KeyRelease>", lambda e: self._on_content_changed())
        
        return frame
    
    def _create_bottom_buttons(self, parent: tb.Frame):
        """创建底部按钮"""
        button_frame = tb.Frame(parent)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # 左侧提示
        self.status_label = tb.Label(
            button_frame,
            text="",
            bootstyle="secondary",
            font=(self.small_font, self.small_size)
        )
        self.status_label.pack(side="left")
        
        # 右侧按钮
        tb.Button(
            button_frame,
            text="取消",
            command=self._on_close,
            bootstyle="secondary",
            width=10
        ).pack(side="right", padx=(5, 0))
        
        tb.Button(
            button_frame,
            text="保存",
            command=self._on_save,
            bootstyle="success",
            width=10
        ).pack(side="right")
    
    # ==================== 状态管理方法 ====================
    
    def _mark_changed(self):
        """标记有未保存的修改"""
        if not self.has_unsaved_changes:
            self.has_unsaved_changes = True
            self.title(f"{self.WINDOW_TITLE} *")
            if self.status_label:
                self.status_label.configure(text="有未保存的修改")
    
    def _clear_changed(self):
        """清除修改标记"""
        self.has_unsaved_changes = False
        self.title(self.WINDOW_TITLE)
        if self.status_label:
            self.status_label.configure(text="")
    
    # ==================== ID 生成方法 ====================
    
    def _generate_unique_id(self, prefix: str = "custom", existing_ids: set = None) -> str:
        """
        生成唯一ID
        
        格式：{prefix}_YYYYMMDD_HHMMSS_xxx
        示例：custom_20241217_163407_a3f
        
        参数:
            prefix: ID前缀
            existing_ids: 已存在的ID集合，用于确保唯一性
        
        返回:
            唯一的ID字符串
        """
        if existing_ids is None:
            existing_ids = set(self.items.keys())
        
        # 日期时间格式
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # 3位随机字符防止同秒冲突
        random_suffix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=3))
        base_id = f"{prefix}_{timestamp}_{random_suffix}"
        
        # 确保唯一性
        unique_id = base_id
        counter = 1
        while unique_id in existing_ids:
            unique_id = f"{base_id}_{counter}"
            counter += 1
        
        return unique_id
    
    # ==================== 项目操作方法 ====================
    
    def _on_item_clicked(self, item_id: str):
        """处理项目点击事件"""
        self._select_item(item_id)
    
    def _on_item_enabled_changed(self, item_id: str):
        """处理项目启用状态变更"""
        if item_id not in self.items:
            return
        
        var = self.checkbox_vars.get(item_id)
        if var:
            self._set_item_enabled(self.items[item_id], var.get())
            self._mark_changed()
            logger.debug(f"项目 {item_id} 启用状态: {var.get()}")
    
    def _select_item(self, item_id: str):
        """选择并显示指定项目"""
        if item_id not in self.items:
            return
        
        # 设置加载标志，防止触发变更回调
        self._is_loading = True
        
        try:
            self.current_item_id = item_id
            item = self.items[item_id]
            
            logger.debug(f"选择项目: {item_id} ({self._get_item_name(item)})")
            
            # 更新基本信息
            if self.name_var:
                self.name_var.set(self._get_item_name(item))
            if self.description_text:
                self.description_text.delete("1.0", tk.END)
                self.description_text.insert("1.0", self._get_item_description(item))
            
            # 更新列表高亮
            self._update_list_highlight()
            
            # 调用子类的额外选择逻辑
            self._on_item_selected(item)
            
        finally:
            # 清除加载标志
            self._is_loading = False
    
    def _on_content_changed(self):
        """处理内容变更"""
        # 加载中或无当前项目时不处理
        if self._is_loading or not self.current_item_id:
            return
        
        self._mark_changed()
        
        # 更新当前项目数据
        item = self.items.get(self.current_item_id)
        if item:
            if self.name_var:
                self._set_item_name(item, self.name_var.get())
            if self.description_text:
                self._set_item_description(item, self.description_text.get("1.0", "end-1c"))
            
            # 刷新列表显示
            self._refresh_item_list()
            self._update_list_highlight()
    
    def _on_move_up(self):
        """上移当前项目"""
        if not self.current_item_id:
            return
        
        index = self.item_order.index(self.current_item_id) if self.current_item_id in self.item_order else -1
        if index <= 0:
            return
        
        # 交换位置
        self.item_order[index], self.item_order[index - 1] = \
            self.item_order[index - 1], self.item_order[index]
        
        self._mark_changed()
        self._refresh_item_list()
        self._update_list_highlight()
        
        logger.debug(f"项目 {self.current_item_id} 上移")
    
    def _on_move_down(self):
        """下移当前项目"""
        if not self.current_item_id:
            return
        
        index = self.item_order.index(self.current_item_id) if self.current_item_id in self.item_order else -1
        if index < 0 or index >= len(self.item_order) - 1:
            return
        
        # 交换位置
        self.item_order[index], self.item_order[index + 1] = \
            self.item_order[index + 1], self.item_order[index]
        
        self._mark_changed()
        self._refresh_item_list()
        self._update_list_highlight()
        
        logger.debug(f"项目 {self.current_item_id} 下移")
    
    def _on_add_item(self):
        """添加新项目"""
        # 生成新ID
        new_id = self._generate_unique_id()
        
        # 创建新项目（由子类实现）
        new_item = self._create_new_item(new_id)
        
        self.items[new_id] = new_item
        self.item_order.append(new_id)
        self._refresh_item_list()
        self._select_item(new_id)
        self._mark_changed()
        
        # 聚焦到名称输入框
        if self.name_entry:
            self.name_entry.focus_set()
            self.name_entry.select_range(0, tk.END)
        
        logger.info(f"添加新项目: {new_id}")
    
    def _on_copy_item(self):
        """复制当前项目"""
        if not self.current_item_id:
            self._show_info("提示", "请先选择一个项目")
            return
        
        source_item = self.items.get(self.current_item_id)
        if not source_item:
            return
        
        # 生成新ID
        new_id = self._generate_unique_id()
        
        # 复制项目（由子类实现）
        new_item = self._copy_item(source_item, new_id)
        
        # 在当前项目后面插入
        current_index = self.item_order.index(self.current_item_id)
        self.items[new_id] = new_item
        self.item_order.insert(current_index + 1, new_id)
        
        self._refresh_item_list()
        self._select_item(new_id)
        self._mark_changed()
        
        logger.info(f"复制项目: {self.current_item_id} -> {new_id}")
    
    def _on_delete_item(self):
        """删除当前项目"""
        if not self.current_item_id:
            self._show_info("提示", "请先选择一个项目")
            return
        
        item = self.items.get(self.current_item_id)
        if not item:
            return
        
        # 检查是否为系统项目
        if self._is_item_system(item):
            self._show_info("无法删除", "系统预设项目不能删除，只能禁用")
            return
        
        # 检查是否可以删除（由子类实现额外检查）
        if not self._can_delete_item(item):
            return
        
        # 确认删除
        item_name = self._get_item_name(item)
        if not self._show_confirm("确认删除", f"确定要删除 '{item_name}' 吗？"):
            return
        
        # 获取当前索引，用于删除后选择下一个
        current_index = self.item_order.index(self.current_item_id) if self.current_item_id in self.item_order else -1
        
        # 删除项目
        deleted_id = self.current_item_id
        del self.items[self.current_item_id]
        if self.current_item_id in self.item_order:
            self.item_order.remove(self.current_item_id)
        
        logger.info(f"删除项目: {deleted_id}")
        
        # 选择另一个项目
        if self.item_order:
            # 尝试选择下一个，如果没有则选择上一个
            next_index = min(current_index, len(self.item_order) - 1)
            self._refresh_item_list()
            self._select_item(self.item_order[next_index])
        else:
            self.current_item_id = None
            self._refresh_item_list()
            self._on_no_item_selected()
        
        self._mark_changed()
    
    # ==================== 事件处理方法 ====================
    
    def _on_close(self):
        """关闭对话框"""
        if self.has_unsaved_changes:
            if not self._show_confirm("确认关闭", "有未保存的修改，确定要关闭吗？"):
                return
        
        self._cleanup()
        self.destroy()
    
    def _cleanup(self):
        """
        清理资源（子类可覆盖）
        
        在对话框关闭前调用，用于清理绑定的事件等资源。
        """
        # 解除鼠标滚轮绑定
        try:
            if self.item_canvas:
                self.item_canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
    
    # ==================== 抽象方法（子类必须实现） ====================
    
    @abstractmethod
    def _on_save(self):
        """保存配置"""
        pass
    
    @abstractmethod
    def _load_config_data(self):
        """加载配置数据"""
        pass
    
    @abstractmethod
    def _create_interface(self):
        """创建界面"""
        pass
    
    @abstractmethod
    def _create_new_item(self, item_id: str) -> Any:
        """创建新项目对象"""
        pass
    
    @abstractmethod
    def _copy_item(self, source_item: Any, new_id: str) -> Any:
        """复制项目对象"""
        pass
    
    # ==================== 项目属性访问方法（子类应覆盖） ====================
    
    def _get_item_name(self, item: Any) -> str:
        """获取项目名称"""
        return getattr(item, 'name', str(item))
    
    def _set_item_name(self, item: Any, name: str):
        """设置项目名称"""
        if hasattr(item, 'name'):
            item.name = name
    
    def _get_item_description(self, item: Any) -> str:
        """获取项目描述"""
        return getattr(item, 'description', '')
    
    def _set_item_description(self, item: Any, description: str):
        """设置项目描述"""
        if hasattr(item, 'description'):
            item.description = description
    
    def _get_item_enabled(self, item: Any) -> bool:
        """获取项目启用状态"""
        return getattr(item, 'enabled', True)
    
    def _set_item_enabled(self, item: Any, enabled: bool):
        """设置项目启用状态"""
        if hasattr(item, 'enabled'):
            item.enabled = enabled
    
    def _is_item_system(self, item: Any) -> bool:
        """检查是否为系统项目"""
        return getattr(item, 'is_system', False)
    
    def _can_delete_item(self, item: Any) -> bool:
        """
        检查是否可以删除项目（子类可覆盖添加额外检查）
        
        返回 True 表示可以删除，返回 False 表示不能删除
        """
        return True
    
    def _on_item_selected(self, item: Any):
        """
        项目被选中后的回调（子类可覆盖）
        
        用于更新项目特有的编辑区域
        """
        pass
    
    def _on_no_item_selected(self):
        """
        没有项目被选中时的回调（子类可覆盖）
        
        用于清空编辑区域
        """
        pass
    
    # ==================== 对话框辅助方法 ====================
    
    def _show_info(self, title: str, message: str):
        """显示信息对话框"""
        try:
            from gongwen_converter.gui.components.base_dialog import MessageBox
            MessageBox.showinfo(title, message, parent=self)
        except Exception:
            import tkinter.messagebox as msgbox
            msgbox.showinfo(title, message, parent=self)
    
    def _show_error(self, title: str, message: str):
        """显示错误对话框"""
        try:
            from gongwen_converter.gui.components.base_dialog import MessageBox
            MessageBox.showerror(title, message, parent=self)
        except Exception:
            import tkinter.messagebox as msgbox
            msgbox.showerror(title, message, parent=self)
    
    def _show_confirm(self, title: str, message: str) -> bool:
        """显示确认对话框"""
        try:
            from gongwen_converter.gui.components.base_dialog import MessageBox
            return MessageBox.askyesno(title, message, parent=self)
        except Exception:
            import tkinter.messagebox as msgbox
            return msgbox.askyesno(title, message, parent=self)
