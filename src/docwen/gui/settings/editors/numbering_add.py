"""
标题序号方案编辑器模块

提供标题序号方案的可视化编辑功能，支持：
- 方案列表管理（新增、删除、复制）
- 级别格式配置（9个级别）
- 实时预览
- 格式验证
- 默认方案设置

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。
"""

import logging
import re
import tkinter as tk
from typing import Dict, List, Tuple, Optional, Callable, Any
from dataclasses import dataclass, field

import ttkbootstrap as tb
from ttkbootstrap.scrolled import ScrolledFrame

from docwen.gui.settings.editors.base import BaseEditorDialog
from docwen.i18n import t

logger = logging.getLogger(__name__)


# ==================== 数据类 ====================

@dataclass
class NumberingScheme:
    """序号方案数据类"""
    scheme_id: str          # 方案ID，如 "gongwen_standard"
    name: str               # 显示名称（用户自定义方案使用）
    description: str        # 说明（用户自定义方案使用）
    enabled: bool           # 是否启用
    is_system: bool         # 是否系统预设
    levels: Dict[int, str] = field(default_factory=dict)  # 级别格式
    locales: List[str] = field(default_factory=lambda: ["*"])  # 适用语言
    name_key: str = ""      # 翻译键（系统通用方案使用）
    description_key: str = ""  # 说明翻译键（系统通用方案使用）
    
    @classmethod
    def from_config(cls, scheme_id: str, config: Dict) -> 'NumberingScheme':
        """从配置字典创建方案对象"""
        levels = {}
        for i in range(1, 10):
            level_key = f"level_{i}"
            if level_key in config:
                levels[i] = config[level_key].get("format", "")
        
        return cls(
            scheme_id=scheme_id,
            name=config.get("name", ""),
            description=config.get("description", ""),
            enabled=config.get("enabled", True),
            is_system=config.get("is_system", False),
            levels=levels,
            locales=config.get("locales", ["*"]),
            name_key=config.get("name_key", ""),
            description_key=config.get("description_key", "")
        )
    
    def get_display_name(self) -> str:
        """获取显示名称（支持国际化）"""
        if self.name_key:
            translated = t(f"editors.numbering_add.names.{self.name_key}")
            # 检查翻译是否成功（翻译失败时通常返回原键或带方括号的键）
            if not translated.startswith("[") and translated != f"editors.numbering_add.names.{self.name_key}":
                return translated
        return self.name if self.name else self.scheme_id
    
    def get_display_description(self) -> str:
        """获取显示说明（支持国际化）"""
        if self.description_key:
            translated = t(f"editors.numbering_add.descriptions.{self.description_key}")
            if translated != f"editors.numbering_add.descriptions.{self.description_key}":
                return translated
        return self.description
    
    def to_config(self) -> Dict:
        """转换为配置字典格式"""
        config = {
            "description": self.description,
            "enabled": self.enabled,
            "is_system": self.is_system,
            "locales": self.locales,
        }
        # 根据是否有 name_key 决定保存方式
        if self.name_key:
            config["name_key"] = self.name_key
        else:
            config["name"] = self.name
        
        for level, format_str in self.levels.items():
            config[f"level_{level}"] = {"format": format_str}
        return config
    
    def copy(self, new_id: str, new_name: str) -> 'NumberingScheme':
        """复制方案（用户自定义，不使用 name_key）"""
        return NumberingScheme(
            scheme_id=new_id,
            name=new_name,
            description=self.description,
            enabled=True,
            is_system=False,
            levels=self.levels.copy(),
            locales=["*"],  # 复制的方案默认所有语言可用
            name_key=""     # 用户自定义方案不使用翻译键
        )


# ==================== 格式验证器 ====================

class FormatValidator:
    """格式验证器"""
    
    PLACEHOLDER_PATTERN = re.compile(r'\{(\d+)\.(\w+)\}')
    
    # 可用的数字样式
    VALID_STYLES = {
        'chinese_lower', 'chinese_upper',
        'arabic_half', 'arabic_full', 'arabic_circled',
        'letter_upper', 'letter_lower',
        'roman_upper', 'roman_lower'
    }
    
    def __init__(self, number_styles: Optional[Dict[str, Dict]] = None):
        """初始化验证器"""
        if number_styles:
            self.valid_styles = set(number_styles.keys())
        else:
            self.valid_styles = self.VALID_STYLES
    
    def validate_format(self, format_str: str, level: int) -> Tuple[bool, str]:
        """
        验证格式字符串
        
        参数:
            format_str: 格式字符串
            level: 当前级别（1-9）
            
        返回:
            Tuple[bool, str]: (是否有效, 错误/警告信息)
        """
        # 空格式是有效的（表示该级别不显示序号）
        if not format_str.strip():
            return True, ""
        
        # 查找所有占位符
        placeholders = self.PLACEHOLDER_PATTERN.findall(format_str)
        
        if not placeholders:
            # 没有占位符，可能是纯文本（允许但警告）
            return True, "⚠️ 格式中没有占位符，将显示固定文本"
        
        # 验证每个占位符
        for ref_level_str, style in placeholders:
            # 验证级别
            try:
                ref_level = int(ref_level_str)
                if not (1 <= ref_level <= 9):
                    return False, f"❌ 级别 {ref_level} 超出范围（应为1-9）"
            except ValueError:
                return False, f"❌ 级别 '{ref_level_str}' 不是有效数字"
            
            # 验证样式
            if style not in self.valid_styles:
                valid_list = ', '.join(sorted(self.valid_styles))
                return False, f"❌ 未知样式 '{style}'\n可用: {valid_list}"
        
        return True, "✓ 格式有效"
    
    def extract_placeholders(self, format_str: str) -> List[Tuple[int, str]]:
        """提取格式字符串中的所有占位符"""
        placeholders = self.PLACEHOLDER_PATTERN.findall(format_str)
        return [(int(level), style) for level, style in placeholders]


# ==================== 主编辑器对话框 ====================

class HeadingNumberingEditorDialog(BaseEditorDialog):
    """
    标题序号方案编辑器对话框
    
    提供标题序号方案的可视化编辑功能。
    """
    
    # 覆盖基类属性
    WINDOW_TITLE = t("editors.numbering_add.window_title")
    CONFIG_FILE_NAME = "heading_numbering_add.toml"
    WINDOW_WIDTH = 1000
    WINDOW_HEIGHT = 700
    MIN_WIDTH = 850
    MIN_HEIGHT = 550
    LEFT_PANEL_WIDTH = 280
    LEFT_PANEL_TITLE = t("editors.numbering_add.scheme_list")
    
    # 预览区高度
    PREVIEW_HEIGHT = 180
    
    @staticmethod
    def get_level_names() -> List[str]:
        """获取级别名称列表（支持国际化）"""
        return [
            t("editors.numbering_add.level_1"),
            t("editors.numbering_add.level_2"),
            t("editors.numbering_add.level_3"),
            t("editors.numbering_add.level_4"),
            t("editors.numbering_add.level_5"),
            t("editors.numbering_add.level_6"),
            t("editors.numbering_add.level_7"),
            t("editors.numbering_add.level_8"),
            t("editors.numbering_add.level_9"),
        ]
    
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
        # 调用基类初始化
        super().__init__(parent, config_manager, on_save)
        
        # 特有数据
        self.number_styles: Dict[str, Dict] = {}
        self.default_scheme_id: str = "gongwen_standard"
        
        # UI组件引用
        self.default_scheme_combo: Optional[tb.Combobox] = None
        self.default_scheme_var: Optional[tk.StringVar] = None
        self.level_entries: Dict[int, tb.Entry] = {}
        self.level_status_labels: Dict[int, tb.Label] = {}
        self.preview_text: Optional[tb.Text] = None
        
        # 格式验证器
        self.validator: Optional[FormatValidator] = None
        
        # 延迟更新定时器ID
        self._preview_update_id: Optional[str] = None
        
        # 加载配置数据
        self._load_config_data()
        
        # 配置窗口
        self._configure_window()
        
        # 创建界面
        self._create_interface()
        
        # 选择第一个方案
        if self.item_order:
            self._select_item(self.item_order[0])
        
        logger.info("序号方案编辑器对话框初始化完成")
    
    # ==================== 数据加载 ====================
    
    def _load_config_data(self):
        """加载配置数据"""
        logger.debug("加载序号方案配置数据")
        
        try:
            from docwen.config.toml_operations import read_toml_file
            from docwen.i18n import get_current_locale
            
            config_data = read_toml_file(str(self.config_file_path))
            if not config_data:
                logger.warning("配置文件为空，使用默认配置")
                config_data = {}
            
            # 获取当前语言
            current_locale = get_current_locale()
            logger.debug(f"当前语言: {current_locale}")
            
            # 加载设置
            settings = config_data.get("settings", {})
            self.default_scheme_id = settings.get("default_scheme", "gongwen_standard")
            
            # 加载数字样式
            self.number_styles = config_data.get("number_styles", {})
            
            # 创建验证器
            self.validator = FormatValidator(self.number_styles)
            
            # 加载方案顺序
            self.item_order = settings.get("order", [])
            
            # 加载方案（按语言过滤）
            schemes_data = config_data.get("schemes", {})
            for scheme_id, scheme_config in schemes_data.items():
                scheme = NumberingScheme.from_config(scheme_id, scheme_config)
                
                # 检查语言适配性
                locales = scheme.locales
                if "*" in locales or current_locale in locales:
                    # 方案适用于当前语言
                    self.items[scheme_id] = scheme
                else:
                    # 方案不适用于当前语言，跳过
                    logger.debug(f"跳过方案 {scheme_id}，不适用于当前语言 {current_locale} (适用语言: {locales})")
            
            # 确保 order 中的所有方案都存在（过滤掉不适用的）
            self.item_order = [sid for sid in self.item_order if sid in self.items]
            
            # 添加不在 order 中的方案
            for scheme_id in self.items:
                if scheme_id not in self.item_order:
                    self.item_order.append(scheme_id)
            
            # 检查默认方案是否可用
            if self.default_scheme_id not in self.items and self.item_order:
                # 默认方案不适用于当前语言，选择第一个可用的
                self.default_scheme_id = self.item_order[0]
                logger.info(f"默认方案不可用，自动选择: {self.default_scheme_id}")
            
            logger.info(f"加载了 {len(self.items)} 个适用于当前语言的序号方案")
            
        except Exception as e:
            logger.error(f"加载配置数据失败: {e}", exc_info=True)
            self._show_error(t("editors.numbering_add.load_failed"), t("editors.numbering_add.load_failed_message", error=e))
    
    # ==================== UI创建方法 ====================
    
    def _create_interface(self):
        """创建界面"""
        logger.debug("创建序号方案编辑器界面")
        
        # 主容器
        main_frame = tb.Frame(self, padding=self.PADDING)
        main_frame.pack(fill="both", expand=True)
        
        # 内容区域（上部）
        content_frame = tb.Frame(main_frame)
        content_frame.pack(fill="both", expand=True)
        
        # 创建左右分栏
        self._create_left_panel_content(content_frame)
        self._create_right_panel(content_frame)
        
        # 创建底部按钮（在内容区域之后）
        self._create_bottom_buttons(main_frame)
        
        logger.debug("序号方案编辑器界面创建完成")
    
    def _create_left_panel_content(self, parent: tb.Frame):
        """创建左侧面板内容"""
        left_frame = self._create_left_panel(parent)
        
        # 默认方案选择
        self._create_default_scheme_selector(left_frame)
        
        # 方案列表
        self._create_item_list(left_frame)
        
        # 刷新列表
        self._refresh_item_list()
        
        # 操作按钮
        self._create_action_buttons(left_frame)
    
    def _create_default_scheme_selector(self, parent: tb.Frame):
        """创建默认方案选择器"""
        frame = tb.Frame(parent)
        frame.pack(fill="x", pady=(0, 10))
        
        tb.Label(
            frame,
            text=t("editors.numbering_add.default_scheme"),
            font=(self.small_font, self.small_size)
        ).pack(side="left")
        
        # 获取方案名称列表（使用国际化显示名称）
        scheme_names = [s.get_display_name() for s in self.items.values()]
        
        self.default_scheme_var = tk.StringVar()
        self.default_scheme_combo = tb.Combobox(
            frame,
            textvariable=self.default_scheme_var,
            values=scheme_names,
            state="readonly",
            width=18
        )
        self.default_scheme_combo.pack(side="left", padx=(5, 0), fill="x", expand=True)
        self.default_scheme_combo.bind("<<ComboboxSelected>>", self._on_default_scheme_changed)
        
        # 设置当前默认方案
        if self.default_scheme_id in self.items:
            self.default_scheme_var.set(self.items[self.default_scheme_id].get_display_name())
    
    def _refresh_item_list(self):
        """刷新方案列表（覆盖基类方法以更新下拉框）"""
        super()._refresh_item_list()
        
        # 更新默认方案下拉框（使用国际化显示名称）
        if self.default_scheme_combo:
            scheme_names = [self.items[sid].get_display_name() for sid in self.item_order if sid in self.items]
            self.default_scheme_combo.configure(values=scheme_names)
    
    def _create_right_panel(self, parent: tb.Frame):
        """创建右侧面板（配置详情）"""
        right_frame = tb.Frame(parent)
        right_frame.pack(side="left", fill="both", expand=True)
        
        # 基本信息
        self._create_item_editor(right_frame)
        
        # 级别配置（可滚动）
        self._create_level_config_section(right_frame)
        
        # 预览区域
        self._create_preview_section(right_frame)
    
    def _create_item_editor(self, parent: tb.Frame):
        """创建项目编辑区（名称和说明）"""
        frame = tb.Labelframe(parent, text=t("editors.common.basic_info"), padding=5)
        frame.pack(fill="x", pady=(0, 10))
        
        # 名称
        name_frame = tb.Frame(frame)
        name_frame.pack(fill="x", pady=(0, 5))
        
        tb.Label(
            name_frame,
            text=t("editors.common.name"),
            width=12,
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
            text=t("editors.common.description"),
            width=12,
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
    
    def _create_level_config_section(self, parent: tb.Frame):
        """创建级别配置区域"""
        frame = tb.Labelframe(parent, text=t("editors.numbering_add.level_format_config"), padding=5)
        frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # 使用ScrolledFrame
        scrolled = ScrolledFrame(frame, autohide=True)
        scrolled.pack(fill="both", expand=True)
        
        inner_frame = scrolled
        
        # 创建9个级别的输入框
        for level in range(1, 10):
            self._create_level_row(inner_frame, level)
        
        # 格式说明按钮
        help_frame = tb.Frame(frame)
        help_frame.pack(fill="x", pady=(5, 0))
        
        tb.Button(
            help_frame,
            text=t("editors.numbering_add.format_help"),
            command=self._show_format_help,
            bootstyle="link"
        ).pack(side="right")
    
    def _create_level_row(self, parent: tb.Frame, level: int):
        """创建单个级别的配置行"""
        row_frame = tb.Frame(parent)
        row_frame.pack(fill="x", pady=2)
        
        # 级别标签
        tb.Label(
            row_frame,
            text=f"{self.get_level_names()[level-1]}:",
            width=10,
            anchor="e",
            font=(self.small_font, self.small_size)
        ).pack(side="left")
        
        # 格式输入框
        entry_var = tk.StringVar()
        entry = tb.Entry(
            row_frame,
            textvariable=entry_var,
            width=35,
            font=(self.small_font, self.small_size)
        )
        entry.pack(side="left", padx=(5, 5), fill="x", expand=True)
        entry_var.trace_add("write", lambda *_, lv=level: self._on_level_format_changed(lv))
        
        self.level_entries[level] = entry
        
        # 快速插入按钮
        tb.Button(
            row_frame,
            text="📋",
            command=lambda lv=level: self._show_placeholder_menu(lv),
            bootstyle="link",
            width=3
        ).pack(side="left")
        
        # 验证状态标签
        status_label = tb.Label(
            row_frame,
            text="",
            width=3,
            font=(self.small_font, self.small_size)
        )
        status_label.pack(side="left")
        self.level_status_labels[level] = status_label
    
    def _create_preview_section(self, parent: tb.Frame):
        """创建预览区域"""
        frame = tb.Labelframe(parent, text=t("editors.numbering_add.realtime_preview"), padding=5)
        frame.pack(fill="x")
        frame.configure(height=self.scale(self.PREVIEW_HEIGHT))
        frame.pack_propagate(False)
        
        self.preview_text = tb.Text(
            frame,
            height=8,
            font=(self.small_font, self.small_size),
            state="disabled",
            wrap="word"
        )
        self.preview_text.pack(fill="both", expand=True)
    
    # ==================== 项目操作方法（实现基类抽象方法） ====================
    
    def _create_new_item(self, item_id: str) -> NumberingScheme:
        """创建新方案"""
        return NumberingScheme(
            scheme_id=item_id,
            name=t("editors.numbering_add.new_scheme"),
            description="",
            enabled=True,
            is_system=False,
            levels={}
        )
    
    def _copy_item(self, source_item: NumberingScheme, new_id: str) -> NumberingScheme:
        """复制方案"""
        copy_suffix = t("editors.numbering_add.copy_suffix")
        return source_item.copy(new_id, f"{source_item.get_display_name()} {copy_suffix}")
    
    def _can_delete_item(self, item: NumberingScheme) -> bool:
        """检查是否可以删除（额外检查默认方案）"""
        if item.scheme_id == self.default_scheme_id:
            self._show_info(t("editors.numbering_add.cannot_delete"), t("editors.numbering_add.cannot_delete_default_scheme"))
            return False
        return True
    
    def _get_item_name(self, item: NumberingScheme) -> str:
        """获取方案显示名称（覆盖基类方法以支持国际化）"""
        return item.get_display_name()
    
    def _get_item_description(self, item: NumberingScheme) -> str:
        """获取方案显示说明（覆盖基类方法以支持国际化）"""
        return item.get_display_description()
    
    def _on_item_selected(self, item: NumberingScheme):
        """方案被选中后的回调"""
        # 更新级别格式
        for level in range(1, 10):
            entry = self.level_entries.get(level)
            if entry:
                entry.delete(0, tk.END)
                format_str = item.levels.get(level, "")
                entry.insert(0, format_str)
        
        # 验证并更新预览
        self._validate_all_levels()
        self._update_preview()
    
    def _on_no_item_selected(self):
        """没有方案被选中时清空编辑区"""
        self._is_loading = True
        try:
            if self.name_var:
                self.name_var.set("")
            if self.description_text:
                self.description_text.delete("1.0", tk.END)
            for level in range(1, 10):
                entry = self.level_entries.get(level)
                if entry:
                    entry.delete(0, tk.END)
            if self.preview_text:
                self.preview_text.configure(state="normal")
                self.preview_text.delete("1.0", tk.END)
                self.preview_text.configure(state="disabled")
        finally:
            self._is_loading = False
    
    # ==================== 事件处理方法 ====================
    
    def _on_default_scheme_changed(self, event=None):
        """处理默认方案变更"""
        selected_name = self.default_scheme_var.get()
        
        # 查找对应的ID（使用国际化显示名称匹配）
        for scheme_id, scheme in self.items.items():
            if scheme.get_display_name() == selected_name:
                self.default_scheme_id = scheme_id
                self._mark_changed()
                logger.info(f"默认方案变更为: {scheme_id}")
                break
    
    def _on_level_format_changed(self, level: int):
        """处理级别格式变更"""
        # 加载中或无当前方案时不处理
        if self._is_loading or not self.current_item_id:
            return
        
        self._mark_changed()
        
        # 更新当前方案数据
        scheme = self.items.get(self.current_item_id)
        entry = self.level_entries.get(level)
        
        if scheme and entry:
            format_str = entry.get()
            scheme.levels[level] = format_str
            
            # 验证格式
            self._validate_level(level, format_str)
        
        # 延迟更新预览
        self._schedule_preview_update()
    
    def _validate_level(self, level: int, format_str: str):
        """验证单个级别的格式"""
        if not self.validator:
            return
        
        is_valid, message = self.validator.validate_format(format_str, level)
        status_label = self.level_status_labels.get(level)
        
        if status_label:
            if not format_str.strip():
                status_label.configure(text="", bootstyle="secondary")
            elif is_valid:
                if message.startswith("⚠️"):
                    status_label.configure(text="⚠️", bootstyle="warning")
                else:
                    status_label.configure(text="✓", bootstyle="success")
            else:
                status_label.configure(text="✗", bootstyle="danger")
    
    def _validate_all_levels(self):
        """验证所有级别"""
        if not self.current_item_id:
            return
        
        scheme = self.items.get(self.current_item_id)
        if not scheme:
            return
        
        for level in range(1, 10):
            format_str = scheme.levels.get(level, "")
            self._validate_level(level, format_str)
    
    def _schedule_preview_update(self):
        """安排预览更新（延迟500ms）"""
        if self._preview_update_id:
            self.after_cancel(self._preview_update_id)
        
        self._preview_update_id = self.after(500, self._update_preview)
    
    def _update_preview(self):
        """更新预览显示"""
        self._preview_update_id = None
        
        if not self.current_item_id or not self.preview_text:
            return
        
        scheme = self.items.get(self.current_item_id)
        if not scheme:
            return
        
        try:
            from docwen.utils.heading_numbering import HeadingFormatter
            
            # 创建临时格式化器
            scheme_config = scheme.to_config()
            formatter = HeadingFormatter(scheme_config)
            
            # 生成预览文本
            preview_lines = []
            level_names = self.get_level_names()
            
            # 找出最大有格式定义的级别
            max_defined_level = 0
            for level in range(1, 10):
                if scheme.levels.get(level, "").strip():
                    max_defined_level = level
            
            # 显示所有级别（包括空格式的级别，以展示完整层级结构）
            shown = 0
            for level in range(1, 10):
                # 如果当前级别超过最大定义级别，停止显示
                if level > max_defined_level:
                    break
                
                format_str = scheme.levels.get(level, "")
                indent = "  " * (level - 1)
                sample_text = t("editors.numbering_add.preview_sample", level_name=level_names[level-1])
                
                if format_str.strip():
                    # 有格式定义：显示带序号的标题
                    formatter.increment_level(level)
                    formatted = formatter.format_heading(level)
                    preview_lines.append(f"{indent}{formatted}（{sample_text}）")
                else:
                    # 空格式：显示无序号的标题（如"层级数字(H2起)"的一级标题）
                    preview_lines.append(f"{indent}（{sample_text}）")
                
                shown += 1
                if shown >= 7:  # 最多显示7级（包含可能的空格式级别）
                    break
            
            if not preview_lines:
                preview_lines.append(f"（{t('editors.numbering_add.no_format')}）")
            
            # 更新预览文本框
            self.preview_text.configure(state="normal")
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert("1.0", "\n".join(preview_lines))
            self.preview_text.configure(state="disabled")
            
        except Exception as e:
            logger.warning(f"更新预览失败: {e}")
            self.preview_text.configure(state="normal")
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert("1.0", f"{t('editors.numbering_add.preview_failed')}: {e}")
            self.preview_text.configure(state="disabled")
    
    # ==================== 占位符菜单 ====================
    
    def _show_placeholder_menu(self, level: int):
        """显示占位符快速插入菜单"""
        entry = self.level_entries.get(level)
        if not entry:
            return
        
        menu = tk.Menu(self, tearoff=0)
        
        # 当前级别占位符
        menu.add_command(
            label=t("editors.numbering_add.menu_current_level", level=level),
            state="disabled"
        )
        
        styles = [
            ("chinese_lower", t("editors.numbering_add.style_chinese_lower")),
            ("chinese_upper", t("editors.numbering_add.style_chinese_upper")),
            ("arabic_half", t("editors.numbering_add.style_arabic_half")),
            ("arabic_circled", t("editors.numbering_add.style_arabic_circled")),
            ("letter_upper", t("editors.numbering_add.style_letter_upper")),
            ("roman_upper", t("editors.numbering_add.style_roman_upper")),
        ]
        
        for style_id, style_name in styles:
            placeholder = f"{{{level}.{style_id}}}"
            menu.add_command(
                label=f"  {placeholder}  {style_name}",
                command=lambda p=placeholder, e=entry: self._insert_placeholder(e, p)
            )
        
        # 常用装饰
        menu.add_separator()
        menu.add_command(label=t("editors.numbering_add.menu_decorations"), state="disabled")
        
        decorations = [
            ("、", t("editors.numbering_add.decoration_dun")),
            (". ", t("editors.numbering_add.decoration_dot_space")),
            ("（）", t("editors.numbering_add.decoration_bracket")),
            ("第...章", t("editors.numbering_add.decoration_chapter")),
        ]
        
        for dec, desc in decorations:
            menu.add_command(
                label=f"  {dec}  {desc}",
                command=lambda d=dec, e=entry: self._insert_placeholder(e, d)
            )
        
        # 层级格式
        if level > 1:
            menu.add_separator()
            menu.add_command(label=t("editors.numbering_add.menu_hierarchical"), state="disabled")
            
            # 生成层级格式
            parts = [f"{{{i}.arabic_half}}" for i in range(1, level + 1)]
            hierarchical = ".".join(parts) + " "
            menu.add_command(
                label=f"  {hierarchical}",
                command=lambda p=hierarchical, e=entry: self._insert_placeholder(e, p)
            )
        
        # 显示菜单
        try:
            x = entry.winfo_rootx() + entry.winfo_width()
            y = entry.winfo_rooty()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()
    
    def _insert_placeholder(self, entry: tb.Entry, text: str):
        """在输入框中插入占位符"""
        try:
            # 获取当前光标位置
            cursor_pos = entry.index(tk.INSERT)
            
            # 插入文本
            entry.insert(cursor_pos, text)
            
            # 聚焦
            entry.focus_set()
        except Exception as e:
            logger.warning(f"插入占位符失败: {e}")
    
    # ==================== 格式说明 ====================
    
    def _show_format_help(self):
        """显示格式说明对话框"""
        help_text = t("editors.numbering_add.format_help_text")
        
        self._show_info(t("editors.numbering_add.format_help_title"), help_text)
    
    # ==================== 保存 ====================
    
    def _on_save(self):
        """保存配置"""
        logger.info("保存序号方案配置")
        
        try:
            from docwen.config.toml_operations import (
                read_toml_document,
                write_toml_document
            )
            from tomlkit import table, array
            
            # 读取现有文档（保留注释）
            doc = read_toml_document(str(self.config_file_path))
            if doc is None:
                from tomlkit import document
                doc = document()
            
            # 更新设置
            if "settings" not in doc:
                doc.add("settings", table())
            doc["settings"]["default_scheme"] = self.default_scheme_id
            
            # 更新方案顺序
            order_array = array()
            for scheme_id in self.item_order:
                order_array.append(scheme_id)
            order_array.multiline(True)
            doc["settings"]["order"] = order_array
            
            # 更新方案
            if "schemes" not in doc:
                doc.add("schemes", table())
            
            schemes_table = doc["schemes"]
            
            # 获取现有方案ID
            existing_ids = set(schemes_table.keys()) if schemes_table else set()
            current_ids = set(self.items.keys())
            
            # 删除已移除的自定义方案（保护系统方案）
            for scheme_id in existing_ids - current_ids:
                if not schemes_table[scheme_id].get("is_system", False):
                    del schemes_table[scheme_id]
                    logger.info(f"从配置文件中删除方案: {scheme_id}")
            
            # 更新/添加方案
            for scheme_id, scheme in self.items.items():
                if scheme_id not in schemes_table:
                    schemes_table.add(scheme_id, table())
                
                scheme_table = schemes_table[scheme_id]
                
                # 根据是否有 name_key 决定保存方式
                if scheme.name_key:
                    scheme_table["name_key"] = scheme.name_key
                    # 删除可能存在的旧 name 字段
                    if "name" in scheme_table:
                        del scheme_table["name"]
                else:
                    scheme_table["name"] = scheme.name
                    # 删除可能存在的旧 name_key 字段
                    if "name_key" in scheme_table:
                        del scheme_table["name_key"]
                
                scheme_table["description"] = scheme.description
                scheme_table["enabled"] = scheme.enabled
                scheme_table["is_system"] = scheme.is_system
                
                # 保存 locales
                locales_array = array()
                for locale in scheme.locales:
                    locales_array.append(locale)
                scheme_table["locales"] = locales_array
                
                # 更新级别格式
                for level in range(1, 10):
                    level_key = f"level_{level}"
                    format_str = scheme.levels.get(level, "")
                    
                    if level_key not in scheme_table:
                        scheme_table.add(level_key, table())
                    
                    scheme_table[level_key]["format"] = format_str
            
            # 写入文件
            success = write_toml_document(str(self.config_file_path), doc)
            
            if success:
                # 重新加载配置
                self.config_manager.reload_configs()
                
                # 调用回调
                if self.on_save_callback:
                    self.on_save_callback()
                
                self._clear_changed()
                logger.info("序号方案配置保存成功")
                
                # 关闭对话框
                self.destroy()
            else:
                self._show_error(t("editors.numbering_add.save_failed"), t("editors.numbering_add.save_failed_message"))
                
        except Exception as e:
            logger.error(f"保存配置失败: {e}", exc_info=True)
            self._show_error(t("editors.numbering_add.save_failed"), t("editors.numbering_add.save_failed_with_error", error=e))
