"""
标题序号方案编辑器模块

提供标题序号方案的可视化编辑功能，支持：
- 方案列表管理（新增、删除、复制）
- 级别格式配置（9个级别）
- 实时预览
- 格式验证
- 默认方案设置
"""

import logging
import re
import tkinter as tk
from typing import Dict, List, Tuple, Optional, Callable, Any
from dataclasses import dataclass, field

import ttkbootstrap as tb
from ttkbootstrap.scrolled import ScrolledFrame

from gongwen_converter.gui.settings.base_editor_dialog import BaseEditorDialog

logger = logging.getLogger(__name__)


# ==================== 数据类 ====================

@dataclass
class NumberingScheme:
    """序号方案数据类"""
    scheme_id: str          # 方案ID，如 "gongwen_standard"
    name: str               # 显示名称
    description: str        # 说明
    enabled: bool           # 是否启用
    is_system: bool         # 是否系统预设
    levels: Dict[int, str] = field(default_factory=dict)  # 级别格式，{1: "{1.chinese_lower}、", ...}
    
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
            name=config.get("name", "未命名方案"),
            description=config.get("description", ""),
            enabled=config.get("enabled", True),
            is_system=config.get("is_system", False),
            levels=levels
        )
    
    def to_config(self) -> Dict:
        """转换为配置字典格式"""
        config = {
            "name": self.name,
            "description": self.description,
            "enabled": self.enabled,
            "is_system": self.is_system,
        }
        for level, format_str in self.levels.items():
            config[f"level_{level}"] = {"format": format_str}
        return config
    
    def copy(self, new_id: str, new_name: str) -> 'NumberingScheme':
        """复制方案"""
        return NumberingScheme(
            scheme_id=new_id,
            name=new_name,
            description=self.description,
            enabled=True,
            is_system=False,
            levels=self.levels.copy()
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
    WINDOW_TITLE = "序号方案编辑器"
    CONFIG_FILE_NAME = "heading_numbering_add.toml"
    WINDOW_WIDTH = 1000
    WINDOW_HEIGHT = 700
    MIN_WIDTH = 850
    MIN_HEIGHT = 550
    LEFT_PANEL_WIDTH = 280
    LEFT_PANEL_TITLE = "方案列表"
    
    # 预览区高度
    PREVIEW_HEIGHT = 180
    
    # 级别中文名称
    LEVEL_NAMES = [
        "一级标题", "二级标题", "三级标题",
        "四级标题", "五级标题", "六级标题",
        "七级标题", "八级标题", "九级标题"
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
            from gongwen_converter.config.toml_operations import read_toml_file
            
            config_data = read_toml_file(str(self.config_file_path))
            if not config_data:
                logger.warning("配置文件为空，使用默认配置")
                config_data = {}
            
            # 加载设置
            settings = config_data.get("settings", {})
            self.default_scheme_id = settings.get("default_scheme", "gongwen_standard")
            
            # 加载数字样式
            self.number_styles = config_data.get("number_styles", {})
            
            # 创建验证器
            self.validator = FormatValidator(self.number_styles)
            
            # 加载方案顺序
            self.item_order = settings.get("order", [])
            
            # 加载方案
            schemes_data = config_data.get("schemes", {})
            for scheme_id, scheme_config in schemes_data.items():
                scheme = NumberingScheme.from_config(scheme_id, scheme_config)
                self.items[scheme_id] = scheme
            
            # 确保 order 中的所有方案都存在
            self.item_order = [sid for sid in self.item_order if sid in self.items]
            
            # 添加不在 order 中的方案
            for scheme_id in self.items:
                if scheme_id not in self.item_order:
                    self.item_order.append(scheme_id)
            
            logger.info(f"加载了 {len(self.items)} 个序号方案")
            
        except Exception as e:
            logger.error(f"加载配置数据失败: {e}", exc_info=True)
            self._show_error("加载失败", f"无法加载配置文件：{e}")
    
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
            text="默认方案:",
            font=(self.small_font, self.small_size)
        ).pack(side="left")
        
        # 获取方案名称列表
        scheme_names = [s.name for s in self.items.values()]
        
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
            self.default_scheme_var.set(self.items[self.default_scheme_id].name)
    
    def _refresh_item_list(self):
        """刷新方案列表（覆盖基类方法以更新下拉框）"""
        super()._refresh_item_list()
        
        # 更新默认方案下拉框
        if self.default_scheme_combo:
            scheme_names = [self.items[sid].name for sid in self.item_order if sid in self.items]
            self.default_scheme_combo.configure(values=scheme_names)
    
    def _create_right_panel(self, parent: tb.Frame):
        """创建右侧面板（配置详情）"""
        right_frame = tb.Frame(parent)
        right_frame.pack(side="left", fill="both", expand=True)
        
        # 基本信息（使用基类方法）
        self._create_basic_info_section(right_frame)
        
        # 级别配置（可滚动）
        self._create_level_config_section(right_frame)
        
        # 预览区域
        self._create_preview_section(right_frame)
    
    def _create_level_config_section(self, parent: tb.Frame):
        """创建级别配置区域"""
        frame = tb.Labelframe(parent, text="级别格式配置", padding=5)
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
            text="📖 格式说明",
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
            text=f"{self.LEVEL_NAMES[level-1]}:",
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
        frame = tb.Labelframe(parent, text="实时预览", padding=5)
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
            name="新建方案",
            description="",
            enabled=True,
            is_system=False,
            levels={}
        )
    
    def _copy_item(self, source_item: NumberingScheme, new_id: str) -> NumberingScheme:
        """复制方案"""
        return source_item.copy(new_id, f"{source_item.name} (副本)")
    
    def _can_delete_item(self, item: NumberingScheme) -> bool:
        """检查是否可以删除（额外检查默认方案）"""
        if item.scheme_id == self.default_scheme_id:
            self._show_info("无法删除", "当前方案是默认方案，请先更改默认方案")
            return False
        return True
    
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
        
        # 查找对应的ID
        for scheme_id, scheme in self.items.items():
            if scheme.name == selected_name:
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
            from gongwen_converter.utils.heading_numbering import HeadingFormatter
            
            # 创建临时格式化器
            scheme_config = scheme.to_config()
            formatter = HeadingFormatter(scheme_config)
            
            # 生成预览文本
            preview_lines = []
            
            # 显示有格式定义的级别
            shown = 0
            for level in range(1, 10):
                format_str = scheme.levels.get(level, "")
                if format_str.strip():
                    formatter.increment_level(level)
                    formatted = formatter.format_heading(level)
                    indent = "  " * (level - 1)
                    preview_lines.append(f"{indent}{formatted}（示例{self.LEVEL_NAMES[level-1]}）")
                    shown += 1
                    
                    if shown >= 6:  # 最多显示6级
                        break
            
            if not preview_lines:
                preview_lines.append("（暂无格式定义）")
            
            # 更新预览文本框
            self.preview_text.configure(state="normal")
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert("1.0", "\n".join(preview_lines))
            self.preview_text.configure(state="disabled")
            
        except Exception as e:
            logger.warning(f"更新预览失败: {e}")
            self.preview_text.configure(state="normal")
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert("1.0", f"预览生成失败: {e}")
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
            label=f"── 当前级别 ({level}级) ──",
            state="disabled"
        )
        
        styles = [
            ("chinese_lower", "小写中文 (一二三)"),
            ("chinese_upper", "大写中文 (壹贰叁)"),
            ("arabic_half", "阿拉伯数字 (123)"),
            ("arabic_circled", "带圈数字 (①②③)"),
            ("letter_upper", "大写字母 (ABC)"),
            ("roman_upper", "罗马数字 (I II III)"),
        ]
        
        for style_id, style_name in styles:
            placeholder = f"{{{level}.{style_id}}}"
            menu.add_command(
                label=f"  {placeholder}  {style_name}",
                command=lambda p=placeholder, e=entry: self._insert_placeholder(e, p)
            )
        
        # 常用装饰
        menu.add_separator()
        menu.add_command(label="── 常用装饰 ──", state="disabled")
        
        decorations = [
            ("、", "顿号"),
            (". ", "点号+空格"),
            ("（）", "全角括号"),
            ("第...章", "章节格式"),
        ]
        
        for dec, desc in decorations:
            menu.add_command(
                label=f"  {dec}  {desc}",
                command=lambda d=dec, e=entry: self._insert_placeholder(e, d)
            )
        
        # 层级格式
        if level > 1:
            menu.add_separator()
            menu.add_command(label="── 层级格式 ──", state="disabled")
            
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
        help_text = """占位符语法说明

基本格式: {级别.样式}

级别范围: 1-9
  • 1 表示一级标题
  • 2 表示二级标题
  • ... 以此类推

可用样式:
  • chinese_lower  - 小写中文 (一、二、三)
  • chinese_upper  - 大写中文 (壹、贰、叁)
  • arabic_half    - 半角数字 (1, 2, 3)
  • arabic_full    - 全角数字 (１、２、３)
  • arabic_circled - 带圈数字 (①②③, 最多50)
  • letter_upper   - 大写字母 (A, B, C)
  • letter_lower   - 小写字母 (a, b, c)
  • roman_upper    - 大写罗马 (I, II, III)
  • roman_lower    - 小写罗马 (i, ii, iii)

示例:
  • {1.chinese_lower}、     → 一、二、三、
  • 第{2.chinese_lower}章　 → 第一章、第二章
  • {1.arabic_half}.{2.arabic_half}  → 1.1、1.2、2.1
  • （{3.chinese_lower}）   → （一）（二）（三）
  • {4.arabic_circled}      → ①②③

提示:
  1. 可以添加任意装饰文本
  2. 支持跨级引用（如3级引用1级）
  3. 空格式表示不显示序号
  4. 全角空格用于"编、章、节、条"后"""
        
        self._show_info("格式说明", help_text)
    
    # ==================== 保存 ====================
    
    def _on_save(self):
        """保存配置"""
        logger.info("保存序号方案配置")
        
        try:
            from gongwen_converter.config.toml_operations import (
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
                scheme_table["name"] = scheme.name
                scheme_table["description"] = scheme.description
                scheme_table["enabled"] = scheme.enabled
                scheme_table["is_system"] = scheme.is_system
                
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
                self._show_error("保存失败", "无法写入配置文件")
                
        except Exception as e:
            logger.error(f"保存配置失败: {e}", exc_info=True)
            self._show_error("保存失败", f"保存时发生错误：{e}")
