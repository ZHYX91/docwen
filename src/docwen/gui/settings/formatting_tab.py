"""
格式设置选项卡

实现设置对话框的格式设置选项卡，包含：
- DOCX → MD 格式设置
- MD → DOCX 格式设置
- Markdown语法选择

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。
使用 ConfigCombobox 组件实现配置值与显示文本的分离，
避免语言切换时的映射问题。
"""

import tkinter as tk
import ttkbootstrap as tb
from typing import Dict, Any, Callable
import logging

from .base_tab import BaseSettingsTab, SectionStyle
from docwen.gui.components.config_combobox import ConfigCombobox
from docwen.i18n import t

logger = logging.getLogger(__name__)

# ========== 配置值常量和翻译键映射 ==========

# DOCX → MD 格式保留选项（布尔值转字符串）
PRESERVE_FORMAT_CONFIG_VALUES = ["true", "false"]
PRESERVE_FORMAT_TRANSLATE_KEYS = {
    "true": "settings.formatting.options.preserve_format",
    "false": "settings.formatting.options.discard_format"
}

# MD → DOCX 格式处理模式
FORMATTING_MODE_CONFIG_VALUES = ["apply", "keep", "remove"]
FORMATTING_MODE_TRANSLATE_KEYS = {
    "apply": "settings.formatting.options.apply_format",
    "keep": "settings.formatting.options.keep_markup",
    "remove": "settings.formatting.options.clean_markup"
}

# DOCX → MD 分隔符映射选项
DOCX_TO_MD_HR_CONFIG_VALUES = ["---", "***", "___", "ignore"]
DOCX_TO_MD_HR_TRANSLATE_KEYS = {
    "---": "settings.formatting.separators.dash",
    "***": "settings.formatting.separators.asterisk",
    "___": "settings.formatting.separators.underscore",
    "ignore": "settings.formatting.separators.ignore"
}

# MD → DOCX 分隔符映射选项
MD_TO_DOCX_HR_CONFIG_VALUES = ["page_break", "section_break", "horizontal_rule_1", "horizontal_rule_2", "horizontal_rule_3", "ignore"]
MD_TO_DOCX_HR_TRANSLATE_KEYS = {
    "page_break": "settings.formatting.separators.page_break",
    "section_break": "settings.formatting.separators.section_break",
    "horizontal_rule_1": "settings.formatting.separators.horizontal_rule_1",
    "horizontal_rule_2": "settings.formatting.separators.horizontal_rule_2",
    "horizontal_rule_3": "settings.formatting.separators.horizontal_rule_3",
    "ignore": "settings.formatting.separators.ignore"
}

# 粗体/斜体语法选项
BOLD_ITALIC_CONFIG_VALUES = ["asterisk", "underscore"]
BOLD_ITALIC_TRANSLATE_KEYS = {
    "asterisk": "settings.formatting.syntax.bold_asterisk",
    "underscore": "settings.formatting.syntax.bold_underscore"
}
ITALIC_TRANSLATE_KEYS = {
    "asterisk": "settings.formatting.syntax.italic_asterisk",
    "underscore": "settings.formatting.syntax.italic_underscore"
}

# 扩展语法选项（删除线、高亮、上标、下标）
EXTENDED_HTML_CONFIG_VALUES = ["extended", "html"]
STRIKETHROUGH_TRANSLATE_KEYS = {
    "extended": "settings.formatting.syntax.strikethrough_extended",
    "html": "settings.formatting.syntax.strikethrough_html"
}
HIGHLIGHT_TRANSLATE_KEYS = {
    "extended": "settings.formatting.syntax.highlight_extended",
    "html": "settings.formatting.syntax.highlight_html"
}
SUPERSCRIPT_TRANSLATE_KEYS = {
    "extended": "settings.formatting.syntax.superscript_extended",
    "html": "settings.formatting.syntax.superscript_html"
}
SUBSCRIPT_TRANSLATE_KEYS = {
    "extended": "settings.formatting.syntax.subscript_extended",
    "html": "settings.formatting.syntax.subscript_html"
}

# 无序列表标记选项
UNORDERED_LIST_CONFIG_VALUES = ["dash", "asterisk", "plus"]
UNORDERED_LIST_TRANSLATE_KEYS = {
    "dash": "settings.formatting.syntax.unordered_dash",
    "asterisk": "settings.formatting.syntax.unordered_asterisk",
    "plus": "settings.formatting.syntax.unordered_plus"
}

# 有序列表编号选项
ORDERED_LIST_CONFIG_VALUES = ["restart", "preserve"]
ORDERED_LIST_TRANSLATE_KEYS = {
    "restart": "settings.formatting.syntax.ordered_restart",
    "preserve": "settings.formatting.syntax.ordered_preserve"
}

# 列表缩进空格数选项
INDENT_SPACES_CONFIG_VALUES = ["2", "4"]
INDENT_SPACES_TRANSLATE_KEYS = {
    "2": "settings.formatting.syntax.indent_2_spaces",
    "4": "settings.formatting.syntax.indent_4_spaces"
}

# 内置表格样式选项
BUILTIN_TABLE_STYLE_CONFIG_VALUES = ["three_line_table", "table_grid"]
BUILTIN_TABLE_STYLE_TRANSLATE_KEYS = {
    "three_line_table": "settings.formatting.table_styles.three_line_table",
    "table_grid": "settings.formatting.table_styles.table_grid"
}


class FormattingTab(BaseSettingsTab):
    """
    格式设置选项卡
    
    管理DOCX与Markdown互转时的格式处理配置。
    """
    
    def __init__(self, parent, config_manager, on_change_callback):
        """
        初始化格式设置选项卡
        
        参数:
            parent: 父容器
            config_manager: 配置管理器实例
            on_change_callback: 配置变更回调函数
        """
        super().__init__(parent, config_manager, on_change_callback)
        logger.debug("初始化格式设置选项卡")
    
    def _create_interface(self):
        """创建选项卡内容（实现抽象方法）"""
        logger.debug("创建格式设置选项卡内容")
        
        # 创建两个主要区域（语法选择已整合到文档转MD section中）
        self._create_docx_to_md_section()
        self._create_md_to_docx_section()
        
        logger.debug("格式设置选项卡内容创建完成")
    
    def _create_config_combobox(
        self,
        parent: tk.Widget,
        label_text: str,
        config_values: list,
        translate_keys: dict,
        initial_value: str,
        tooltip: str,
        on_change: Callable[[str], None],
        disabled: bool = False
    ) -> ConfigCombobox:
        """
        在指定列中创建带标签和信息图标的配置下拉框
        
        参数:
            parent: 父组件
            label_text: 标签文本
            config_values: 配置值列表
            translate_keys: 翻译键映射
            initial_value: 初始配置值
            tooltip: 工具提示文本
            on_change: 值变更回调函数
            disabled: 是否禁用下拉框
            
        返回:
            ConfigCombobox: 创建的下拉框组件
        """
        from docwen.utils.gui_utils import create_info_icon
        
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
        
        # 创建配置下拉框
        combobox = ConfigCombobox(
            container,
            config_values=config_values,
            translate_keys=translate_keys,
            initial_value=initial_value,
            on_change=on_change,
            state="disabled" if disabled else "readonly"
        )
        combobox.pack(fill="x")
        
        return combobox
    
    def _create_docx_to_md_section(self):
        """
        创建DOCX → MD 格式设置区域（三列布局）
        
        配置路径：conversion_config.docx_to_md.*
        """
        logger.debug("创建DOCX → MD 格式设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.formatting.docx_to_md_section"),
            SectionStyle.PRIMARY
        )
        
        # 获取当前配置
        preserve_formatting = self.config_manager.get_docx_to_md_preserve_formatting()
        preserve_heading_formatting = self.config_manager.get_docx_to_md_preserve_heading_formatting()
        preserve_table_header_formatting = self.config_manager.get_docx_to_md_preserve_table_header_formatting()
        
        # 格式处理小标题
        from docwen.utils.gui_utils import create_info_icon
        
        format_title_frame = tb.Frame(frame)
        format_title_frame.pack(fill="x", pady=(0, 10))
        
        format_title_label = tb.Label(
            format_title_frame,
            text=t("settings.formatting.format_processing"),
            font=(self.small_font, self.small_size, "bold")
        )
        format_title_label.pack(side="left")
        
        format_info = create_info_icon(
            format_title_frame,
            t("settings.formatting.format_processing_tooltip"),
            "info"
        )
        format_info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 创建三列容器
        columns_frame = tb.Frame(frame)
        columns_frame.pack(fill="x")
        
        # 配置列权重使三列等宽
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        columns_frame.columnconfigure(2, weight=1)
        
        # 第一列
        first_column = tb.Frame(columns_frame)
        first_column.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        # 第二列
        second_column = tb.Frame(columns_frame)
        second_column.grid(row=0, column=1, sticky="nsew", padx=(5, 5))
        
        # 第三列
        third_column = tb.Frame(columns_frame)
        third_column.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        
        # 第一列 - 正文格式（布尔值转字符串）
        self.preserve_formatting_combo = self._create_config_combobox(
            first_column,
            t("settings.formatting.body_format_label"),
            PRESERVE_FORMAT_CONFIG_VALUES,
            PRESERVE_FORMAT_TRANSLATE_KEYS,
            "true" if preserve_formatting else "false",
            t("settings.formatting.body_format_tooltip"),
            lambda v: self.on_change("preserve_formatting", v == "true")
        )
        
        # 第二列 - 小标题格式
        self.preserve_heading_formatting_combo = self._create_config_combobox(
            second_column,
            t("settings.formatting.heading_format_label"),
            PRESERVE_FORMAT_CONFIG_VALUES,
            PRESERVE_FORMAT_TRANSLATE_KEYS,
            "true" if preserve_heading_formatting else "false",
            t("settings.formatting.heading_format_tooltip"),
            lambda v: self.on_change("preserve_heading_formatting", v == "true")
        )
        
        # 第三列 - 表头格式
        self.preserve_table_header_formatting_combo = self._create_config_combobox(
            third_column,
            t("settings.formatting.table_header_format_label"),
            PRESERVE_FORMAT_CONFIG_VALUES,
            PRESERVE_FORMAT_TRANSLATE_KEYS,
            "true" if preserve_table_header_formatting else "false",
            t("settings.formatting.table_header_format_tooltip"),
            lambda v: self.on_change("preserve_table_header_formatting", v == "true")
        )
        
        # ========== 分隔符映射配置（文档转MD）==========
        self._create_docx_to_md_hr_config(frame)
        
        # ========== Markdown语法选择 ==========
        self._create_syntax_config(frame)
        
        logger.debug("DOCX → MD 格式设置区域创建完成")
    
    def _create_docx_to_md_hr_config(self, parent_frame):
        """
        创建文档转MD的分隔符映射配置
        
        配置路径：conversion_config.horizontal_rule.docx_to_md.*
        
        三种Word元素与三种MD分隔符的映射：
        - 分页符 (page_break) ↔ ---/***/___ 
        - 分节符 (section_break) ↔ ---/***/___ (所有分节符类型统一处理)
        - 分隔线 (horizontal_rule) ↔ ---/***/___ (段落底部边框)
        """
        from docwen.utils.gui_utils import create_info_icon
        
        # 分隔线
        separator = tb.Separator(parent_frame, orient="horizontal")
        separator.pack(fill="x", pady=(15, 10))
        
        # 小节标题
        title_frame = tb.Frame(parent_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            title_frame,
            text=t("settings.formatting.separator_mapping"),
            font=(self.small_font, self.small_size, "bold")
        )
        title_label.pack(side="left")
        
        info = create_info_icon(
            title_frame,
            t("settings.formatting.separator_mapping_tooltip"),
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 获取当前配置
        hr_config = self.config_manager.get_horizontal_rule_docx_to_md_config()
        
        # 创建三列容器
        columns_frame = tb.Frame(parent_frame)
        columns_frame.pack(fill="x")
        
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        columns_frame.columnconfigure(2, weight=1)
        
        # 第一列 - 分页符
        first_column = tb.Frame(columns_frame)
        first_column.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        # 第二列 - 分节符
        second_column = tb.Frame(columns_frame)
        second_column.grid(row=0, column=1, sticky="nsew", padx=(5, 5))
        
        # 第三列 - 分隔线
        third_column = tb.Frame(columns_frame)
        third_column.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        
        # 第一列 - 分页符
        self.docx_hr_page_break_combo = self._create_config_combobox(
            first_column,
            t("settings.formatting.page_break_label"),
            DOCX_TO_MD_HR_CONFIG_VALUES,
            DOCX_TO_MD_HR_TRANSLATE_KEYS,
            hr_config.get("page_break", "---"),
            t("settings.formatting.page_break_tooltip"),
            lambda v: self.on_change("docx_hr_page_break", v)
        )
        
        # 第二列 - 分节符
        self.docx_hr_section_break_combo = self._create_config_combobox(
            second_column,
            t("settings.formatting.section_break_label"),
            DOCX_TO_MD_HR_CONFIG_VALUES,
            DOCX_TO_MD_HR_TRANSLATE_KEYS,
            hr_config.get("section_break", "***"),
            t("settings.formatting.section_break_tooltip"),
            lambda v: self.on_change("docx_hr_section_break", v)
        )
        
        # 第三列 - 分隔线
        self.docx_hr_horizontal_rule_combo = self._create_config_combobox(
            third_column,
            t("settings.formatting.horizontal_rule_label"),
            DOCX_TO_MD_HR_CONFIG_VALUES,
            DOCX_TO_MD_HR_TRANSLATE_KEYS,
            hr_config.get("horizontal_rule", "___"),
            t("settings.formatting.horizontal_rule_tooltip"),
            lambda v: self.on_change("docx_hr_horizontal_rule", v)
        )
    
    def _create_syntax_config(self, parent_frame):
        """
        创建Markdown语法选择子区块（两列布局）
        
        配置路径：conversion_config.syntax.*
        
        作为"文档转为MarkDown" section的子区块
        """
        from docwen.utils.gui_utils import create_info_icon
        
        # 分隔线
        separator = tb.Separator(parent_frame, orient="horizontal")
        separator.pack(fill="x", pady=(15, 10))
        
        # 小节标题
        title_frame = tb.Frame(parent_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            title_frame,
            text=t("settings.formatting.syntax_selection"),
            font=(self.small_font, self.small_size, "bold")
        )
        title_label.pack(side="left")
        
        info = create_info_icon(
            title_frame,
            t("settings.formatting.syntax_selection_tooltip"),
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 获取当前配置
        syntax_config = self.config_manager.get_all_syntax_settings()
        
        # 创建两列容器
        columns_frame = tb.Frame(parent_frame)
        columns_frame.pack(fill="x")
        
        # 配置列权重使两列等宽
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        
        # 左列
        left_column = tb.Frame(columns_frame)
        left_column.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        # 右列
        right_column = tb.Frame(columns_frame)
        right_column.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        
        # ========== 左列设置项 ==========
        
        # 粗体语法
        self.bold_combo = self._create_config_combobox(
            left_column,
            t("settings.formatting.bold_syntax_label"),
            BOLD_ITALIC_CONFIG_VALUES,
            BOLD_ITALIC_TRANSLATE_KEYS,
            syntax_config.get("bold", "asterisk"),
            t("settings.formatting.bold_syntax_tooltip"),
            lambda v: self.on_change("bold", v)
        )
        
        # 删除线语法
        self.strikethrough_combo = self._create_config_combobox(
            left_column,
            t("settings.formatting.strikethrough_syntax_label"),
            EXTENDED_HTML_CONFIG_VALUES,
            STRIKETHROUGH_TRANSLATE_KEYS,
            syntax_config.get("strikethrough", "extended"),
            t("settings.formatting.strikethrough_syntax_tooltip"),
            lambda v: self.on_change("strikethrough", v)
        )
        
        # 上标语法
        self.superscript_combo = self._create_config_combobox(
            left_column,
            t("settings.formatting.superscript_syntax_label"),
            EXTENDED_HTML_CONFIG_VALUES,
            SUPERSCRIPT_TRANSLATE_KEYS,
            syntax_config.get("superscript", "html"),
            t("settings.formatting.superscript_syntax_tooltip"),
            lambda v: self.on_change("superscript", v)
        )
        
        # 无序列表标记
        self.unordered_list_combo = self._create_config_combobox(
            left_column,
            t("settings.formatting.unordered_list_label"),
            UNORDERED_LIST_CONFIG_VALUES,
            UNORDERED_LIST_TRANSLATE_KEYS,
            self.config_manager.get_unordered_list_marker(),
            t("settings.formatting.unordered_list_tooltip"),
            lambda v: self.on_change("unordered_list", v)
        )
        
        # 列表缩进空格数
        indent_spaces_value = self.config_manager.get_list_indent_spaces()
        self.indent_spaces_combo = self._create_config_combobox(
            left_column,
            t("settings.formatting.indent_spaces_label"),
            INDENT_SPACES_CONFIG_VALUES,
            INDENT_SPACES_TRANSLATE_KEYS,
            str(indent_spaces_value),
            t("settings.formatting.indent_spaces_tooltip"),
            lambda v: self.on_change("indent_spaces", int(v))
        )
        
        # ========== 右列设置项 ==========
        
        # 斜体语法
        self.italic_combo = self._create_config_combobox(
            right_column,
            t("settings.formatting.italic_syntax_label"),
            BOLD_ITALIC_CONFIG_VALUES,
            ITALIC_TRANSLATE_KEYS,
            syntax_config.get("italic", "asterisk"),
            t("settings.formatting.italic_syntax_tooltip"),
            lambda v: self.on_change("italic", v)
        )
        
        # 高亮语法
        self.highlight_combo = self._create_config_combobox(
            right_column,
            t("settings.formatting.highlight_syntax_label"),
            EXTENDED_HTML_CONFIG_VALUES,
            HIGHLIGHT_TRANSLATE_KEYS,
            syntax_config.get("highlight", "extended"),
            t("settings.formatting.highlight_syntax_tooltip"),
            lambda v: self.on_change("highlight", v)
        )
        
        # 下标语法
        self.subscript_combo = self._create_config_combobox(
            right_column,
            t("settings.formatting.subscript_syntax_label"),
            EXTENDED_HTML_CONFIG_VALUES,
            SUBSCRIPT_TRANSLATE_KEYS,
            syntax_config.get("subscript", "html"),
            t("settings.formatting.subscript_syntax_tooltip"),
            lambda v: self.on_change("subscript", v)
        )
        
        # 有序列表编号（禁用状态）
        self.ordered_list_combo = self._create_config_combobox(
            right_column,
            t("settings.formatting.ordered_list_label"),
            ORDERED_LIST_CONFIG_VALUES,
            ORDERED_LIST_TRANSLATE_KEYS,
            self.config_manager.get_ordered_list_mode(),
            t("settings.formatting.ordered_list_tooltip"),
            lambda v: self.on_change("ordered_list", v),
            disabled=True
        )
        
        logger.debug("Markdown语法选择子区块创建完成")
    
    def _create_md_to_docx_section(self):
        """
        创建MD → DOCX 格式设置区域（三列布局）
        
        配置路径：conversion_config.md_to_docx.*
        """
        logger.debug("创建MD → DOCX 格式设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            t("settings.formatting.md_to_docx_section"),
            SectionStyle.PRIMARY
        )
        
        # 获取当前配置
        formatting_mode = self.config_manager.get_md_to_docx_formatting_mode()
        heading_formatting_mode = self.config_manager.get_md_to_docx_heading_formatting_mode()
        table_header_formatting_mode = self.config_manager.get_md_to_docx_table_header_formatting_mode()
        
        # 格式处理小标题
        from docwen.utils.gui_utils import create_info_icon
        
        format_title_frame = tb.Frame(frame)
        format_title_frame.pack(fill="x", pady=(0, 10))
        
        format_title_label = tb.Label(
            format_title_frame,
            text=t("settings.formatting.md_format_processing"),
            font=(self.small_font, self.small_size, "bold")
        )
        format_title_label.pack(side="left")
        
        format_info = create_info_icon(
            format_title_frame,
            t("settings.formatting.md_format_processing_tooltip"),
            "info"
        )
        format_info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 创建三列容器
        columns_frame = tb.Frame(frame)
        columns_frame.pack(fill="x")
        
        # 配置列权重使三列等宽
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        columns_frame.columnconfigure(2, weight=1)
        
        # 第一列
        first_column = tb.Frame(columns_frame)
        first_column.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        # 第二列
        second_column = tb.Frame(columns_frame)
        second_column.grid(row=0, column=1, sticky="nsew", padx=(5, 5))
        
        # 第三列
        third_column = tb.Frame(columns_frame)
        third_column.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        
        # 第一列 - 正文格式
        self.formatting_mode_combo = self._create_config_combobox(
            first_column,
            t("settings.formatting.md_body_format_label"),
            FORMATTING_MODE_CONFIG_VALUES,
            FORMATTING_MODE_TRANSLATE_KEYS,
            formatting_mode,
            t("settings.formatting.md_body_format_tooltip"),
            lambda v: self.on_change("formatting_mode", v)
        )
        
        # 第二列 - 小标题格式
        self.heading_formatting_mode_combo = self._create_config_combobox(
            second_column,
            t("settings.formatting.md_heading_format_label"),
            FORMATTING_MODE_CONFIG_VALUES,
            FORMATTING_MODE_TRANSLATE_KEYS,
            heading_formatting_mode,
            t("settings.formatting.md_heading_format_tooltip"),
            lambda v: self.on_change("heading_formatting_mode", v)
        )
        
        # 第三列 - 表头格式
        self.table_header_formatting_mode_combo = self._create_config_combobox(
            third_column,
            t("settings.formatting.md_table_header_format_label"),
            FORMATTING_MODE_CONFIG_VALUES,
            FORMATTING_MODE_TRANSLATE_KEYS,
            table_header_formatting_mode,
            t("settings.formatting.md_table_header_format_tooltip"),
            lambda v: self.on_change("table_header_formatting_mode", v)
        )
        
        # ========== YAML列表拼接符配置 ==========
        self._create_list_separator_config(frame)
        
        # ========== 分隔符映射配置（MD转文档）==========
        self._create_md_to_docx_hr_config(frame)
        
        # ========== 表格样式配置 ==========
        self._create_table_style_config(frame)
        
        logger.debug("MD → DOCX 格式设置区域创建完成")
    
    def _create_list_separator_config(self, parent_frame):
        """
        创建YAML列表拼接符配置区域
        
        配置路径：conversion_config.md_to_docx.list_separator
        
        当YAML字段值为列表时，使用此分隔符拼接为字符串。
        常见值：顿号（、）、逗号（，）、逗号空格（, ）、空字符串
        """
        from docwen.utils.gui_utils import create_info_icon
        
        # 分隔线
        separator = tb.Separator(parent_frame, orient="horizontal")
        separator.pack(fill="x", pady=(15, 10))
        
        # 小节标题
        title_frame = tb.Frame(parent_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            title_frame,
            text=t("settings.formatting.list_separator_label"),
            font=(self.small_font, self.small_size, "bold")
        )
        title_label.pack(side="left")
        
        info = create_info_icon(
            title_frame,
            t("settings.formatting.list_separator_tooltip"),
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 获取当前配置
        current_separator = self.config_manager.get_yaml_list_separator()
        
        # 创建输入框容器
        input_frame = tb.Frame(parent_frame)
        input_frame.pack(fill="x")
        
        # 输入框变量
        self.list_separator_var = tk.StringVar(value=current_separator)
        
        # 创建输入框
        self.list_separator_entry = tb.Entry(
            input_frame,
            textvariable=self.list_separator_var,
            width=20,
            bootstyle="secondary"
        )
        self.list_separator_entry.pack(side="left")
        self.list_separator_entry.bind("<FocusOut>", self._on_list_separator_changed)
        self.list_separator_entry.bind("<Return>", self._on_list_separator_changed)
        
        logger.debug("YAML列表拼接符配置区域创建完成")
    
    def _on_list_separator_changed(self, event=None):
        """处理列表拼接符变更"""
        separator = self.list_separator_var.get()
        logger.info(f"YAML列表拼接符变更: '{separator}'")
        self.on_change("list_separator", separator)
    
    def _create_md_to_docx_hr_config(self, parent_frame):
        """
        创建MD转文档的分隔符映射配置
        
        配置路径：conversion_config.horizontal_rule.md_to_docx.*
        """
        from docwen.utils.gui_utils import create_info_icon
        
        # 分隔线
        separator = tb.Separator(parent_frame, orient="horizontal")
        separator.pack(fill="x", pady=(15, 10))
        
        # 小节标题
        title_frame = tb.Frame(parent_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            title_frame,
            text=t("settings.formatting.md_separator_mapping"),
            font=(self.small_font, self.small_size, "bold")
        )
        title_label.pack(side="left")
        
        info = create_info_icon(
            title_frame,
            t("settings.formatting.md_separator_mapping_tooltip"),
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 获取当前配置
        hr_config = self.config_manager.get_horizontal_rule_md_to_docx_config()
        
        # 创建三列容器
        columns_frame = tb.Frame(parent_frame)
        columns_frame.pack(fill="x")
        
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        columns_frame.columnconfigure(2, weight=1)
        
        # 第一列 - ---
        first_column = tb.Frame(columns_frame)
        first_column.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        # 第二列 - ***
        second_column = tb.Frame(columns_frame)
        second_column.grid(row=0, column=1, sticky="nsew", padx=(5, 5))
        
        # 第三列 - ___
        third_column = tb.Frame(columns_frame)
        third_column.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        
        # 第一列 - ---
        self.md_hr_dash_combo = self._create_config_combobox(
            first_column,
            t("settings.formatting.dash_label"),
            MD_TO_DOCX_HR_CONFIG_VALUES,
            MD_TO_DOCX_HR_TRANSLATE_KEYS,
            hr_config.get("dash", "page_break"),
            t("settings.formatting.dash_tooltip"),
            lambda v: self.on_change("md_hr_dash", v)
        )
        
        # 第二列 - ***
        self.md_hr_asterisk_combo = self._create_config_combobox(
            second_column,
            t("settings.formatting.asterisk_label"),
            MD_TO_DOCX_HR_CONFIG_VALUES,
            MD_TO_DOCX_HR_TRANSLATE_KEYS,
            hr_config.get("asterisk", "section_continuous"),
            t("settings.formatting.asterisk_tooltip"),
            lambda v: self.on_change("md_hr_asterisk", v)
        )
        
        # 第三列 - ___
        self.md_hr_underscore_combo = self._create_config_combobox(
            third_column,
            t("settings.formatting.underscore_label"),
            MD_TO_DOCX_HR_CONFIG_VALUES,
            MD_TO_DOCX_HR_TRANSLATE_KEYS,
            hr_config.get("underscore", "section_break"),
            t("settings.formatting.underscore_tooltip"),
            lambda v: self.on_change("md_hr_underscore", v)
        )
    
    def _create_table_style_config(self, parent_frame):
        """
        创建表格样式配置区域
        
        配置路径：style_table.md_to_docx.*
        
        支持两种模式：
        - 内置样式：三线表 / 网格表
        - 自定义样式：用户输入样式名称
        """
        from docwen.utils.gui_utils import create_info_icon
        
        # 分隔线
        separator = tb.Separator(parent_frame, orient="horizontal")
        separator.pack(fill="x", pady=(15, 10))
        
        # 小节标题
        title_frame = tb.Frame(parent_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            title_frame,
            text=t("settings.formatting.table_style"),
            font=(self.small_font, self.small_size, "bold")
        )
        title_label.pack(side="left")
        
        info = create_info_icon(
            title_frame,
            t("settings.formatting.table_style_tooltip"),
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 获取当前配置
        style_mode = self.config_manager.get_table_style_mode()
        builtin_key = self.config_manager.get_builtin_table_style_key()
        custom_name = self.config_manager.get_custom_table_style_name()
        
        # 创建两列容器
        columns_frame = tb.Frame(parent_frame)
        columns_frame.pack(fill="x")
        
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        
        # 第一列 - 内置样式
        first_column = tb.Frame(columns_frame)
        first_column.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        # 第二列 - 自定义样式
        second_column = tb.Frame(columns_frame)
        second_column.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        
        # 创建单选按钮变量
        self.table_style_mode_var = tk.StringVar(value=style_mode)
        
        # === 第一列：内置样式 ===
        builtin_radio = tb.Radiobutton(
            first_column,
            text=t("settings.formatting.builtin_style_radio"),
            variable=self.table_style_mode_var,
            value="builtin",
            command=self._on_table_style_mode_changed
        )
        builtin_radio.pack(anchor="w", pady=(0, 5))
        
        # 内置样式下拉框（使用 ConfigCombobox）
        self.builtin_style_combo = ConfigCombobox(
            first_column,
            config_values=BUILTIN_TABLE_STYLE_CONFIG_VALUES,
            translate_keys=BUILTIN_TABLE_STYLE_TRANSLATE_KEYS,
            initial_value=builtin_key,
            on_change=lambda v: self.on_change("builtin_style_key", v),
            state="readonly" if style_mode == "builtin" else "disabled"
        )
        self.builtin_style_combo.pack(fill="x", pady=(0, 5))
        
        # === 第二列：自定义样式 ===
        custom_radio = tb.Radiobutton(
            second_column,
            text=t("settings.formatting.custom_style_radio"),
            variable=self.table_style_mode_var,
            value="custom",
            command=self._on_table_style_mode_changed
        )
        custom_radio.pack(anchor="w", pady=(0, 5))
        
        # 自定义样式输入框
        self.custom_style_var = tk.StringVar(value=custom_name)
        
        self.custom_style_entry = tb.Entry(
            second_column,
            textvariable=self.custom_style_var,
            state="normal" if style_mode == "custom" else "disabled",
            bootstyle="secondary"
        )
        self.custom_style_entry.pack(fill="x", pady=(0, 5))
        self.custom_style_entry.bind("<FocusOut>", self._on_custom_style_changed)
        self.custom_style_entry.bind("<Return>", self._on_custom_style_changed)
        
        logger.debug("表格样式配置区域创建完成")
    
    def _on_table_style_mode_changed(self):
        """处理表格样式模式变更（单选按钮）"""
        mode = self.table_style_mode_var.get()
        logger.info(f"表格样式模式变更: {mode}")
        
        # 更新控件状态
        if mode == "builtin":
            self.builtin_style_combo.configure(state="readonly")
            self.custom_style_entry.configure(state="disabled")
        else:
            self.builtin_style_combo.configure(state="disabled")
            self.custom_style_entry.configure(state="normal")
        
        self.on_change("table_style_mode", mode)
    
    def _on_custom_style_changed(self, event=None):
        """处理自定义样式名称变更"""
        custom_name = self.custom_style_var.get().strip()
        logger.info(f"自定义表格样式变更: {custom_name}")
        self.on_change("custom_style_name", custom_name)
    
    # --------------------------
    # 抽象方法实现
    # --------------------------
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前设置（实现抽象方法）
        
        返回:
            Dict: 配置字典
        """
        # 获取缩进空格数配置值并转换为整数
        indent_spaces_str = self.indent_spaces_combo.get_config_value()
        indent_spaces = int(indent_spaces_str) if indent_spaces_str else 2
        
        return {
            # DOCX → MD 格式保留（布尔值）
            "preserve_formatting": self.preserve_formatting_combo.get_config_value() == "true",
            "preserve_heading_formatting": self.preserve_heading_formatting_combo.get_config_value() == "true",
            "preserve_table_header_formatting": self.preserve_table_header_formatting_combo.get_config_value() == "true",
            # MD → DOCX 格式处理模式
            "formatting_mode": self.formatting_mode_combo.get_config_value(),
            "heading_formatting_mode": self.heading_formatting_mode_combo.get_config_value(),
            "table_header_formatting_mode": self.table_header_formatting_mode_combo.get_config_value(),
            # YAML列表拼接符
            "list_separator": self.list_separator_var.get(),
            # Markdown 语法设置
            "bold": self.bold_combo.get_config_value(),
            "italic": self.italic_combo.get_config_value(),
            "strikethrough": self.strikethrough_combo.get_config_value(),
            "highlight": self.highlight_combo.get_config_value(),
            "superscript": self.superscript_combo.get_config_value(),
            "subscript": self.subscript_combo.get_config_value(),
            "unordered_list": self.unordered_list_combo.get_config_value(),
            "ordered_list": self.ordered_list_combo.get_config_value(),
            "indent_spaces": indent_spaces,
            # 分隔符转换配置 - 文档转MD
            "docx_hr_page_break": self.docx_hr_page_break_combo.get_config_value(),
            "docx_hr_section_break": self.docx_hr_section_break_combo.get_config_value(),
            "docx_hr_horizontal_rule": self.docx_hr_horizontal_rule_combo.get_config_value(),
            # 分隔符转换配置 - MD转文档
            "md_hr_dash": self.md_hr_dash_combo.get_config_value(),
            "md_hr_asterisk": self.md_hr_asterisk_combo.get_config_value(),
            "md_hr_underscore": self.md_hr_underscore_combo.get_config_value(),
            # 表格样式配置
            "table_style_mode": self.table_style_mode_var.get(),
            "builtin_style_key": self.builtin_style_combo.get_config_value(),
            "custom_style_name": self.custom_style_var.get().strip()
        }
    
    def apply_settings(self) -> bool:
        """
        应用设置到配置文件（实现抽象方法）
        
        返回:
            bool: 保存是否成功
        """
        settings = self.get_settings()
        logger.info(f"保存格式配置: {settings}")
        
        try:
            # === DOCX → MD 配置 ===
            if not self.config_manager.update_config_value(
                "conversion_config", "docx_to_md", "preserve_formatting", 
                settings["preserve_formatting"]
            ):
                logger.error("保存 preserve_formatting 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "conversion_config", "docx_to_md", "preserve_heading_formatting",
                settings["preserve_heading_formatting"]
            ):
                logger.error("保存 preserve_heading_formatting 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "conversion_config", "docx_to_md", "preserve_table_header_formatting",
                settings["preserve_table_header_formatting"]
            ):
                logger.error("保存 preserve_table_header_formatting 失败")
                return False
            
            # === MD → DOCX 配置 ===
            if not self.config_manager.update_config_value(
                "conversion_config", "md_to_docx", "formatting_mode",
                settings["formatting_mode"]
            ):
                logger.error("保存 formatting_mode 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "conversion_config", "md_to_docx", "heading_formatting_mode",
                settings["heading_formatting_mode"]
            ):
                logger.error("保存 heading_formatting_mode 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "conversion_config", "md_to_docx", "table_header_formatting_mode",
                settings["table_header_formatting_mode"]
            ):
                logger.error("保存 table_header_formatting_mode 失败")
                return False
            
            # === YAML列表拼接符 ===
            if not self.config_manager.update_config_value(
                "conversion_config", "md_to_docx", "list_separator",
                settings["list_separator"]
            ):
                logger.error("保存 list_separator 失败")
                return False
            
            # === 语法配置 ===
            syntax_keys = ["bold", "italic", "strikethrough", "highlight", "superscript", "subscript", "unordered_list", "indent_spaces"]
            for key in syntax_keys:
                if not self.config_manager.update_config_value(
                    "conversion_config", "syntax", key, settings[key]
                ):
                    logger.error(f"保存 {key} 语法配置失败")
                    return False
            
            # ordered_list 暂不保存（功能未实现）
            
            # === 分隔符转换配置 - 文档转MD ===
            if not self.config_manager.update_config_value(
                "conversion_config", "horizontal_rule.docx_to_md", "page_break",
                settings["docx_hr_page_break"]
            ):
                logger.error("保存 docx_hr_page_break 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "conversion_config", "horizontal_rule.docx_to_md", "section_break",
                settings["docx_hr_section_break"]
            ):
                logger.error("保存 docx_hr_section_break 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "conversion_config", "horizontal_rule.docx_to_md", "horizontal_rule",
                settings["docx_hr_horizontal_rule"]
            ):
                logger.error("保存 docx_hr_horizontal_rule 失败")
                return False
            
            # === 分隔符转换配置 - MD转文档 ===
            if not self.config_manager.update_config_value(
                "conversion_config", "horizontal_rule.md_to_docx", "dash",
                settings["md_hr_dash"]
            ):
                logger.error("保存 md_hr_dash 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "conversion_config", "horizontal_rule.md_to_docx", "asterisk",
                settings["md_hr_asterisk"]
            ):
                logger.error("保存 md_hr_asterisk 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "conversion_config", "horizontal_rule.md_to_docx", "underscore",
                settings["md_hr_underscore"]
            ):
                logger.error("保存 md_hr_underscore 失败")
                return False
            
            # === 表格样式配置 ===
            if not self.config_manager.update_config_value(
                "style_table", "md_to_docx", "table_style_mode",
                settings["table_style_mode"]
            ):
                logger.error("保存 table_style_mode 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "style_table", "md_to_docx", "builtin_style_key",
                settings["builtin_style_key"]
            ):
                logger.error("保存 builtin_style_key 失败")
                return False
            
            if not self.config_manager.update_config_value(
                "style_table", "md_to_docx", "custom_style_name",
                settings["custom_style_name"]
            ):
                logger.error("保存 custom_style_name 失败")
                return False
            
            logger.info("格式配置保存成功")
            return True
            
        except Exception as e:
            logger.error(f"保存格式配置失败: {e}")
            return False
