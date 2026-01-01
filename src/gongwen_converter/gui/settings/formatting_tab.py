"""
格式设置选项卡

实现设置对话框的格式设置选项卡，包含：
- DOCX → MD 格式设置
- MD → DOCX 格式设置
- Markdown语法选择
"""

import tkinter as tk
import ttkbootstrap as tb
from typing import Dict, Any
import logging

from .base_tab import BaseSettingsTab, SectionStyle

logger = logging.getLogger(__name__)


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
    
    def _create_docx_to_md_section(self):
        """
        创建DOCX → MD 格式设置区域（三列布局）
        
        配置路径：conversion_config.docx_to_md.*
        """
        logger.debug("创建DOCX → MD 格式设置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "文档转为MarkDown",
            SectionStyle.PRIMARY
        )
        
        # 获取当前配置
        preserve_formatting = self.config_manager.get_docx_to_md_preserve_formatting()
        preserve_heading_formatting = self.config_manager.get_docx_to_md_preserve_heading_formatting()
        preserve_table_header_formatting = self.config_manager.get_docx_to_md_preserve_table_header_formatting()
        
        # 定义格式处理选项的映射（简化文本）
        self._preserve_formatting_mapping = {
            True: "保留格式",
            False: "丢弃格式"
        }
        
        # 定义小标题格式处理选项的映射（简化文本）
        self._preserve_heading_formatting_mapping = {
            True: "保留格式",
            False: "丢弃格式"
        }
        
        # 定义表头格式处理选项的映射（简化文本）
        self._preserve_table_header_formatting_mapping = {
            True: "保留格式",
            False: "丢弃格式"
        }
        
        # 根据当前配置获取显示值
        current_display = self._preserve_formatting_mapping.get(preserve_formatting, "保留格式")
        self.preserve_formatting_var = tk.StringVar(value=current_display)
        
        current_heading_display = self._preserve_heading_formatting_mapping.get(preserve_heading_formatting, "丢弃格式")
        self.preserve_heading_formatting_var = tk.StringVar(value=current_heading_display)
        
        current_table_header_display = self._preserve_table_header_formatting_mapping.get(preserve_table_header_formatting, "丢弃格式")
        self.preserve_table_header_formatting_var = tk.StringVar(value=current_table_header_display)
        
        # 格式处理小标题
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        format_title_frame = tb.Frame(frame)
        format_title_frame.pack(fill="x", pady=(0, 10))
        
        format_title_label = tb.Label(
            format_title_frame,
            text="格式处理",
            font=(self.small_font, self.small_size, "bold")
        )
        format_title_label.pack(side="left")
        
        format_info = create_info_icon(
            format_title_frame,
            '配置 Word 文档转换为 Markdown 时的格式处理方式\n\n部分选项对“针对公文优化”的文档转MarkDown不起效',
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
        self._create_column_combobox(
            first_column,
            "正文格式处理:",
            self.preserve_formatting_var,
            list(self._preserve_formatting_mapping.values()),
            "保留格式：粗体→**粗体**、斜体→*斜体*、删除线→~~删除线~~等\n丢弃格式：只保留纯文本内容",
            self._on_preserve_formatting_changed
        )
        
        # 第二列 - 小标题格式
        self._create_column_combobox(
            second_column,
            "小标题格式处理:",
            self.preserve_heading_formatting_var,
            list(self._preserve_heading_formatting_mapping.values()),
            "保留格式：标题内的部分加粗转为 **加粗** 等\n丢弃格式：标题只保留纯文本，避免冗余标记\n注意：通常Word标题样式已有固定格式，保留会导致冗余标记",
            self._on_preserve_heading_formatting_changed
        )
        
        # 第三列 - 表头格式
        self._create_column_combobox(
            third_column,
            "表头格式处理:",
            self.preserve_table_header_formatting_var,
            list(self._preserve_table_header_formatting_mapping.values()),
            "保留格式：表头单元格中的粗体转为 **粗体** 等\n丢弃格式：表头只保留纯文本，避免冗余标记\n注意：表格样式通常已定义表头加粗，保留会导致重复标记",
            self._on_preserve_table_header_formatting_changed
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
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 分隔线
        separator = tb.Separator(parent_frame, orient="horizontal")
        separator.pack(fill="x", pady=(15, 10))
        
        # 小节标题
        title_frame = tb.Frame(parent_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            title_frame,
            text="分隔符映射",
            font=(self.small_font, self.small_size, "bold")
        )
        title_label.pack(side="left")
        
        info = create_info_icon(
            title_frame,
            "配置 Word 分页符/分节符/分隔线转换为哪种 Markdown 分隔符\n\n"
            "分页符：Word 中的分页符\n"
            "分节符：所有类型的分节符（下一页/连续/奇数页/偶数页）\n"
            "分隔线：段落底部边框（用户输入 --- 后回车产生）",
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 定义映射关系（MD分隔符选项）
        self._docx_to_md_hr_mapping = {
            "---": "---",
            "***": "***",
            "___": "___",
            "ignore": "忽略"
        }
        
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
        
        # 获取当前值的显示文本
        page_break_value = hr_config.get("page_break", "---")
        section_break_value = hr_config.get("section_break", "***")
        horizontal_rule_value = hr_config.get("horizontal_rule", "___")
        
        page_break_display = self._docx_to_md_hr_mapping.get(page_break_value, "---")
        section_break_display = self._docx_to_md_hr_mapping.get(section_break_value, "***")
        horizontal_rule_display = self._docx_to_md_hr_mapping.get(horizontal_rule_value, "___")
        
        self.docx_hr_page_break_var = tk.StringVar(value=page_break_display)
        self.docx_hr_section_break_var = tk.StringVar(value=section_break_display)
        self.docx_hr_horizontal_rule_var = tk.StringVar(value=horizontal_rule_display)
        
        # 第一列 - 分页符
        self._create_column_combobox(
            first_column,
            "分页符:",
            self.docx_hr_page_break_var,
            list(self._docx_to_md_hr_mapping.values()),
            "Word 分页符转换为 Markdown 分隔符\n分页符通常用于在特定位置强制换页",
            self._on_docx_hr_page_break_changed
        )
        
        # 第二列 - 分节符
        self._create_column_combobox(
            second_column,
            "分节符:",
            self.docx_hr_section_break_var,
            list(self._docx_to_md_hr_mapping.values()),
            "Word 分节符转换为 Markdown 分隔符\n所有分节符类型（下一页/连续/奇偶页）统一处理",
            self._on_docx_hr_section_break_changed
        )
        
        # 第三列 - 分隔线
        self._create_column_combobox(
            third_column,
            "分隔线:",
            self.docx_hr_horizontal_rule_var,
            list(self._docx_to_md_hr_mapping.values()),
            "Word 分隔线（Horizontal Rule 1/2/3 样式）转换为 Markdown 分隔符",
            self._on_docx_hr_horizontal_rule_changed
        )
    
    def _create_syntax_config(self, parent_frame):
        """
        创建Markdown语法选择子区块（两列布局）
        
        配置路径：conversion_config.syntax.*
        
        作为"文档转为MarkDown" section的子区块
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 分隔线
        separator = tb.Separator(parent_frame, orient="horizontal")
        separator.pack(fill="x", pady=(15, 10))
        
        # 小节标题
        title_frame = tb.Frame(parent_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            title_frame,
            text="语法选择",
            font=(self.small_font, self.small_size, "bold")
        )
        title_label.pack(side="left")
        
        info = create_info_icon(
            title_frame,
            "选择文档转为Markdown时输出的语法格式\n（输入时会自动识别所有格式语法）",
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 获取当前配置
        syntax_config = self.config_manager.get_all_syntax_settings()
        
        # 定义语法选项的映射关系（显示文本 <-> 配置值）
        self._syntax_mappings = {
            "bold": {
                "asterisk": "星号 (**文本**)",
                "underscore": "下划线 (__文本__)"
            },
            "italic": {
                "asterisk": "星号 (*文本*)",
                "underscore": "下划线 (_文本_)"
            },
            "strikethrough": {
                "extended": "扩展 (~~文本~~)",
                "html": "HTML (<del>)"
            },
            "highlight": {
                "extended": "扩展 (==文本==)",
                "html": "HTML (<mark>)"
            },
            "superscript": {
                "extended": "扩展 (^文本^)",
                "html": "HTML (<sup>)"
            },
            "subscript": {
                "extended": "扩展 (~文本~)",
                "html": "HTML (<sub>)"
            },
            "unordered_list": {
                "dash": "短横线 (- 内容)",
                "asterisk": "星号 (* 内容)",
                "plus": "加号 (+ 内容)"
            },
            "ordered_list": {
                "restart": "从1开始",
                "preserve": "保留原编号"
            },
            "indent_spaces": {
                2: "2个空格 (推荐)",
                4: "4个空格"
            }
        }
        
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
        bold_value = syntax_config.get("bold", "asterisk")
        bold_display = self._syntax_mappings["bold"].get(bold_value, "星号 (**文本**)")
        self.bold_var = tk.StringVar(value=bold_display)
        self._create_column_combobox(
            left_column,
            "粗体语法:",
            self.bold_var,
            list(self._syntax_mappings["bold"].values()),
            "选择粗体的输出语法格式",
            self._on_bold_syntax_changed
        )
        
        # 删除线语法
        strikethrough_value = syntax_config.get("strikethrough", "extended")
        strikethrough_display = self._syntax_mappings["strikethrough"].get(strikethrough_value, "扩展 (~~文本~~)")
        self.strikethrough_var = tk.StringVar(value=strikethrough_display)
        self._create_column_combobox(
            left_column,
            "删除线语法:",
            self.strikethrough_var,
            list(self._syntax_mappings["strikethrough"].values()),
            "选择删除线的输出语法格式",
            self._on_strikethrough_syntax_changed
        )
        
        # 上标语法
        superscript_value = syntax_config.get("superscript", "html")
        superscript_display = self._syntax_mappings["superscript"].get(superscript_value, "HTML (<sup>)")
        self.superscript_var = tk.StringVar(value=superscript_display)
        self._create_column_combobox(
            left_column,
            "上标语法:",
            self.superscript_var,
            list(self._syntax_mappings["superscript"].values()),
            "选择上标的输出语法格式",
            self._on_superscript_syntax_changed
        )
        
        # 无序列表标记
        unordered_list_value = self.config_manager.get_unordered_list_marker()
        unordered_list_display = self._syntax_mappings["unordered_list"].get(unordered_list_value, "短横线 (- 内容)")
        self.unordered_list_var = tk.StringVar(value=unordered_list_display)
        self._create_column_combobox(
            left_column,
            "无序列表:",
            self.unordered_list_var,
            list(self._syntax_mappings["unordered_list"].values()),
            "选择无序列表的输出标记符号",
            self._on_unordered_list_changed
        )
        
        # 列表缩进空格数
        indent_spaces_value = self.config_manager.get_list_indent_spaces()
        indent_spaces_display = self._syntax_mappings["indent_spaces"].get(indent_spaces_value, "2个空格 (推荐)")
        self.indent_spaces_var = tk.StringVar(value=indent_spaces_display)
        self._create_column_combobox(
            left_column,
            "列表缩进:",
            self.indent_spaces_var,
            list(self._syntax_mappings["indent_spaces"].values()),
            "Markdown列表每级缩进的空格数，用于识别多级嵌套\n2个空格：Obsidian、Typora 等常用\n4个空格：VS Code 等常用",
            self._on_indent_spaces_changed
        )
        
        # ========== 右列设置项 ==========
        
        # 斜体语法
        italic_value = syntax_config.get("italic", "asterisk")
        italic_display = self._syntax_mappings["italic"].get(italic_value, "星号 (*文本*)")
        self.italic_var = tk.StringVar(value=italic_display)
        self._create_column_combobox(
            right_column,
            "斜体语法:",
            self.italic_var,
            list(self._syntax_mappings["italic"].values()),
            "选择斜体的输出语法格式",
            self._on_italic_syntax_changed
        )
        
        # 高亮语法
        highlight_value = syntax_config.get("highlight", "extended")
        highlight_display = self._syntax_mappings["highlight"].get(highlight_value, "扩展 (==文本==)")
        self.highlight_var = tk.StringVar(value=highlight_display)
        self._create_column_combobox(
            right_column,
            "高亮语法:",
            self.highlight_var,
            list(self._syntax_mappings["highlight"].values()),
            "选择高亮的输出语法格式",
            self._on_highlight_syntax_changed
        )
        
        # 下标语法
        subscript_value = syntax_config.get("subscript", "html")
        subscript_display = self._syntax_mappings["subscript"].get(subscript_value, "HTML (<sub>)")
        self.subscript_var = tk.StringVar(value=subscript_display)
        self._create_column_combobox(
            right_column,
            "下标语法:",
            self.subscript_var,
            list(self._syntax_mappings["subscript"].values()),
            "选择下标的输出语法格式",
            self._on_subscript_syntax_changed
        )
        
        # 有序列表编号（禁用状态）
        ordered_list_value = self.config_manager.get_ordered_list_mode()
        ordered_list_display = self._syntax_mappings["ordered_list"].get(ordered_list_value, "从1开始")
        self.ordered_list_var = tk.StringVar(value=ordered_list_display)
        self._ordered_list_combobox = self._create_column_combobox(
            right_column,
            "有序列表:",
            self.ordered_list_var,
            list(self._syntax_mappings["ordered_list"].values()),
            "有序列表编号方式（此选项暂未实现，将在后续版本支持）",
            self._on_ordered_list_changed,
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
            "MarkDown转为文档",
            SectionStyle.PRIMARY
        )
        
        # 获取当前配置
        formatting_mode = self.config_manager.get_md_to_docx_formatting_mode()
        heading_formatting_mode = self.config_manager.get_md_to_docx_heading_formatting_mode()
        table_header_formatting_mode = self.config_manager.get_md_to_docx_table_header_formatting_mode()
        
        # 定义格式处理模式的映射（简化文本）
        self._formatting_mode_mapping = {
            "apply": "应用格式",
            "keep": "保留标记",
            "remove": "清理标记"
        }
        
        # 定义小标题格式处理模式的映射（简化文本）
        self._heading_formatting_mode_mapping = {
            "apply": "应用格式",
            "keep": "保留标记",
            "remove": "清理标记"
        }
        
        # 定义表头格式处理模式的映射（简化文本）
        self._table_header_formatting_mode_mapping = {
            "apply": "应用格式",
            "keep": "保留标记",
            "remove": "清理标记"
        }
        
        # 根据当前配置获取显示值
        current_display = self._formatting_mode_mapping.get(formatting_mode, "应用格式")
        self.formatting_mode_var = tk.StringVar(value=current_display)
        
        current_heading_display = self._heading_formatting_mode_mapping.get(heading_formatting_mode, "清理标记")
        self.heading_formatting_mode_var = tk.StringVar(value=current_heading_display)
        
        current_table_header_display = self._table_header_formatting_mode_mapping.get(table_header_formatting_mode, "清理标记")
        self.table_header_formatting_mode_var = tk.StringVar(value=current_table_header_display)
        
        # 格式处理小标题
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        format_title_frame = tb.Frame(frame)
        format_title_frame.pack(fill="x", pady=(0, 10))
        
        format_title_label = tb.Label(
            format_title_frame,
            text="格式处理",
            font=(self.small_font, self.small_size, "bold")
        )
        format_title_label.pack(side="left")
        
        format_info = create_info_icon(
            format_title_frame,
            "配置 Markdown 转换为 Word 文档时的格式标记处理方式",
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
        self._create_column_combobox(
            first_column,
            "正文格式标记处理:",
            self.formatting_mode_var,
            list(self._formatting_mode_mapping.values()),
            '应用格式：**粗体**转为真正的粗体格式\n保留标记：保留**粗体**原样显示\n清理标记：只显示"粗体"文本，无格式',
            self._on_formatting_mode_changed
        )
        
        # 第二列 - 小标题格式
        self._create_column_combobox(
            second_column,
            "小标题格式标记处理:",
            self.heading_formatting_mode_var,
            list(self._heading_formatting_mode_mapping.values()),
            '应用格式：覆盖标题样式默认格式，可实现部分加粗效果\n保留标记：保留标记原样显示\n清理标记：让Word标题样式格式自然生效\n注意：当标题样式本身是加粗时，"应用格式"可实现部分加粗',
            self._on_heading_formatting_mode_changed
        )
        
        # 第三列 - 表头格式
        self._create_column_combobox(
            third_column,
            "表头格式标记处理:",
            self.table_header_formatting_mode_var,
            list(self._table_header_formatting_mode_mapping.values()),
            '应用格式：**粗体**转为真正的粗体（可能与表格样式重复）\n保留标记：保留标记原样显示\n清理标记：让表格样式（如 firstRow 加粗）自然生效\n注意：表格样式通常已定义表头加粗，建议使用"清理标记"',
            self._on_table_header_formatting_mode_changed
        )
        
        # ========== 分隔符映射配置（MD转文档）==========
        self._create_md_to_docx_hr_config(frame)
        
        # ========== 表格样式配置 ==========
        self._create_table_style_config(frame)
        
        logger.debug("MD → DOCX 格式设置区域创建完成")
    
    def _create_md_to_docx_hr_config(self, parent_frame):
        """
        创建MD转文档的分隔符映射配置
        
        配置路径：conversion_config.horizontal_rule.md_to_docx.*
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 分隔线
        separator = tb.Separator(parent_frame, orient="horizontal")
        separator.pack(fill="x", pady=(15, 10))
        
        # 小节标题
        title_frame = tb.Frame(parent_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            title_frame,
            text="分隔符映射",
            font=(self.small_font, self.small_size, "bold")
        )
        title_label.pack(side="left")
        
        info = create_info_icon(
            title_frame,
            "配置 Markdown 分隔符转换为哪种 Word 分页符/分节符/分隔线\n\n"
            "分页符：Word 中的分页符\n"
            "分节符：Word 中的分节符（下一页类型）\n"
            "分隔线：段落底部边框（视觉分隔线）",
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 定义映射关系（Word分隔符选项）
        self._md_to_docx_hr_mapping = {
            "page_break": "分页符",
            "section_break": "分节符",
            "horizontal_rule_1": "分隔线 1",
            "horizontal_rule_2": "分隔线 2",
            "horizontal_rule_3": "分隔线 3",
            "ignore": "忽略"
        }
        
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
        
        # 获取当前值的显示文本
        dash_value = hr_config.get("dash", "page_break")
        asterisk_value = hr_config.get("asterisk", "section_continuous")
        underscore_value = hr_config.get("underscore", "section_break")
        
        dash_display = self._md_to_docx_hr_mapping.get(dash_value, "分页符")
        asterisk_display = self._md_to_docx_hr_mapping.get(asterisk_value, "分节符(连续)")
        underscore_display = self._md_to_docx_hr_mapping.get(underscore_value, "分节符(分页)")
        
        self.md_hr_dash_var = tk.StringVar(value=dash_display)
        self.md_hr_asterisk_var = tk.StringVar(value=asterisk_display)
        self.md_hr_underscore_var = tk.StringVar(value=underscore_display)
        
        # 第一列 - ---
        self._create_column_combobox(
            first_column,
            "--- (三个短横线):",
            self.md_hr_dash_var,
            list(self._md_to_docx_hr_mapping.values()),
            "Markdown --- 分隔符转换为 Word 分页符/分节符",
            self._on_md_hr_dash_changed
        )
        
        # 第二列 - ***
        self._create_column_combobox(
            second_column,
            "*** (三个星号):",
            self.md_hr_asterisk_var,
            list(self._md_to_docx_hr_mapping.values()),
            "Markdown *** 分隔符转换为 Word 分页符/分节符",
            self._on_md_hr_asterisk_changed
        )
        
        # 第三列 - ___
        self._create_column_combobox(
            third_column,
            "___ (三个下划线):",
            self.md_hr_underscore_var,
            list(self._md_to_docx_hr_mapping.values()),
            "Markdown ___ 分隔符转换为 Word 分页符/分节符\n注：分节符(分页) 实际生成为 分节符(下一页)",
            self._on_md_hr_underscore_changed
        )
    
    def _create_column_combobox(
        self,
        parent: tk.Widget,
        label_text: str,
        variable: tk.StringVar,
        values: list,
        tooltip: str,
        command=None,
        disabled: bool = False
    ) -> tb.Frame:
        """
        在指定列中创建带标签和信息图标的下拉框
        
        参数:
            parent: 父组件（左列或右列）
            label_text: 标签文本
            variable: 绑定的StringVar变量
            values: 下拉选项列表
            tooltip: 工具提示文本
            command: 选择改变时的回调函数
            disabled: 是否禁用下拉框
            
        返回:
            tb.Frame: 包含标签和下拉框的容器框架
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
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
            state="disabled" if disabled else "readonly",
            bootstyle="secondary"
        )
        combobox.pack(fill="x")
        
        # 绑定事件
        if command and not disabled:
            combobox.bind("<<ComboboxSelected>>", command)
        
        return container
    
    def _get_syntax_config_value(self, syntax_type: str, display_value: str) -> str:
        """
        将显示文本转换为配置值
        
        参数:
            syntax_type: 语法类型（bold, italic等）
            display_value: 下拉框显示的文本
        
        返回:
            str: 配置文件中使用的值
        """
        mapping = self._syntax_mappings.get(syntax_type, {})
        for config_value, display_text in mapping.items():
            if display_text == display_value:
                return config_value
        return "asterisk" if syntax_type in ["bold", "italic"] else "extended"
    
    def _get_preserve_formatting_config_value(self, display_value: str) -> bool:
        """
        将保留格式的显示文本转换为配置值（布尔值）
        
        参数:
            display_value: 下拉框显示的文本
        
        返回:
            bool: 配置文件中使用的布尔值
        """
        for config_value, display_text in self._preserve_formatting_mapping.items():
            if display_text == display_value:
                return config_value
        return True  # 默认保留格式
    
    def _get_formatting_mode_config_value(self, display_value: str) -> str:
        """
        将格式标记处理的显示文本转换为配置值
        
        参数:
            display_value: 下拉框显示的文本
        
        返回:
            str: 配置文件中使用的值
        """
        for config_value, display_text in self._formatting_mode_mapping.items():
            if display_text == display_value:
                return config_value
        return "apply"  # 默认应用格式
    
    def _get_preserve_heading_formatting_config_value(self, display_value: str) -> bool:
        """
        将小标题保留格式的显示文本转换为配置值（布尔值）
        
        参数:
            display_value: 下拉框显示的文本
        
        返回:
            bool: 配置文件中使用的布尔值
        """
        for config_value, display_text in self._preserve_heading_formatting_mapping.items():
            if display_text == display_value:
                return config_value
        return False  # 默认不保留格式
    
    def _get_heading_formatting_mode_config_value(self, display_value: str) -> str:
        """
        将小标题格式标记处理的显示文本转换为配置值
        
        参数:
            display_value: 下拉框显示的文本
        
        返回:
            str: 配置文件中使用的值
        """
        for config_value, display_text in self._heading_formatting_mode_mapping.items():
            if display_text == display_value:
                return config_value
        return "remove"  # 默认清理标记
    
    # --------------------------
    # 事件处理方法
    # --------------------------
    
    def _on_preserve_formatting_changed(self, event=None):
        """处理保留格式设置变更"""
        display_value = self.preserve_formatting_var.get()
        config_value = self._get_preserve_formatting_config_value(display_value)
        logger.info(f"DOCX转MD保留格式变更: {display_value} -> {config_value}")
        self.on_change("preserve_formatting", config_value)
    
    def _on_preserve_heading_formatting_changed(self, event=None):
        """处理小标题保留格式设置变更"""
        display_value = self.preserve_heading_formatting_var.get()
        config_value = self._get_preserve_heading_formatting_config_value(display_value)
        logger.info(f"DOCX转MD小标题保留格式变更: {display_value} -> {config_value}")
        self.on_change("preserve_heading_formatting", config_value)
    
    def _on_preserve_table_header_formatting_changed(self, event=None):
        """处理表头保留格式设置变更"""
        display_value = self.preserve_table_header_formatting_var.get()
        config_value = self._get_preserve_table_header_formatting_config_value(display_value)
        logger.info(f"DOCX转MD表头保留格式变更: {display_value} -> {config_value}")
        self.on_change("preserve_table_header_formatting", config_value)
    
    def _get_preserve_table_header_formatting_config_value(self, display_value: str) -> bool:
        """
        将表头保留格式的显示文本转换为配置值（布尔值）
        
        参数:
            display_value: 下拉框显示的文本
        
        返回:
            bool: 配置文件中使用的布尔值
        """
        for config_value, display_text in self._preserve_table_header_formatting_mapping.items():
            if display_text == display_value:
                return config_value
        return False  # 默认不保留格式
    
    def _on_formatting_mode_changed(self, event=None):
        """处理格式处理模式变更"""
        display_value = self.formatting_mode_var.get()
        config_value = self._get_formatting_mode_config_value(display_value)
        logger.info(f"MD转DOCX格式处理模式变更: {display_value} -> {config_value}")
        self.on_change("formatting_mode", config_value)
    
    def _on_heading_formatting_mode_changed(self, event=None):
        """处理小标题格式处理模式变更"""
        display_value = self.heading_formatting_mode_var.get()
        config_value = self._get_heading_formatting_mode_config_value(display_value)
        logger.info(f"MD转DOCX小标题格式处理模式变更: {display_value} -> {config_value}")
        self.on_change("heading_formatting_mode", config_value)
    
    def _on_table_header_formatting_mode_changed(self, event=None):
        """处理表头格式处理模式变更"""
        display_value = self.table_header_formatting_mode_var.get()
        config_value = self._get_table_header_formatting_mode_config_value(display_value)
        logger.info(f"MD转DOCX表头格式处理模式变更: {display_value} -> {config_value}")
        self.on_change("table_header_formatting_mode", config_value)
    
    def _get_table_header_formatting_mode_config_value(self, display_value: str) -> str:
        """
        将表头格式标记处理的显示文本转换为配置值
        
        参数:
            display_value: 下拉框显示的文本
        
        返回:
            str: 配置文件中使用的值
        """
        for config_value, display_text in self._table_header_formatting_mode_mapping.items():
            if display_text == display_value:
                return config_value
        return "remove"  # 默认清理标记
    
    def _on_bold_syntax_changed(self, event=None):
        """处理粗体语法变更"""
        display_value = self.bold_var.get()
        config_value = self._get_syntax_config_value("bold", display_value)
        logger.info(f"粗体语法变更: {display_value} -> {config_value}")
        self.on_change("bold", config_value)
    
    def _on_italic_syntax_changed(self, event=None):
        """处理斜体语法变更"""
        display_value = self.italic_var.get()
        config_value = self._get_syntax_config_value("italic", display_value)
        logger.info(f"斜体语法变更: {display_value} -> {config_value}")
        self.on_change("italic", config_value)
    
    def _on_strikethrough_syntax_changed(self, event=None):
        """处理删除线语法变更"""
        display_value = self.strikethrough_var.get()
        config_value = self._get_syntax_config_value("strikethrough", display_value)
        logger.info(f"删除线语法变更: {display_value} -> {config_value}")
        self.on_change("strikethrough", config_value)
    
    def _on_highlight_syntax_changed(self, event=None):
        """处理高亮语法变更"""
        display_value = self.highlight_var.get()
        config_value = self._get_syntax_config_value("highlight", display_value)
        logger.info(f"高亮语法变更: {display_value} -> {config_value}")
        self.on_change("highlight", config_value)
    
    def _on_superscript_syntax_changed(self, event=None):
        """处理上标语法变更"""
        display_value = self.superscript_var.get()
        config_value = self._get_syntax_config_value("superscript", display_value)
        logger.info(f"上标语法变更: {display_value} -> {config_value}")
        self.on_change("superscript", config_value)
    
    def _on_subscript_syntax_changed(self, event=None):
        """处理下标语法变更"""
        display_value = self.subscript_var.get()
        config_value = self._get_syntax_config_value("subscript", display_value)
        logger.info(f"下标语法变更: {display_value} -> {config_value}")
        self.on_change("subscript", config_value)
    
    def _on_unordered_list_changed(self, event=None):
        """处理无序列表标记变更"""
        display_value = self.unordered_list_var.get()
        config_value = self._get_syntax_config_value("unordered_list", display_value)
        logger.info(f"无序列表标记变更: {display_value} -> {config_value}")
        self.on_change("unordered_list", config_value)
    
    def _on_ordered_list_changed(self, event=None):
        """处理有序列表编号变更（暂未实现）"""
        display_value = self.ordered_list_var.get()
        config_value = self._get_syntax_config_value("ordered_list", display_value)
        logger.info(f"有序列表编号变更: {display_value} -> {config_value}")
        self.on_change("ordered_list", config_value)
    
    def _on_indent_spaces_changed(self, event=None):
        """处理列表缩进空格数变更"""
        display_value = self.indent_spaces_var.get()
        config_value = self._get_syntax_config_value("indent_spaces", display_value)
        logger.info(f"列表缩进空格数变更: {display_value} -> {config_value}")
        self.on_change("indent_spaces", config_value)
    
    # --------------------------
    # 分隔符转换事件处理方法
    # --------------------------
    
    def _get_docx_to_md_hr_config_value(self, display_value: str) -> str:
        """将文档转MD分隔符显示文本转换为配置值"""
        for config_value, display_text in self._docx_to_md_hr_mapping.items():
            if display_text == display_value:
                return config_value
        return "---"
    
    def _get_md_to_docx_hr_config_value(self, display_value: str) -> str:
        """将MD转文档分隔符显示文本转换为配置值"""
        for config_value, display_text in self._md_to_docx_hr_mapping.items():
            if display_text == display_value:
                return config_value
        return "page_break"
    
    def _on_docx_hr_page_break_changed(self, event=None):
        """处理文档转MD分页符映射变更"""
        display_value = self.docx_hr_page_break_var.get()
        config_value = self._get_docx_to_md_hr_config_value(display_value)
        logger.info(f"文档转MD分页符映射变更: {display_value} -> {config_value}")
        self.on_change("docx_hr_page_break", config_value)
    
    def _on_docx_hr_section_continuous_changed(self, event=None):
        """处理文档转MD分节符(连续)映射变更"""
        display_value = self.docx_hr_section_continuous_var.get()
        config_value = self._get_docx_to_md_hr_config_value(display_value)
        logger.info(f"文档转MD分节符(连续)映射变更: {display_value} -> {config_value}")
        self.on_change("docx_hr_section_continuous", config_value)
    
    def _on_docx_hr_section_break_changed(self, event=None):
        """处理文档转MD分节符映射变更"""
        display_value = self.docx_hr_section_break_var.get()
        config_value = self._get_docx_to_md_hr_config_value(display_value)
        logger.info(f"文档转MD分节符映射变更: {display_value} -> {config_value}")
        self.on_change("docx_hr_section_break", config_value)
    
    def _on_docx_hr_horizontal_rule_changed(self, event=None):
        """处理文档转MD分隔线映射变更"""
        display_value = self.docx_hr_horizontal_rule_var.get()
        config_value = self._get_docx_to_md_hr_config_value(display_value)
        logger.info(f"文档转MD分隔线映射变更: {display_value} -> {config_value}")
        self.on_change("docx_hr_horizontal_rule", config_value)
    
    def _on_md_hr_dash_changed(self, event=None):
        """处理MD转文档 --- 映射变更"""
        display_value = self.md_hr_dash_var.get()
        config_value = self._get_md_to_docx_hr_config_value(display_value)
        logger.info(f"MD转文档 --- 映射变更: {display_value} -> {config_value}")
        self.on_change("md_hr_dash", config_value)
    
    def _on_md_hr_asterisk_changed(self, event=None):
        """处理MD转文档 *** 映射变更"""
        display_value = self.md_hr_asterisk_var.get()
        config_value = self._get_md_to_docx_hr_config_value(display_value)
        logger.info(f"MD转文档 *** 映射变更: {display_value} -> {config_value}")
        self.on_change("md_hr_asterisk", config_value)
    
    def _on_md_hr_underscore_changed(self, event=None):
        """处理MD转文档 ___ 映射变更"""
        display_value = self.md_hr_underscore_var.get()
        config_value = self._get_md_to_docx_hr_config_value(display_value)
        logger.info(f"MD转文档 ___ 映射变更: {display_value} -> {config_value}")
        self.on_change("md_hr_underscore", config_value)
    
    # --------------------------
    # 表格样式配置
    # --------------------------
    
    def _create_table_style_config(self, parent_frame):
        """
        创建表格样式配置区域
        
        配置路径：style_table.md_to_docx.*
        
        支持两种模式：
        - 内置样式：三线表 / 网格表
        - 自定义样式：用户输入样式名称
        """
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 分隔线
        separator = tb.Separator(parent_frame, orient="horizontal")
        separator.pack(fill="x", pady=(15, 10))
        
        # 小节标题
        title_frame = tb.Frame(parent_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            title_frame,
            text="表格样式",
            font=(self.small_font, self.small_size, "bold")
        )
        title_label.pack(side="left")
        
        info = create_info_icon(
            title_frame,
            "配置 Markdown 表格转换为 Word 表格时使用的样式\n\n"
            "【内置样式】\n"
            "三线表/网格表：自动检测模板是否存在，若不存在则注入样式定义\n\n"
            "【自定义样式】\n"
            "检测模板中是否存在该样式，若不存在则创建一个网格表格式的同名样式",
            "info"
        )
        info.pack(side="left", padx=(self.layout_config.widget_spacing, 0))
        
        # 定义内置样式映射
        self._builtin_style_mapping = {
            "three_line_table": "三线表",
            "table_grid": "网格表"
        }
        
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
            text="使用内置样式",
            variable=self.table_style_mode_var,
            value="builtin",
            command=self._on_table_style_mode_changed
        )
        builtin_radio.pack(anchor="w", pady=(0, 5))
        
        # 内置样式下拉框
        builtin_display = self._builtin_style_mapping.get(builtin_key, "三线表")
        self.builtin_style_var = tk.StringVar(value=builtin_display)
        
        self.builtin_style_combobox = tb.Combobox(
            first_column,
            textvariable=self.builtin_style_var,
            values=list(self._builtin_style_mapping.values()),
            state="readonly" if style_mode == "builtin" else "disabled",
            bootstyle="secondary"
        )
        self.builtin_style_combobox.pack(fill="x", pady=(0, 5))
        self.builtin_style_combobox.bind("<<ComboboxSelected>>", self._on_builtin_style_changed)
        
        # === 第二列：自定义样式 ===
        custom_radio = tb.Radiobutton(
            second_column,
            text="使用自定义样式名称",
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
            self.builtin_style_combobox.configure(state="readonly")
            self.custom_style_entry.configure(state="disabled")
        else:
            self.builtin_style_combobox.configure(state="disabled")
            self.custom_style_entry.configure(state="normal")
        
        self.on_change("table_style_mode", mode)
    
    def _on_builtin_style_changed(self, event=None):
        """处理内置样式选择变更"""
        display_value = self.builtin_style_var.get()
        config_value = self._get_builtin_style_config_value(display_value)
        logger.info(f"内置表格样式变更: {display_value} -> {config_value}")
        self.on_change("builtin_style_key", config_value)
    
    def _on_custom_style_changed(self, event=None):
        """处理自定义样式名称变更"""
        custom_name = self.custom_style_var.get().strip()
        logger.info(f"自定义表格样式变更: {custom_name}")
        self.on_change("custom_style_name", custom_name)
    
    def _get_builtin_style_config_value(self, display_value: str) -> str:
        """将内置样式显示文本转换为配置值"""
        for config_value, display_text in self._builtin_style_mapping.items():
            if display_text == display_value:
                return config_value
        return "three_line_table"
    
    # --------------------------
    # 抽象方法实现
    # --------------------------
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前设置（实现抽象方法）
        
        返回:
            Dict: 配置字典
        """
        return {
            "preserve_formatting": self._get_preserve_formatting_config_value(self.preserve_formatting_var.get()),
            "preserve_heading_formatting": self._get_preserve_heading_formatting_config_value(self.preserve_heading_formatting_var.get()),
            "preserve_table_header_formatting": self._get_preserve_table_header_formatting_config_value(self.preserve_table_header_formatting_var.get()),
            "formatting_mode": self._get_formatting_mode_config_value(self.formatting_mode_var.get()),
            "heading_formatting_mode": self._get_heading_formatting_mode_config_value(self.heading_formatting_mode_var.get()),
            "table_header_formatting_mode": self._get_table_header_formatting_mode_config_value(self.table_header_formatting_mode_var.get()),
            "bold": self._get_syntax_config_value("bold", self.bold_var.get()),
            "italic": self._get_syntax_config_value("italic", self.italic_var.get()),
            "strikethrough": self._get_syntax_config_value("strikethrough", self.strikethrough_var.get()),
            "highlight": self._get_syntax_config_value("highlight", self.highlight_var.get()),
            "superscript": self._get_syntax_config_value("superscript", self.superscript_var.get()),
            "subscript": self._get_syntax_config_value("subscript", self.subscript_var.get()),
            "unordered_list": self._get_syntax_config_value("unordered_list", self.unordered_list_var.get()),
            "ordered_list": self._get_syntax_config_value("ordered_list", self.ordered_list_var.get()),
            "indent_spaces": self._get_syntax_config_value("indent_spaces", self.indent_spaces_var.get()),
            # 分隔符转换配置 - 文档转MD
            "docx_hr_page_break": self._get_docx_to_md_hr_config_value(self.docx_hr_page_break_var.get()),
            "docx_hr_section_break": self._get_docx_to_md_hr_config_value(self.docx_hr_section_break_var.get()),
            "docx_hr_horizontal_rule": self._get_docx_to_md_hr_config_value(self.docx_hr_horizontal_rule_var.get()),
            # 分隔符转换配置 - MD转文档
            "md_hr_dash": self._get_md_to_docx_hr_config_value(self.md_hr_dash_var.get()),
            "md_hr_asterisk": self._get_md_to_docx_hr_config_value(self.md_hr_asterisk_var.get()),
            "md_hr_underscore": self._get_md_to_docx_hr_config_value(self.md_hr_underscore_var.get()),
            # 表格样式配置
            "table_style_mode": self.table_style_mode_var.get(),
            "builtin_style_key": self._get_builtin_style_config_value(self.builtin_style_var.get()),
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
