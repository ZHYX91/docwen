"""
映射编辑器对话框模块
实现错误符号和自定义错别字的可视化编辑功能

重构版本：
- 提取配置常量
- 统一数据刷新流程
- 改进代码组织
- 增强类型提示
"""

import logging
import os
import tkinter as tk
from typing import Dict, List, Callable, Optional
from dataclasses import dataclass

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.tableview import Tableview

from gongwen_converter.utils.font_utils import get_small_font
from gongwen_converter.utils.dpi_utils import ScalableMixin

# 配置日志记录器
logger = logging.getLogger(__name__)


@dataclass
class LayoutConfig:
    """布局配置类"""
    # 窗口尺寸
    WINDOW_WIDTH: int = 850   # 加宽以容纳三列完整显示
    WINDOW_HEIGHT: int = 500  # 统一固定高度
    
    # 表格参数
    TABLE_HEIGHT: int = 20    # 固定表格行数
    PAGE_SIZE: int = 20       # 每页显示行数
    
    # 列宽配置（像素）- 总和约820px，适配850px窗口（预留空间给滚动条等组件）
    COL_WIDTH_KEY: int = 120        # 第1列：键（正确符号/词语）
    COL_WIDTH_VALUES: int = 480     # 第2列：值列表（错误列表）- 加宽以完整显示
    COL_WIDTH_COMMENT: int = 220    # 第3列：备注
    
    # 间距
    DESCRIPTION_WRAPLENGTH: int = 800  # 适应新窗口宽度
    BUTTON_PADDING: int = 5


@dataclass
class MappingType:
    """映射类型枚举"""
    SYMBOL: str = "symbol"
    TYPO: str = "typo"
    SENSITIVE: str = "sensitive"


class MappingEditorDialog(tb.Toplevel, ScalableMixin):
    """
    映射编辑器对话框类
    提供错误符号和自定义错别字的可视化编辑功能
    """
    
    def __init__(
        self, 
        parent: tk.Widget, 
        mapping_type: str, 
        current_mapping: Dict[str, List[str]],
        on_save: Callable[[Dict[str, List[str]]], None],
        config_file_path: str = None
    ):
        """
        初始化映射编辑器对话框
        
        参数:
            parent: 父窗口
            mapping_type: 映射类型 ('symbol', 'typo', 'sensitive')
            current_mapping: 当前映射数据
            on_save: 保存回调函数
            config_file_path: 配置文件路径（用于读取注释）
        """
        super().__init__(parent)
        
        # 初始化基本属性
        self.parent = parent
        self.mapping_type = mapping_type
        self.current_mapping = current_mapping.copy()
        self.on_save = on_save
        self.config = LayoutConfig()
        self.config_file_path = config_file_path
        
        # 初始化注释数据
        self.comments_dict: Dict[str, str] = {}
        self._load_comments()
        
        # 初始化UI组件引用
        self.table: Optional[Tableview] = None
        self.table_container: Optional[tb.Frame] = None
        
        # 配置窗口
        self._configure_window()
        
        # 创建界面
        self._create_interface()
        
        # 加载并显示数据
        self._refresh_data()
        
        logger.info(f"映射编辑器对话框初始化完成: {mapping_type}")

    # ==================== 工具方法 ====================
    
    def _configure_window(self) -> None:
        """配置窗口基本属性"""
        width = self.scale(self.config.WINDOW_WIDTH)
        height = self.scale(self.config.WINDOW_HEIGHT)
        
        self.title(self._get_dialog_title())
        self.geometry(f"{width}x{height}")
        self.resizable(True, True)
        
        # 设置最小窗口尺寸，防止窗口过小导致显示问题
        self.minsize(width, height)
        
        # 设置窗口图标
        self._setup_icon()
        
        # 使对话框模态
        self.transient(self.parent)
        self.grab_set()
    
    def _setup_icon(self) -> None:
        """设置窗口图标"""
        try:
            from gongwen_converter.utils.icon_utils import IconManager
            IconManager.set_window_icon(self)
        except Exception as e:
            logger.error(f"设置映射编辑器图标失败: {e}")
    
    # ==================== 文本生成方法 ====================
    
    def _get_dialog_title(self) -> str:
        """获取对话框标题"""
        titles = {
            MappingType.SYMBOL: "错误符号映射编辑器",
            MappingType.TYPO: "自定义错别字映射编辑器",
            MappingType.SENSITIVE: "敏感词映射编辑器"
        }
        return titles.get(self.mapping_type, "映射编辑器")
    
    def _get_column_headers(self) -> List[str]:
        """获取表格列头"""
        headers = {
            MappingType.SYMBOL: ["正确符号", "错误符号列表", "备注"],
            MappingType.TYPO: ["正确词语", "错误词语列表", "备注"],
            MappingType.SENSITIVE: ["敏感词", "例外情况列表", "备注"]
        }
        return headers.get(self.mapping_type, ["键", "值列表", "备注"])
    
    def _get_description_text(self) -> str:
        """获取说明文本"""
        descriptions = {
            MappingType.SYMBOL: "在此编辑错误符号映射。每行表示一个正确符号及其对应的常见错误形式。使用下方的按钮添加、删除或修改映射条目。",
            MappingType.TYPO: "在此编辑自定义错别字映射。每行表示一个正确词语及其对应的常见错误形式。使用下方的按钮添加、删除或修改映射条目。",
            MappingType.SENSITIVE: "在此编辑敏感词映射。每行表示一个敏感词及其例外情况。使用下方的按钮添加、删除或修改映射条目。"
        }
        return descriptions.get(self.mapping_type, "编辑映射数据")
    
    # ==================== UI创建方法 ====================
    
    def _create_interface(self) -> None:
        """创建对话框界面"""
        logger.debug("创建映射编辑器对话框界面")
        
        # 创建主框架
        main_frame = tb.Frame(self, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # 创建说明标签
        self._create_description_label(main_frame)
        
        # 创建操作按钮 - 先打包到底部
        self._create_operation_buttons(main_frame)
        
        # 创建表格 - 后打包，占据剩余空间
        self._create_table_container(main_frame)
        
        logger.debug("映射编辑器对话框界面创建完成")
    
    def _create_description_label(self, parent: tb.Frame) -> None:
        """创建说明标签"""
        description_label = tb.Label(
            parent,
            text=self._get_description_text(),
            bootstyle="secondary",
            wraplength=self.config.DESCRIPTION_WRAPLENGTH
        )
        description_label.pack(anchor="w", pady=(0, 10))
    
    def _create_operation_buttons(self, parent: tb.Frame) -> None:
        """创建操作按钮"""
        logger.debug("创建操作按钮")
        
        button_frame = tb.Frame(parent)
        button_frame.pack(side="bottom", fill="x", pady=(10, 0))
        
        # 左侧按钮组
        self._create_left_buttons(button_frame)
        
        # 右侧按钮组
        self._create_right_buttons(button_frame)
        
        logger.debug("操作按钮创建完成")
    
    def _create_left_buttons(self, parent: tb.Frame) -> None:
        """创建左侧按钮组（添加、编辑、删除）"""
        left_frame = tb.Frame(parent)
        left_frame.pack(side="left")
        
        buttons = [
            ("添加条目", "success", self._on_add_entry),
            ("编辑条目", "primary", self._on_edit_entry),
            ("删除条目", "danger", self._on_delete_entry)
        ]
        
        for text, style, command in buttons:
            btn = tb.Button(
                left_frame,
                text=text,
                bootstyle=style,
                command=command
            )
            btn.pack(side="left", padx=(0, self.config.BUTTON_PADDING))
    
    def _create_right_buttons(self, parent: tb.Frame) -> None:
        """创建右侧按钮组（保存、取消）"""
        right_frame = tb.Frame(parent)
        right_frame.pack(side="right")
        
        # 保存按钮
        save_button = tb.Button(
            right_frame,
            text="保存",
            bootstyle="success",
            command=self._on_save
        )
        save_button.pack(side="right", padx=(self.config.BUTTON_PADDING, 0))
        
        # 取消按钮
        cancel_button = tb.Button(
            right_frame,
            text="取消",
            bootstyle="secondary",
            command=self.destroy
        )
        cancel_button.pack(side="right", padx=(0, self.config.BUTTON_PADDING))
    
    def _create_table_container(self, parent: tb.Frame) -> None:
        """创建表格容器"""
        logger.debug("创建映射表格容器")
        
        # 创建表格框架，给予足够空间
        table_frame = tb.Frame(parent)
        table_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # 创建表格容器 - Tableview 内置滚动条和分页控件
        self.table_container = tb.Frame(table_frame)
        self.table_container.pack(fill="both", expand=True)
        
        # 初始创建表格
        self._rebuild_table()
        
        logger.debug("映射表格创建完成")
    
    def _rebuild_table(self) -> None:
        """重建表格（使用固定高度和自定义列宽）"""
        # 如果表格已存在，先销毁
        if self.table is not None:
            self.table.destroy()
        
        # 创建新表格 - 使用固定高度
        self.table = Tableview(
            self.table_container,
            coldata=self._get_column_headers(),
            rowdata=[],
            paginated=True,
            searchable=True,
            bootstyle="primary",
            height=self.config.TABLE_HEIGHT,
            pagesize=self.config.PAGE_SIZE
        )
        self.table.pack(fill="both", expand=True)
        
        # 手动设置列宽 - 使用 Treeview 的 column 方法
        try:
            self.table.view.column(0, width=self.scale(self.config.COL_WIDTH_KEY))
            self.table.view.column(1, width=self.scale(self.config.COL_WIDTH_VALUES))
            self.table.view.column(2, width=self.scale(self.config.COL_WIDTH_COMMENT))
            logger.debug(f"列宽设置成功: {self.config.COL_WIDTH_KEY}/{self.config.COL_WIDTH_VALUES}/{self.config.COL_WIDTH_COMMENT}")
        except Exception as e:
            logger.warning(f"设置列宽失败: {e}")
        
        # 绑定双击事件
        self.table.view.bind("<Double-Button-1>", self._on_table_double_click)
        
        logger.debug(f"表格重建完成，固定高度: {self.config.TABLE_HEIGHT} 行")
    
    # ==================== 数据操作方法 ====================
    
    def _load_comments(self) -> None:
        """从配置文件加载注释"""
        if not self.config_file_path:
            return
        
        try:
            from gongwen_converter.config.toml_operations import extract_inline_comments
            
            # 根据映射类型确定section名称
            section_map = {
                MappingType.SYMBOL: "symbol_map",
                MappingType.TYPO: "typos",
                MappingType.SENSITIVE: "sensitive_words"
            }
            section = section_map.get(self.mapping_type, "")
            
            if section:
                self.comments_dict = extract_inline_comments(self.config_file_path, section)
                logger.info(f"加载了 {len(self.comments_dict)} 条注释")
        except Exception as e:
            logger.error(f"加载注释失败: {e}")
            self.comments_dict = {}
    
    def _refresh_data(self) -> None:
        """
        刷新数据显示
        统一的数据刷新流程：加载数据
        """
        self._load_mapping_data()
    
    def _load_mapping_data(self) -> None:
        """加载映射数据到表格"""
        logger.debug("加载映射数据到表格")
        
        # 清空表格
        self.table.delete_rows()
        
        # 添加数据行（包含备注）
        for key, values in self.current_mapping.items():
            values_str = ", ".join(values)
            comment = self.comments_dict.get(key, "")
            self.table.insert_row("end", [key, values_str, comment])
        
        # 刷新表格
        self.table.load_table_data()
        
        logger.debug(f"加载了 {len(self.current_mapping)} 条映射数据")
    
    # ==================== 事件处理方法 ====================
    
    def _on_add_entry(self) -> None:
        """处理添加条目按钮点击事件"""
        logger.info("添加映射条目按钮被点击")
        
        dialog = EntryEditDialog(
            self,
            self.mapping_type,
            operation="add"
        )
        
        self.wait_window(dialog)
        
        if dialog.result:
            key = dialog.key_value.get().strip()
            values_str = dialog.values_value.get().strip()
            comment = dialog.comment_value.get().strip()
            
            if key and values_str:
                values = [v.strip() for v in values_str.split("\n") if v.strip()]
                self.current_mapping[key] = values
                
                # 保存备注
                if comment:
                    self.comments_dict[key] = comment
                
                # 统一刷新流程
                self._refresh_data()
                
                logger.info(f"添加映射条目: {key} -> {values}, 备注: {comment}")
    
    def _on_edit_entry(self) -> None:
        """处理编辑条目按钮点击事件"""
        logger.info("编辑映射条目按钮被点击")
        
        selected_rows = self.table.get_rows(selected=True)
        if not selected_rows:
            self._show_info_dialog("提示", "请先选择一个要编辑的条目。")
            return
        
        row_data = selected_rows[0].values
        key = row_data[0]
        current_values = self.current_mapping.get(key, [])
        current_comment = self.comments_dict.get(key, "")
        
        dialog = EntryEditDialog(
            self,
            self.mapping_type,
            operation="edit",
            current_key=key,
            current_values="\n".join(current_values),
            current_comment=current_comment
        )
        
        self.wait_window(dialog)
        
        if dialog.result:
            new_key = dialog.key_value.get().strip()
            values_str = dialog.values_value.get().strip()
            comment = dialog.comment_value.get().strip()
            
            if new_key and values_str:
                values = [v.strip() for v in values_str.split("\n") if v.strip()]
                
                # 如果键名改变，删除旧条目和旧备注
                if new_key != key:
                    del self.current_mapping[key]
                    if key in self.comments_dict:
                        del self.comments_dict[key]
                
                self.current_mapping[new_key] = values
                
                # 更新或删除备注
                if comment:
                    self.comments_dict[new_key] = comment
                elif new_key in self.comments_dict:
                    del self.comments_dict[new_key]
                
                # 统一刷新流程
                self._refresh_data()
                
                logger.info(f"编辑映射条目: {key} -> {new_key} -> {values}, 备注: {comment}")
    
    def _on_delete_entry(self) -> None:
        """处理删除条目按钮点击事件"""
        logger.info("删除映射条目按钮被点击")
        
        selected_rows = self.table.get_rows(selected=True)
        if not selected_rows:
            self._show_info_dialog("提示", "请先选择一个要删除的条目。")
            return
        
        row_data = selected_rows[0].values
        key = row_data[0]
        
        if self._show_confirm_dialog("确认删除", f"确定要删除条目 '{key}' 吗？"):
            if key in self.current_mapping:
                del self.current_mapping[key]
            
            # 同时删除备注
            if key in self.comments_dict:
                del self.comments_dict[key]
                
            # 统一刷新流程
            self._refresh_data()
            
            logger.info(f"删除映射条目: {key}")
    
    def _on_save(self) -> None:
        """处理保存按钮点击事件"""
        logger.info("保存映射数据和备注")
        
        try:
            # 如果有配置文件路径，直接保存映射和备注
            if self.config_file_path:
                from gongwen_converter.config.toml_operations import save_mapping_with_comments
                
                # 根据映射类型确定section名称
                section_map = {
                    MappingType.SYMBOL: "symbol_map",
                    MappingType.TYPO: "typos",
                    MappingType.SENSITIVE: "sensitive_words"
                }
                section = section_map.get(self.mapping_type, "")
                
                if section:
                    success = save_mapping_with_comments(
                        self.config_file_path,
                        section,
                        self.current_mapping,
                        self.comments_dict
                    )
                    
                    if not success:
                        raise Exception("保存到配置文件失败")
            
            # 调用回调函数通知外部（用于刷新配置等）
            self.on_save(self.current_mapping)
            
            # 直接关闭对话框，不再显示成功提示
            self.destroy()
        except Exception as e:
            logger.error(f"保存映射数据失败: {e}")
            self._show_error_dialog("错误", f"保存映射数据时发生错误: {str(e)}")
    
    def _on_table_double_click(self, event) -> None:
        """处理表格双击事件"""
        logger.debug("表格双击事件触发")
        
        item = self.table.view.identify_row(event.y)
        if item:
            self.table.view.selection_set(item)
            self._on_edit_entry()
    
    # ==================== 对话框辅助方法 ====================
    
    def _show_info_dialog(self, title: str, message: str) -> None:
        """显示信息对话框"""
        try:
            from gongwen_converter.gui.components.base_dialog import MessageBox
            MessageBox.showinfo(title, message, parent=self)
        except Exception as e:
            logger.error(f"显示信息对话框失败: {e}")
    
    def _show_error_dialog(self, title: str, message: str) -> None:
        """显示错误对话框"""
        try:
            from gongwen_converter.gui.components.base_dialog import MessageBox
            MessageBox.showerror(title, message, parent=self)
        except Exception as e:
            logger.error(f"显示错误对话框失败: {e}")
    
    def _show_confirm_dialog(self, title: str, message: str) -> bool:
        """显示确认对话框"""
        try:
            from gongwen_converter.gui.components.base_dialog import MessageBox
            return MessageBox.askyesno(title, message, parent=self)
        except Exception as e:
            logger.error(f"显示确认对话框失败: {e}")
            return False


class EntryEditDialog(tb.Toplevel, ScalableMixin):
    """
    条目编辑对话框类
    用于添加或编辑单个映射条目
    """
    
    # 布局常量
    DIALOG_WIDTH = 500
    DIALOG_HEIGHT = 600
    TEXT_HEIGHT = 6
    PADDING = 15
    
    def __init__(
        self, 
        parent: tk.Widget, 
        mapping_type: str, 
        operation: str,
        current_key: str = "",
        current_values: str = "",
        current_comment: str = ""
    ):
        """
        初始化条目编辑对话框
        
        参数:
            parent: 父窗口
            mapping_type: 映射类型
            operation: 操作类型 ('add' 或 'edit')
            current_key: 当前键值（编辑时使用）
            current_values: 当前值字符串（编辑时使用）
            current_comment: 当前备注（编辑时使用）
        """
        super().__init__(parent)
        
        self.parent = parent
        self.mapping_type = mapping_type
        self.operation = operation
        self.result = False
        
        # 获取字体配置
        self.small_font, self.small_size = get_small_font()
        
        # 创建界面变量
        self.key_value = tk.StringVar(value=current_key)
        self.values_value = tk.StringVar(value=current_values)
        self.comment_value = tk.StringVar(value=current_comment)
        self.values_text: Optional[tb.Text] = None
        self.comment_entry: Optional[tb.Entry] = None
        
        # 配置窗口
        self._configure_window()
        
        # 创建界面
        self._create_interface()
        
        logger.debug("条目编辑对话框初始化完成")
    
    def _configure_window(self) -> None:
        """配置窗口属性"""
        width = self.scale(self.DIALOG_WIDTH)
        height = self.scale(self.DIALOG_HEIGHT)
        
        self.title(self._get_dialog_title())
        self.geometry(f"{width}x{height}")
        self.resizable(False, False)
        
        # 设置窗口图标
        self._setup_icon()
        
        # 使对话框模态
        self.transient(self.parent)
        self.grab_set()
    
    def _setup_icon(self) -> None:
        """设置窗口图标"""
        try:
            from gongwen_converter.utils.icon_utils import IconManager
            IconManager.set_window_icon(self)
        except Exception as e:
            logger.error(f"设置条目编辑对话框图标失败: {e}")
    
    def _get_dialog_title(self) -> str:
        """获取对话框标题"""
        operation_text = "添加" if self.operation == "add" else "编辑"
        
        type_texts = {
            MappingType.SYMBOL: "错误符号映射",
            MappingType.TYPO: "错别字映射",
            MappingType.SENSITIVE: "敏感词映射"
        }
        type_text = type_texts.get(self.mapping_type, "映射")
        
        return f"{operation_text}{type_text}"
    
    def _get_key_label(self) -> str:
        """获取键标签文本"""
        labels = {
            MappingType.SYMBOL: "正确符号:",
            MappingType.TYPO: "正确词语:",
            MappingType.SENSITIVE: "敏感词:"
        }
        return labels.get(self.mapping_type, "键:")
    
    def _get_values_label(self) -> str:
        """获取值标签文本"""
        labels = {
            MappingType.SYMBOL: "错误符号列表 (每行一个):",
            MappingType.TYPO: "错误词语列表 (每行一个):",
            MappingType.SENSITIVE: "例外情况列表 (每行一个):"
        }
        return labels.get(self.mapping_type, "值列表 (每行一个):")
    
    def _get_example_text(self) -> str:
        """获取示例文本"""
        examples = {
            MappingType.SYMBOL: "示例:\n正确符号: （\n错误符号:\n【\n{",
            MappingType.TYPO: "示例:\n正确词语: 份号\n错误词语:\n分号\n分好",
            MappingType.SENSITIVE: "示例:\n敏感词: 违禁词\n例外情况:\n合法使用场景"
        }
        return examples.get(self.mapping_type, "示例:\n键: 值1\n值2")
    
    def _create_interface(self) -> None:
        """创建对话框界面"""
        logger.debug("创建条目编辑对话框界面")
        
        main_frame = tb.Frame(self, padding=self.PADDING)
        main_frame.pack(fill="both", expand=True)
        
        # 创建键输入区域
        self._create_key_input(main_frame)
        
        # 创建值输入区域
        self._create_values_input(main_frame)
        
        # 创建备注输入区域
        self._create_comment_input(main_frame)
        
        # 创建示例提示
        self._create_example_label(main_frame)
        
        # 创建按钮区域
        self._create_buttons(main_frame)
        
        logger.debug("条目编辑对话框界面创建完成")
    
    def _create_key_input(self, parent: tb.Frame) -> None:
        """创建键输入区域"""
        key_frame = tb.Frame(parent)
        key_frame.pack(fill="x", pady=(0, 10))
        
        key_label = tb.Label(
            key_frame,
            text=self._get_key_label(),
            bootstyle="primary"
        )
        key_label.pack(anchor="w")
        
        key_entry = tb.Entry(
            key_frame,
            textvariable=self.key_value,
            width=50
        )
        key_entry.pack(fill="x", pady=(5, 0))
    
    def _create_values_input(self, parent: tb.Frame) -> None:
        """创建值输入区域"""
        values_frame = tb.Frame(parent)
        values_frame.pack(fill="x", pady=(0, 10))
        
        values_label = tb.Label(
            values_frame,
            text=self._get_values_label(),
            bootstyle="primary"
        )
        values_label.pack(anchor="w")
        
        self.values_text = tb.Text(
            values_frame,
            width=50,
            height=self.TEXT_HEIGHT
        )
        self.values_text.pack(fill="x", pady=(5, 0))
        self.values_text.insert("1.0", self.values_value.get())
    
    def _create_comment_input(self, parent: tb.Frame) -> None:
        """创建备注输入区域"""
        comment_frame = tb.Frame(parent)
        comment_frame.pack(fill="x", pady=(0, 10))
        
        comment_label = tb.Label(
            comment_frame,
            text="备注说明 (可选):",
            bootstyle="primary"
        )
        comment_label.pack(anchor="w")
        
        self.comment_entry = tb.Entry(
            comment_frame,
            textvariable=self.comment_value,
            width=50
        )
        self.comment_entry.pack(fill="x", pady=(5, 0))
    
    def _create_example_label(self, parent: tb.Frame) -> None:
        """创建示例标签"""
        example_frame = tb.Frame(parent)
        example_frame.pack(fill="x", pady=(0, 15))
        
        example_label = tb.Label(
            example_frame,
            text=self._get_example_text(),
            bootstyle="secondary",
            font=(self.small_font, self.small_size)
        )
        example_label.pack(anchor="w")
    
    def _create_buttons(self, parent: tb.Frame) -> None:
        """创建按钮区域"""
        button_frame = tb.Frame(parent)
        button_frame.pack(fill="x")
        
        # 确定按钮
        ok_button = tb.Button(
            button_frame,
            text="确定",
            bootstyle="success",
            command=self._on_ok
        )
        ok_button.pack(side="right", padx=(5, 0))
        
        # 取消按钮
        cancel_button = tb.Button(
            button_frame,
            text="取消",
            bootstyle="secondary",
            command=self.destroy
        )
        cancel_button.pack(side="right", padx=(0, 5))
    
    def _on_ok(self) -> None:
        """处理确定按钮点击事件"""
        key = self.key_value.get().strip()
        values_str = self.values_text.get("1.0", "end-1c").strip()
        
        # 验证输入
        if not key:
            self._show_error_dialog("错误", "请输入正确的符号或词语。")
            return
        
        if not values_str:
            self._show_error_dialog("错误", "请输入错误符号或词语列表。")
            return
        
        # 设置结果并关闭
        self.values_value.set(values_str)
        self.result = True
        self.destroy()
    
    def _show_error_dialog(self, title: str, message: str) -> None:
        """显示错误对话框"""
        try:
            from gongwen_converter.gui.components.base_dialog import MessageBox
            MessageBox.showerror(title, message, parent=self)
        except Exception as e:
            logger.error(f"显示错误对话框失败: {e}")
