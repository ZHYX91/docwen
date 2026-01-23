"""
序号清理规则编辑器模块

提供序号清理规则的可视化编辑功能，支持：
- 规则新增/删除/复制
- 规则启用/禁用
- 规则顺序调整
- 正则表达式编辑
- 实时正则测试

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。
"""

import logging
import re
import tkinter as tk
from typing import Dict, List, Optional, Callable, Any
from dataclasses import dataclass, field

import ttkbootstrap as tb

from docwen.gui.settings.editors.base import BaseEditorDialog
from docwen.i18n import t

logger = logging.getLogger(__name__)


# ==================== 数据类 ====================

@dataclass
class CleaningRule:
    """清理规则数据类"""
    rule_id: str           # 规则ID（唯一标识）
    name: str              # 规则名称（用户自定义规则使用）
    description: str       # 规则说明（用户自定义规则使用）
    enabled: bool          # 是否启用
    regex: str             # 正则表达式（含占位符）
    is_system: bool        # 是否为系统规则
    examples: List[str] = field(default_factory=list)  # 示例文本
    name_key: str = ""     # 翻译键（系统规则使用）
    description_key: str = ""  # 说明翻译键（系统规则使用）
    
    @classmethod
    def from_config(cls, rule_id: str, config: Dict) -> 'CleaningRule':
        """从配置字典创建规则对象"""
        return cls(
            rule_id=rule_id,
            name=config.get("name", ""),
            description=config.get("description", ""),
            enabled=config.get("enabled", True),
            regex=config.get("regex", ""),
            is_system=config.get("is_system", False),
            examples=config.get("examples", []),
            name_key=config.get("name_key", ""),
            description_key=config.get("description_key", "")
        )
    
    def get_display_name(self) -> str:
        """获取显示名称（支持国际化）"""
        if self.name_key:
            translated = t(f"editors.numbering_clean.names.{self.name_key}")
            # 检查翻译是否成功（翻译失败时通常返回原键或带方括号的键）
            if not translated.startswith("[") and translated != f"editors.numbering_clean.names.{self.name_key}":
                return translated
        return self.name if self.name else self.rule_id
    
    def get_display_description(self) -> str:
        """获取显示说明（支持国际化）"""
        if self.description_key:
            translated = t(f"editors.numbering_clean.descriptions.{self.description_key}")
            if translated != f"editors.numbering_clean.descriptions.{self.description_key}":
                return translated
        return self.description
    
    def to_config(self) -> Dict:
        """转换为配置字典格式"""
        config = {
            "description": self.description,
            "enabled": self.enabled,
            "is_system": self.is_system,
            "regex": self.regex,
        }
        # 根据是否有 name_key 决定保存方式
        if self.name_key:
            config["name_key"] = self.name_key
        else:
            config["name"] = self.name
        
        if self.examples:
            config["examples"] = self.examples
        return config
    
    def copy(self, new_id: str, new_name: str) -> 'CleaningRule':
        """复制规则（用户自定义，不使用 name_key）"""
        return CleaningRule(
            rule_id=new_id,
            name=new_name,
            description=self.description,
            enabled=self.enabled,
            regex=self.regex,
            is_system=False,  # 复制的规则不是系统规则
            examples=self.examples.copy(),
            name_key=""       # 用户自定义规则不使用翻译键
        )


# ==================== 主编辑器对话框 ====================

class NumberingPatternsEditorDialog(BaseEditorDialog):
    """
    序号清理规则编辑器对话框
    
    提供序号清理规则的可视化编辑功能。
    """
    
    # 覆盖基类属性
    WINDOW_TITLE = t("editors.numbering_clean.window_title")
    CONFIG_FILE_NAME = "heading_numbering_clean.toml"
    LEFT_PANEL_TITLE = t("editors.numbering_clean.rule_list")
    
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
        self.placeholders: Dict[str, str] = {}
        
        # UI组件引用
        self.regex_var: Optional[tk.StringVar] = None
        self.regex_entry: Optional[tb.Entry] = None
        self.test_input_var: Optional[tk.StringVar] = None
        self.test_input_entry: Optional[tb.Entry] = None
        self.test_result_label: Optional[tb.Label] = None
        self.test_status_label: Optional[tb.Label] = None
        
        # 延迟测试定时器ID
        self._test_update_id: Optional[str] = None
        
        # 加载配置数据
        self._load_config_data()
        
        # 配置窗口
        self._configure_window()
        
        # 创建界面
        self._create_interface()
        
        # 选择第一个规则
        if self.item_order:
            self._select_item(self.item_order[0])
        
        logger.info("序号清理规则编辑器对话框初始化完成")
    
    # ==================== 数据加载 ====================
    
    def _load_config_data(self):
        """加载配置数据"""
        logger.debug("加载序号清理规则配置数据")
        
        try:
            # 从 heading_utils.py 加载占位符定义
            from docwen.utils.heading_utils import NUMBERING_PLACEHOLDERS
            self.placeholders = NUMBERING_PLACEHOLDERS
            
            # 从配置文件加载规则
            from docwen.config.toml_operations import read_toml_file
            
            config_data = read_toml_file(str(self.config_file_path))
            if not config_data:
                logger.warning("配置文件为空，使用默认配置")
                config_data = {}
            
            # 加载设置
            settings = config_data.get("settings", {})
            self.item_order = settings.get("order", [])
            
            # 加载规则
            rules_data = config_data.get("rules", {})
            for rule_id, rule_config in rules_data.items():
                rule = CleaningRule.from_config(rule_id, rule_config)
                self.items[rule_id] = rule
            
            # 确保 order 中的所有规则都存在
            self.item_order = [rid for rid in self.item_order if rid in self.items]
            
            # 添加不在 order 中的规则
            for rule_id in self.items:
                if rule_id not in self.item_order:
                    self.item_order.append(rule_id)
            
            logger.info(f"加载了 {len(self.items)} 条清理规则")
            
        except Exception as e:
            logger.error(f"加载配置数据失败: {e}", exc_info=True)
            self._show_error(t("editors.numbering_clean.load_failed"), t("editors.numbering_clean.load_failed_message", error=e))
    
    # ==================== UI创建方法 ====================
    
    def _create_interface(self):
        """创建界面"""
        logger.debug("创建序号清理规则编辑器界面")
        
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
        
        logger.debug("序号清理规则编辑器界面创建完成")
    
    def _create_left_panel_content(self, parent: tb.Frame):
        """创建左侧面板内容"""
        left_frame = self._create_left_panel(parent)
        
        # 规则列表
        self._create_item_list(left_frame)
        
        # 刷新列表
        self._refresh_item_list()
        
        # 操作按钮
        self._create_action_buttons(left_frame)
    
    def _create_right_panel(self, parent: tb.Frame):
        """创建右侧面板（规则编辑）"""
        right_frame = tb.Frame(parent)
        right_frame.pack(side="left", fill="both", expand=True)
        
        # 规则编辑区域
        self._create_item_editor(right_frame)
        
        # 正则测试区域
        self._create_regex_test(right_frame)
        
        # 占位符说明
        self._create_placeholder_info(right_frame)
    
    def _create_item_editor(self, parent: tb.Frame):
        """创建项目编辑区（名称、说明和正则）"""
        frame = tb.Labelframe(parent, text=t("editors.numbering_clean.rule_edit"), padding=5)
        frame.pack(fill="x", pady=(0, 10))
        
        # 规则名称
        name_frame = tb.Frame(frame)
        name_frame.pack(fill="x", pady=(0, 5))
        
        tb.Label(
            name_frame,
            text=t("editors.numbering_clean.name"),
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
        
        # 规则说明
        desc_frame = tb.Frame(frame)
        desc_frame.pack(fill="x", pady=(0, 5))
        
        tb.Label(
            desc_frame,
            text=t("editors.numbering_clean.description"),
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
        
        # 正则表达式
        regex_frame = tb.Frame(frame)
        regex_frame.pack(fill="x")
        
        tb.Label(
            regex_frame,
            text=t("editors.numbering_clean.regex"),
            width=12,
            anchor="e"
        ).pack(side="left")
        
        self.regex_var = tk.StringVar()
        self.regex_entry = tb.Entry(
            regex_frame,
            textvariable=self.regex_var,
            width=40,
            font=(self.small_font, self.small_size)
        )
        self.regex_entry.pack(side="left", padx=(5, 0), fill="x", expand=True)
        self.regex_var.trace_add("write", lambda *_: self._on_regex_changed())
    
    def _create_regex_test(self, parent: tb.Frame):
        """创建正则测试区域"""
        frame = tb.Labelframe(parent, text=t("editors.numbering_clean.regex_test"), padding=5)
        frame.pack(fill="x", pady=(0, 10))
        
        # 测试输入
        input_frame = tb.Frame(frame)
        input_frame.pack(fill="x", pady=(0, 5))
        
        tb.Label(
            input_frame,
            text=t("editors.numbering_clean.input"),
            width=12,
            anchor="e"
        ).pack(side="left")
        
        self.test_input_var = tk.StringVar()
        self.test_input_entry = tb.Entry(
            input_frame,
            textvariable=self.test_input_var,
            width=40
        )
        self.test_input_entry.pack(side="left", padx=(5, 0), fill="x", expand=True)
        self.test_input_var.trace_add("write", lambda *_: self._schedule_test_update())
        
        # 测试结果
        result_frame = tb.Frame(frame)
        result_frame.pack(fill="x", pady=(0, 5))
        
        tb.Label(
            result_frame,
            text=t("editors.numbering_clean.result"),
            width=12,
            anchor="e"
        ).pack(side="left")
        
        self.test_result_label = tb.Label(
            result_frame,
            text="",
            font=(self.small_font, self.small_size),
            anchor="w"
        )
        self.test_result_label.pack(side="left", padx=(5, 0), fill="x", expand=True)
        
        # 测试状态
        status_frame = tb.Frame(frame)
        status_frame.pack(fill="x")
        
        tb.Label(
            status_frame,
            text=t("editors.numbering_clean.status"),
            width=12,
            anchor="e"
        ).pack(side="left")
        
        self.test_status_label = tb.Label(
            status_frame,
            text="",
            font=(self.small_font, self.small_size)
        )
        self.test_status_label.pack(side="left", padx=(5, 0))
        
        # 示例按钮
        tb.Button(
            status_frame,
            text=t("editors.numbering_clean.load_example"),
            command=self._load_example,
            bootstyle="link"
        ).pack(side="right")
    
    def _create_placeholder_info(self, parent: tb.Frame):
        """创建占位符说明区域"""
        frame = tb.Labelframe(parent, text=t("editors.numbering_clean.available_placeholders"), padding=5)
        frame.pack(fill="both", expand=True)
        
        # 创建可滚动的文本区域
        info_text = tb.Text(
            frame,
            height=8,
            font=(self.small_font, self.small_size),
            state="normal",
            wrap="word"
        )
        info_text.pack(fill="both", expand=True)
        
        # 填充占位符说明
        placeholder_info = t("editors.numbering_clean.placeholder_info")
        
        info_text.insert("1.0", placeholder_info)
        info_text.configure(state="disabled")
    
    # ==================== 项目操作方法（实现基类抽象方法） ====================
    
    def _create_new_item(self, item_id: str) -> CleaningRule:
        """创建新规则"""
        return CleaningRule(
            rule_id=item_id,
            name=t("editors.numbering_clean.new_rule"),
            description="",
            enabled=True,
            regex="",
            is_system=False,
            examples=[]
        )
    
    def _copy_item(self, source_item: CleaningRule, new_id: str) -> CleaningRule:
        """复制规则"""
        copy_suffix = t("editors.numbering_clean.copy_suffix")
        return source_item.copy(new_id, f"{source_item.get_display_name()} {copy_suffix}")
    
    def _get_item_name(self, item: CleaningRule) -> str:
        """获取规则显示名称（覆盖基类方法以支持国际化）"""
        return item.get_display_name()
    
    def _get_item_description(self, item: CleaningRule) -> str:
        """获取规则显示说明（覆盖基类方法以支持国际化）"""
        return item.get_display_description()
    
    def _on_item_selected(self, item: CleaningRule):
        """规则被选中后的回调"""
        # 更新正则表达式
        if self.regex_var:
            self.regex_var.set(item.regex)
        
        # 清空测试区域
        if self.test_input_var:
            self.test_input_var.set("")
        if self.test_result_label:
            self.test_result_label.configure(text="")
        if self.test_status_label:
            self.test_status_label.configure(text="")
        
        # 如果有示例，加载第一个
        if item.examples:
            self.test_input_var.set(item.examples[0])
            self._run_regex_test()
    
    def _on_no_item_selected(self):
        """没有规则被选中时清空编辑区"""
        self._is_loading = True
        try:
            if self.name_var:
                self.name_var.set("")
            if self.description_text:
                self.description_text.delete("1.0", tk.END)
            if self.regex_var:
                self.regex_var.set("")
            if self.test_input_var:
                self.test_input_var.set("")
            if self.test_result_label:
                self.test_result_label.configure(text="")
            if self.test_status_label:
                self.test_status_label.configure(text="")
        finally:
            self._is_loading = False
    
    # ==================== 事件处理方法 ====================
    
    def _on_regex_changed(self):
        """处理正则表达式变更"""
        if self._is_loading or not self.current_item_id:
            return
        
        self._mark_changed()
        
        # 更新当前规则数据
        rule = self.items.get(self.current_item_id)
        if rule and self.regex_var:
            rule.regex = self.regex_var.get()
        
        # 延迟执行测试
        self._schedule_test_update()
    
    def _schedule_test_update(self):
        """安排测试更新（延迟300ms）"""
        if self._test_update_id:
            self.after_cancel(self._test_update_id)
        
        self._test_update_id = self.after(300, self._run_regex_test)
    
    def _run_regex_test(self):
        """执行正则表达式测试"""
        self._test_update_id = None
        
        if not self.regex_var or not self.test_input_var:
            return
        
        regex_str = self.regex_var.get()
        test_input = self.test_input_var.get()
        
        if not regex_str or not test_input:
            if self.test_result_label:
                self.test_result_label.configure(text="")
            if self.test_status_label:
                self.test_status_label.configure(text="")
            return
        
        try:
            # 替换占位符
            compiled_regex = self._replace_placeholders(regex_str)
            
            # 编译正则表达式
            pattern = re.compile(compiled_regex)
            
            # 执行匹配
            result = pattern.sub('', test_input)
            
            if result != test_input:
                # 匹配成功
                if self.test_result_label:
                    self.test_result_label.configure(text=result if result else "(空)")
                if self.test_status_label:
                    self.test_status_label.configure(text=f"✅ {t('editors.numbering_clean.match_success')}", bootstyle="success")
            else:
                # 无匹配
                if self.test_result_label:
                    self.test_result_label.configure(text=test_input)
                if self.test_status_label:
                    self.test_status_label.configure(text=f"⚠️ {t('editors.numbering_clean.no_match')}", bootstyle="warning")
                
        except re.error as e:
            if self.test_result_label:
                self.test_result_label.configure(text="")
            if self.test_status_label:
                self.test_status_label.configure(text=f"❌ {t('editors.numbering_clean.regex_error')}: {e}", bootstyle="danger")
        except Exception as e:
            if self.test_result_label:
                self.test_result_label.configure(text="")
            if self.test_status_label:
                self.test_status_label.configure(text=f"❌ {t('editors.numbering_clean.test_error')}: {e}", bootstyle="danger")
    
    def _replace_placeholders(self, regex: str) -> str:
        """替换正则表达式中的占位符"""
        result = regex
        for name, value in self.placeholders.items():
            result = result.replace(f"{{{name}}}", value)
        return result
    
    def _load_example(self):
        """加载示例文本"""
        if not self.current_item_id:
            return
        
        rule = self.items.get(self.current_item_id)
        if not rule or not rule.examples:
            self._show_info(t("editors.numbering_clean.prompt"), t("editors.numbering_clean.no_example_text"))
            return
        
        # 创建示例选择菜单
        menu = tk.Menu(self, tearoff=0)
        
        for example in rule.examples:
            menu.add_command(
                label=example,
                command=lambda ex=example: self._set_test_input(ex)
            )
        
        # 显示菜单
        try:
            if self.test_input_entry:
                x = self.test_input_entry.winfo_rootx()
                y = self.test_input_entry.winfo_rooty() + self.test_input_entry.winfo_height()
                menu.tk_popup(x, y)
        finally:
            menu.grab_release()
    
    def _set_test_input(self, text: str):
        """设置测试输入"""
        if self.test_input_var:
            self.test_input_var.set(text)
        self._run_regex_test()
    
    # ==================== 保存 ====================
    
    def _on_save(self):
        """保存配置"""
        logger.info("保存序号清理规则配置")
        
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
            
            # 更新设置 - order
            if "settings" not in doc:
                doc.add("settings", table())
            
            order_array = array()
            for rule_id in self.item_order:
                order_array.append(rule_id)
            order_array.multiline(True)  # 保持多行格式
            doc["settings"]["order"] = order_array
            
            # 更新规则
            if "rules" not in doc:
                doc.add("rules", table())
            
            rules_table = doc["rules"]
            
            # 删除不再存在的规则（保护系统规则）
            existing_rule_ids = list(rules_table.keys())
            current_rule_ids = set(self.items.keys())
            for rule_id in existing_rule_ids:
                if rule_id not in current_rule_ids:
                    # 双重保险：检查原规则是否为系统规则
                    if not rules_table[rule_id].get("is_system", False):
                        del rules_table[rule_id]
                        logger.info(f"从配置文件中删除规则: {rule_id}")
            
            # 更新每条规则
            for rule_id, rule in self.items.items():
                if rule_id not in rules_table:
                    rules_table.add(rule_id, table())
                
                rule_table = rules_table[rule_id]
                
                # 根据是否有 name_key 决定保存方式
                if rule.name_key:
                    rule_table["name_key"] = rule.name_key
                    # 删除可能存在的旧 name 字段
                    if "name" in rule_table:
                        del rule_table["name"]
                else:
                    rule_table["name"] = rule.name
                    # 删除可能存在的旧 name_key 字段
                    if "name_key" in rule_table:
                        del rule_table["name_key"]
                
                rule_table["description"] = rule.description
                rule_table["enabled"] = rule.enabled
                rule_table["is_system"] = rule.is_system
                rule_table["regex"] = rule.regex
                
                if rule.examples:
                    examples_array = array()
                    for ex in rule.examples:
                        examples_array.append(ex)
                    examples_array.multiline(True)  # 保持多行格式
                    rule_table["examples"] = examples_array
            
            # 写入文件
            success = write_toml_document(str(self.config_file_path), doc)
            
            if success:
                # 重新加载配置
                self.config_manager.reload_configs()
                
                # 调用回调
                if self.on_save_callback:
                    self.on_save_callback()
                
                self._clear_changed()
                logger.info("序号清理规则配置保存成功")
                
                # 关闭对话框
                self.destroy()
            else:
                self._show_error(t("editors.numbering_clean.save_failed"), t("editors.numbering_clean.save_failed_message"))
                
        except Exception as e:
            logger.error(f"保存配置失败: {e}", exc_info=True)
            self._show_error(t("editors.numbering_clean.save_failed"), t("editors.numbering_clean.save_failed_with_error", error=e))
