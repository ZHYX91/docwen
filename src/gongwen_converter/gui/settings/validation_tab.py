"""
校对设置选项卡模块

实现设置对话框的校对设置选项卡，包含：
- 校对选项开关（标点配对、符号校对、错别字、敏感词）
- 错误符号映射配置
- 错别字映射配置
- 敏感词映射配置
"""

import logging
import tkinter as tk
from typing import Dict, Any, List, Callable

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from gongwen_converter.gui.settings.base_tab import BaseSettingsTab
from gongwen_converter.gui.settings.config import SectionStyle

logger = logging.getLogger(__name__)


class ValidationTab(BaseSettingsTab):
    """
    校对设置选项卡类
    
    管理文档校对相关的所有配置选项。
    包含校对开关、符号映射、错别字和敏感词配置。
    """
    
    def __init__(self, parent, config_manager: any, on_change: Callable[[str, Any], None]):
        """
        初始化校对设置选项卡
        
        参数：
            parent: 父组件
            config_manager: 配置管理器实例
            on_change: 设置变更回调函数
        """
        super().__init__(parent, config_manager, on_change)
        logger.info("校对设置选项卡初始化完成")
    
    def _create_interface(self):
        """
        创建选项卡界面
        
        创建四个主要设置区域：
        1. 校对选项 - 各类校对功能开关
        2. 错误符号配置 - 符号映射编辑
        3. 错别字配置 - 错别字映射编辑
        4. 敏感词配置 - 敏感词映射编辑
        """
        logger.debug("开始创建校对设置选项卡界面")
        
        self._create_validation_options_section()
        self._create_symbol_mapping_section()
        self._create_custom_typos_section()
        self._create_sensitive_words_section()
        
        logger.debug("校对设置选项卡界面创建完成")
    
    def _create_validation_options_section(self):
        """
        创建校对选项设置区域
        
        包含四个校对功能开关：
        - 标点配对检查
        - 符号校对
        - 错别字校对
        - 敏感词匹配
        
        配置路径：
        - symbol_settings.engine_settings.enable_symbol_pairing
        - symbol_settings.engine_settings.enable_symbol_correction
        - typos_settings.engine_settings.enable_typos_rule
        - sensitive_words_settings.engine_settings.enable_sensitive_word
        """
        logger.debug("创建校对选项设置区域")
        
        # 创建区域框架
        frame = self.create_section_frame(
            self.scrollable_frame,
            "校对选项",
            SectionStyle.PRIMARY
        )
        
        # 合并所有引擎设置
        symbol_engine_settings = self.config_manager.get_symbol_engine_settings()
        typos_engine_settings = self.config_manager.get_typos_engine_settings()
        sensitive_engine_settings = self.config_manager.get_sensitive_words_engine_settings()
        engine_options = {**symbol_engine_settings, **typos_engine_settings, **sensitive_engine_settings}
        
        # 标点配对检查
        self.symbol_pairing_var = tk.BooleanVar(value=engine_options.get("enable_symbol_pairing", True))
        self.create_checkbox_with_info(
            frame,
            "启用标点配对检查",
            self.symbol_pairing_var,
            "检查括号、引号等是否成对出现。",
            lambda: self.on_change("enable_symbol_pairing", self.symbol_pairing_var.get())
        )
        
        # 符号校对
        self.symbol_correction_var = tk.BooleanVar(value=engine_options.get("enable_symbol_correction", True))
        self.create_checkbox_with_info(
            frame,
            "启用符号校对",
            self.symbol_correction_var,
            "自动纠正全角/半角符号，如将英文逗号纠正为中文逗号。",
            lambda: self.on_change("enable_symbol_correction", self.symbol_correction_var.get())
        )
        
        # 错别字校对
        self.typos_rule_var = tk.BooleanVar(value=engine_options.get("enable_typos_rule", True))
        self.create_checkbox_with_info(
            frame,
            "启用错别字校对",
            self.typos_rule_var,
            "根据自定义词典检查并纠正错别字。",
            lambda: self.on_change("enable_typos_rule", self.typos_rule_var.get())
        )
        
        # 敏感词匹配
        self.sensitive_word_var = tk.BooleanVar(value=engine_options.get("enable_sensitive_word", True))
        self.create_checkbox_with_info(
            frame,
            "启用敏感词匹配",
            self.sensitive_word_var,
            "检查文本中是否包含需要关注的敏感词。",
            lambda: self.on_change("enable_sensitive_word", self.sensitive_word_var.get())
        )
        
        logger.debug("校对选项设置区域创建完成")
    
    def _create_symbol_mapping_section(self):
        """
        创建错误符号配置区域
        
        提供按钮打开映射编辑器，用于配置符号的正确形式及其常见错误。
        
        配置路径：symbol_settings.toml - symbol_map
        """
        logger.debug("创建错误符号配置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "错误符号配置",
            SectionStyle.INFO
        )
        
        # 说明文本
        desc_label = tb.Label(
            frame,
            text="定义标点符号的正确形式及其常见错误。",
            bootstyle="secondary",
            wraplength=400
        )
        desc_label.pack(anchor="w", pady=(0, 10))
        
        # 编辑按钮
        edit_button = tb.Button(
            frame,
            text="查看和编辑：错误符号",
            command=self._on_edit_symbol_mapping,
            bootstyle="info"
        )
        edit_button.pack(anchor="e", pady=(10, 0))
        
        logger.debug("错误符号配置区域创建完成")
    
    def _create_custom_typos_section(self):
        """
        创建错别字配置区域
        
        提供按钮打开映射编辑器，用于配置词语的正确形式及其常见错误。
        
        配置路径：typos_settings.toml - typos
        """
        logger.debug("创建错别字配置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "错别字配置",
            SectionStyle.WARNING
        )
        
        # 说明文本
        desc_label = tb.Label(
            frame,
            text="定义词语的正确形式及其常见错误。",
            bootstyle="secondary",
            wraplength=400
        )
        desc_label.pack(anchor="w", pady=(0, 10))
        
        # 编辑按钮
        edit_button = tb.Button(
            frame,
            text="查看和编辑：错别字",
            command=self._on_edit_custom_typos,
            bootstyle="warning"
        )
        edit_button.pack(anchor="e", pady=(10, 0))
        
        logger.debug("错别字配置区域创建完成")
    
    def _create_sensitive_words_section(self):
        """
        创建敏感词配置区域
        
        提供按钮打开映射编辑器，用于配置敏感词及其例外情况。
        
        配置路径：sensitive_words.toml - sensitive_words
        """
        logger.debug("创建敏感词配置区域")
        
        frame = self.create_section_frame(
            self.scrollable_frame,
            "敏感词配置",
            SectionStyle.DANGER
        )
        
        # 说明文本
        desc_label = tb.Label(
            frame,
            text="定义需要匹配的敏感词及其例外情况。",
            bootstyle="secondary",
            wraplength=400
        )
        desc_label.pack(anchor="w", pady=(0, 10))
        
        # 编辑按钮
        edit_button = tb.Button(
            frame,
            text="查看和编辑：敏感词",
            command=self._on_edit_sensitive_words,
            bootstyle="danger"
        )
        edit_button.pack(anchor="e", pady=(10, 0))
        
        logger.debug("敏感词配置区域创建完成")
    
    def _on_edit_symbol_mapping(self):
        """处理编辑符号映射按钮点击事件"""
        logger.info("打开符号映射编辑器")
        config_file_path = self._get_config_file_path("symbol_settings.toml")
        from gongwen_converter.config.toml_operations import read_toml_file
        config_data = read_toml_file(config_file_path)
        symbol_map = config_data.get("symbol_map", {})
        self._open_editor("symbol", symbol_map, self._on_symbol_mapping_saved, config_file_path)
    
    def _on_edit_custom_typos(self):
        """处理编辑自定义错别字按钮点击事件"""
        logger.info("打开错别字映射编辑器")
        config_file_path = self._get_config_file_path("typos_settings.toml")
        from gongwen_converter.config.toml_operations import read_toml_file
        config_data = read_toml_file(config_file_path)
        custom_typos = config_data.get("typos", {})
        self._open_editor("typo", custom_typos, self._on_custom_typos_saved, config_file_path)
    
    def _on_edit_sensitive_words(self):
        """处理编辑敏感词按钮点击事件"""
        logger.info("打开敏感词映射编辑器")
        config_file_path = self._get_config_file_path("sensitive_words.toml")
        from gongwen_converter.config.toml_operations import read_toml_file
        config_data = read_toml_file(config_file_path)
        sensitive_words = config_data.get("sensitive_words", {})
        self._open_editor("sensitive", sensitive_words, self._on_sensitive_words_saved, config_file_path)
    
    def _get_config_file_path(self, filename: str) -> str:
        """
        获取配置文件的完整路径
        
        参数：
            filename: 配置文件名
            
        返回：
            str: 配置文件的完整路径
        """
        import os
        try:
            config_dir = self.config_manager._config_dir
            return os.path.join(config_dir, filename)
        except Exception as e:
            logger.error(f"获取配置文件路径失败: {e}")
            return None
    
    def _open_editor(self, editor_type, data, save_callback, config_file_path=None):
        """
        打开映射编辑器对话框
        
        参数：
            editor_type: 编辑器类型 ('symbol', 'typo', 'sensitive')
            data: 映射数据
            save_callback: 保存回调函数
            config_file_path: 配置文件路径
        """
        try:
            from .mapping_editor import MappingEditorDialog
            editor = MappingEditorDialog(
                self,
                editor_type,
                data,
                save_callback,
                config_file_path=config_file_path
            )
            self.wait_window(editor)
        except ImportError as e:
            logger.error(f"导入映射编辑器失败: {e}")
            self._show_error_dialog("错误", "无法加载映射编辑器。")
    
    def _on_symbol_mapping_saved(self, new_mapping: Dict[str, List[str]]):
        """处理符号映射保存事件"""
        logger.info("符号映射已保存")
        config_file_path = self._get_config_file_path("symbol_settings.toml")
        self._save_mapping_with_comments(config_file_path, "symbol_map", new_mapping, "符号映射")
    
    def _on_custom_typos_saved(self, new_mapping: Dict[str, List[str]]):
        """处理错别字映射保存事件"""
        logger.info("错别字映射已保存")
        config_file_path = self._get_config_file_path("typos_settings.toml")
        self._save_mapping_with_comments(config_file_path, "typos", new_mapping, "错别字映射")
    
    def _on_sensitive_words_saved(self, new_mapping: Dict[str, List[str]]):
        """处理敏感词映射保存事件"""
        logger.info("敏感词映射已保存")
        config_file_path = self._get_config_file_path("sensitive_words.toml")
        self._save_mapping_with_comments(config_file_path, "sensitive_words", new_mapping, "敏感词映射")
    
    def _save_mapping_with_comments(self, filepath, section, mapping_data, title):
        """
        保存映射数据和备注到TOML文件
        
        参数：
            filepath: 文件路径
            section: 配置节名称
            mapping_data: 映射数据
            title: 标题（用于日志）
        """
        try:
            from gongwen_converter.config.toml_operations import (
                save_mapping_with_comments,
                extract_inline_comments
            )
            
            # 重新提取注释（保留现有注释）
            comments_data = extract_inline_comments(filepath, section)
            
            # 保存映射和注释
            success = save_mapping_with_comments(
                filepath,
                section,
                mapping_data,
                comments_data
            )
            
            if success:
                logger.info(f"{title}已成功保存到配置文件（包含备注）")
                self._refresh_config_manager_cache(section)
            else:
                logger.error(f"保存{title}到配置文件失败")
                self._show_error_dialog("错误", f"保存{title}失败。")
                
        except Exception as e:
            logger.error(f"保存{title}时发生错误: {e}")
            self._show_error_dialog("错误", f"保存时发生错误: {str(e)}")
    
    def _refresh_config_manager_cache(self, section: str):
        """
        刷新配置管理器缓存
        
        参数：
            section: 配置节名称
        """
        try:
            self.config_manager.reload_configs()
            logger.debug(f"已刷新 {section} 配置管理器缓存")
        except Exception as e:
            logger.warning(f"刷新配置管理器缓存失败: {e}")
    
    def _show_error_dialog(self, title: str, message: str):
        """显示错误对话框"""
        try:
            from gongwen_converter.gui.components.base_dialog import MessageBox
            MessageBox.showerror(title, message, parent=self.winfo_toplevel())
            logger.debug(f"显示错误对话框: {title}")
        except Exception as e:
            logger.error(f"显示错误对话框失败: {e}")
    
    def get_settings(self) -> Dict[str, Any]:
        """
        获取当前选项卡的设置
        
        返回：
            Dict[str, Any]: 当前所有设置项的值
        """
        settings = {
            "enable_symbol_pairing": self.symbol_pairing_var.get(),
            "enable_symbol_correction": self.symbol_correction_var.get(),
            "enable_typos_rule": self.typos_rule_var.get(),
            "enable_sensitive_word": self.sensitive_word_var.get()
        }
        
        logger.debug(f"获取校对设置: {settings}")
        return settings
    
    def apply_settings(self) -> bool:
        """
        应用当前设置到配置文件
        
        将所有设置项保存到对应的配置路径。
        
        返回：
            bool: 应用是否成功
        """
        logger.debug("开始应用校对设置到配置文件")
        
        try:
            settings = self.get_settings()
            success = True
            
            # 分别更新不同配置文件
            symbol_settings_keys = ["enable_symbol_pairing", "enable_symbol_correction"]
            typos_settings_keys = ["enable_typos_rule"]
            sensitive_settings_keys = ["enable_sensitive_word"]
            
            for key in symbol_settings_keys:
                if key in settings:
                    result = self.config_manager.update_config_value(
                        "symbol_settings", "engine_settings", key, settings[key]
                    )
                    if not result:
                        success = False
            
            for key in typos_settings_keys:
                if key in settings:
                    result = self.config_manager.update_config_value(
                        "typos_settings", "engine_settings", key, settings[key]
                    )
                    if not result:
                        success = False
            
            for key in sensitive_settings_keys:
                if key in settings:
                    result = self.config_manager.update_config_value(
                        "sensitive_words_settings", "engine_settings", key, settings[key]
                    )
                    if not result:
                        success = False
            
            if success:
                logger.info("✓ 校对设置已成功应用到配置文件")
            else:
                logger.error("✗ 部分校对设置更新失败")
            
            return success
            
        except Exception as e:
            logger.error(f"应用校对设置失败: {e}", exc_info=True)
            return False
