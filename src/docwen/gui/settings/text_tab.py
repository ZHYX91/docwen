"""
文本设置选项卡模块

实现设置对话框的文本文件（Markdown）设置选项卡，包含：
- MD转DOCX序号设置
- 序号方案管理
- 模板设置
- 校对设置
- 词库配置

国际化说明：
所有用户可见文本通过 i18n 模块的 t() 函数获取，
支持中英文界面切换。
"""

from __future__ import annotations

import logging
import tkinter as tk
from collections.abc import Callable
from typing import Any

import ttkbootstrap as tb

from docwen.gui.settings.base_tab import BaseSettingsTab
from docwen.gui.settings.config import SectionStyle
from docwen.i18n import t
from docwen.utils.dpi_utils import scale

logger = logging.getLogger(__name__)


class TextTab(BaseSettingsTab):
    """
    文本设置选项卡类

    管理Markdown文件相关的所有配置选项。
    包含MD转DOCX序号设置、序号方案管理、模板设置、校对设置、词库配置。
    """

    def __init__(self, parent, config_manager: Any, on_change: Callable[[str, Any], None]):
        """初始化文本设置选项卡"""
        super().__init__(parent, config_manager, on_change)
        logger.info("文本设置选项卡初始化完成")

    def _get_scheme_mappings(self):
        """
        动态获取序号方案ID和名称的双向映射（支持国际化和语言过滤）

        Returns:
            tuple: (id_to_name, name_to_id) 两个字典
        """
        from docwen.i18n import get_current_locale, t

        id_to_name = {}
        try:
            locale = get_current_locale()
            schemes = self.config_manager.get_heading_schemes()
            settings = self.config_manager.get_numbering_settings()
            order = settings.get("order", list(schemes.keys()))

            for sid in order:
                sconfig = schemes.get(sid)
                if not isinstance(sconfig, dict):
                    continue

                # 检查 locales 过滤
                locales = sconfig.get("locales", ["*"])
                if "*" not in locales and locale not in locales:
                    continue

                # 获取名称（支持 name_key 国际化）
                if "name_key" in sconfig:
                    name_key = sconfig["name_key"]
                    name = t(f"editors.numbering_add.names.{name_key}")
                    if name == f"editors.numbering_add.names.{name_key}":
                        name = sconfig.get("name", sid)
                else:
                    name = sconfig.get("name", sid)

                id_to_name[sid] = name

        except Exception as e:
            logger.warning(f"动态获取序号方案映射失败: {e}")
            # 后备默认值
            id_to_name = {"hierarchical_standard": t("editors.numbering_add.names.hierarchical_standard")}

        name_to_id = {v: k for k, v in id_to_name.items()}
        return id_to_name, name_to_id

    def _create_interface(self):
        """创建选项卡界面"""
        logger.debug("开始创建文本设置选项卡界面")

        self._create_numbering_section()
        self._create_md_to_docx_numbering_section()
        self._create_field_processors_section()
        self._create_template_section()
        self._create_validation_section()
        self._create_dictionary_section()

        logger.debug("文本设置选项卡界面创建完成")

    def _create_field_processors_section(self):
        """创建字段处理器（字段优化）设置区域"""
        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.text.field_processors_section"), SectionStyle.WARNING
        )

        try:
            from docwen.converter.md2docx.field_registry import get_available_processors

            processors = get_available_processors()
        except Exception as e:
            logger.error(f"读取字段处理器列表失败: {e}", exc_info=True)
            desc = tb.Label(frame, text=t("settings.text.field_processors_load_failed"), bootstyle="secondary")
            desc.pack(anchor="w")
            self.bind_label_wraplength(desc, frame, min_wraplength=scale(320))
            return

        if not processors:
            desc = tb.Label(frame, text=t("settings.text.field_processors_empty"), bootstyle="secondary")
            desc.pack(anchor="w")
            self.bind_label_wraplength(desc, frame, min_wraplength=scale(320))
            return

        self._field_processor_vars: dict[str, tk.BooleanVar] = {}

        for proc in processors:
            proc_id = str(proc.get("id") or "")
            if not proc_id:
                continue

            name = str(proc.get("name") or "").strip()
            name_key = str(proc.get("name_key") or "").strip()
            if not name and name_key:
                translated = t(f"field_processors.names.{name_key}")
                name = translated if translated != f"field_processors.names.{name_key}" else name_key
            if not name:
                name = proc_id

            description = str(proc.get("description") or "").strip()
            enabled = bool(proc.get("enabled", True))
            load_error = proc.get("load_error")

            proc_frame = tb.Frame(frame)
            proc_frame.pack(fill="x", pady=(0, scale(5)))

            var = tk.BooleanVar(value=enabled)
            self._field_processor_vars[proc_id] = var

            checkbox = tb.Checkbutton(
                proc_frame,
                text=name,
                variable=var,
                command=lambda pid=proc_id, v=var, n=name: self._on_field_processor_toggled(pid, v, n),
                bootstyle="round-toggle",
            )
            checkbox.pack(side="left")

            if description:
                desc_label = tb.Label(
                    frame,
                    text=description,
                    bootstyle="secondary",
                    font=(self.small_font, self.small_size),
                )
                desc_label.pack(anchor="w", padx=(scale(28), 0), pady=(0, scale(5)))
                self.bind_label_wraplength(desc_label, frame, min_wraplength=scale(320))

            if load_error:
                err_label = tb.Label(
                    frame,
                    text=t("settings.text.field_processors_load_error", error=str(load_error)),
                    bootstyle="danger",
                    font=(self.small_font, self.small_size),
                )
                err_label.pack(anchor="w", padx=(scale(28), 0), pady=(0, scale(5)))
                self.bind_label_wraplength(err_label, frame, min_wraplength=scale(320))

    def _on_field_processor_toggled(self, processor_id: str, var: tk.BooleanVar, display_name: str) -> None:
        enabled = bool(var.get())
        try:
            from docwen.converter.md2docx.field_registry import set_processor_enabled
        except Exception as e:
            logger.error(f"导入字段处理器开关API失败: {e}", exc_info=True)
            var.set(not enabled)
            return

        success = False
        try:
            success = bool(set_processor_enabled(processor_id, enabled))
        except Exception as e:
            logger.error(f"更新字段处理器状态失败: {e}", exc_info=True)

        if success:
            logger.info(f"字段处理器 '{processor_id}' 已{'启用' if enabled else '禁用'}")
            return

        var.set(not enabled)
        try:
            from docwen.gui.components.base_dialog import MessageBox

            MessageBox.showerror(
                t("common.error"),
                t("settings.text.field_processors_toggle_failed", name=display_name),
                parent=self,
            )
        except Exception:
            pass

    def _create_md_to_docx_numbering_section(self):
        """创建MD转DOCX序号设置区域"""
        logger.debug("创建MD转DOCX序号设置区域")

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.text.md_to_docx_section"), SectionStyle.PRIMARY
        )

        # 获取当前配置
        try:
            remove_numbering = self.config_manager.get_md_to_docx_remove_numbering()
            add_numbering = self.config_manager.get_md_to_docx_add_numbering()
            default_scheme = self.config_manager.get_md_to_docx_default_scheme()
        except Exception as e:
            logger.warning(f"读取MD转DOCX序号配置失败: {e}")
            remove_numbering, add_numbering = True, True
            default_scheme = "gongwen_standard"

        # 获取序号方案列表
        try:
            scheme_names = self.config_manager.get_scheme_names()
            if not scheme_names:
                scheme_names = ["公文标准", "层级数字标准", "法律条文标准"]
        except Exception as e:
            logger.warning(f"获取序号方案列表失败: {e}")
            scheme_names = ["公文标准", "层级数字标准", "法律条文标准"]

        # 动态获取方案ID到名称的映射
        scheme_id_to_name, _ = self._get_scheme_mappings()
        default_scheme_name = scheme_id_to_name.get(default_scheme, "公文标准")

        # 默认清除原有Markdown小标题序号
        self.md_to_docx_remove_var = tk.BooleanVar(value=remove_numbering)
        self.create_checkbox_with_info(
            frame,
            t("settings.text.remove_numbering"),
            self.md_to_docx_remove_var,
            t("settings.text.remove_numbering_tooltip"),
            lambda: self.on_change("md_to_docx_remove_numbering", self.md_to_docx_remove_var.get()),
        )

        # 默认新增小标题序号 + 默认序号方案（同一行，左右各50%）
        add_scheme_frame = tb.Frame(frame)
        add_scheme_frame.pack(fill="x", pady=(scale(10), 0))

        # 左侧 - 复选框（占50%）
        left_frame = tb.Frame(add_scheme_frame)
        left_frame.pack(side="left", expand=True, anchor="w")

        self.md_to_docx_add_var = tk.BooleanVar(value=add_numbering)
        add_checkbox = tb.Checkbutton(
            left_frame,
            text=t("settings.text.add_numbering"),
            variable=self.md_to_docx_add_var,
            command=lambda: self.on_change("md_to_docx_add_numbering", self.md_to_docx_add_var.get()),
            bootstyle="round-toggle",
        )
        add_checkbox.pack(side="left")

        # 右侧 - 标签 + 下拉框（占50%）
        right_frame = tb.Frame(add_scheme_frame)
        right_frame.pack(side="left", expand=True, anchor="w")

        scheme_label = tb.Label(
            right_frame, text=t("settings.text.scheme_label"), font=(self.small_font, self.small_size)
        )
        scheme_label.pack(side="left", padx=(0, scale(10)))

        self.md_to_docx_scheme_var = tk.StringVar(value=default_scheme_name)
        self.md_to_docx_scheme_combo = tb.Combobox(
            right_frame, textvariable=self.md_to_docx_scheme_var, values=scheme_names, state="readonly", width=15
        )
        self.md_to_docx_scheme_combo.pack(side="left")
        self.md_to_docx_scheme_combo.bind("<<ComboboxSelected>>", lambda e: self._on_md_to_docx_scheme_changed())

    def _create_numbering_section(self):
        """创建序号设置区域（合并序号方案和清理规则）"""
        logger.debug("创建序号设置区域")

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.text.numbering_settings_section"), SectionStyle.INFO
        )

        # 说明文本
        desc_label = tb.Label(
            frame,
            text=t("settings.text.numbering_settings_desc"),
            bootstyle="secondary",
        )
        desc_label.pack(anchor="w", pady=(0, scale(10)))
        self.bind_label_wraplength(desc_label, frame, min_wraplength=scale(320))

        # 按钮容器
        button_frame = tb.Frame(frame)
        button_frame.pack(fill="x")

        # 两个按钮并排
        tb.Button(
            button_frame,
            text=t("settings.text.edit_numbering_schemes"),
            command=self._on_edit_numbering_schemes,
            bootstyle="success",
            width=18,
        ).pack(side="left", padx=(0, scale(8)))

        tb.Button(
            button_frame,
            text=t("settings.text.edit_patterns"),
            command=self._on_edit_numbering_patterns,
            bootstyle="warning",
            width=18,
        ).pack(side="left")

        logger.debug("序号设置区域创建完成")

    def _create_template_section(self):
        """
        创建模板设置区域

        配置MD文件的默认模板类型选择。

        配置路径：gui_config.template.md_default_template
        """
        logger.debug("创建模板设置区域")

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.text.template_section"), SectionStyle.SUCCESS
        )

        # 获取当前配置
        current_template_type = self.config_manager.get_default_md_template_type()
        current_display = (
            t("settings.text.template_docx") if current_template_type == "docx" else t("settings.text.template_xlsx")
        )

        # 模板类型选项（国际化）
        template_options = [t("settings.text.template_docx"), t("settings.text.template_xlsx")]

        # 模板类型选择
        self.template_type_var = tk.StringVar(value=current_display)
        self.create_combobox_with_info(
            frame,
            t("settings.text.template_label"),
            self.template_type_var,
            template_options,
            t("settings.text.template_tooltip"),
            self._on_template_changed,
        )

        logger.debug("模板设置区域创建完成")

    def _create_validation_section(self):
        """创建校对设置区域"""
        logger.debug("创建校对设置区域")

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.text.validation_section"), SectionStyle.INFO
        )

        # 获取当前配置
        try:
            document_config = self.config_manager.get_document_defaults()
            symbol_pairing = document_config.get("enable_symbol_pairing", True)
            symbol_correction = document_config.get("enable_symbol_correction", True)
            typos_rule = document_config.get("enable_typos_rule", True)
            sensitive_word = document_config.get("enable_sensitive_word", True)
        except Exception as e:
            logger.warning(f"读取校对配置失败: {e}")
            symbol_pairing, symbol_correction = True, True
            typos_rule, sensitive_word = True, True

        # 获取跳过配置
        try:
            skip_code_blocks = self.config_manager.is_skip_code_blocks_enabled()
            skip_quote_blocks = self.config_manager.is_skip_quote_blocks_enabled()
        except Exception as e:
            logger.warning(f"读取跳过配置失败: {e}")
            skip_code_blocks, skip_quote_blocks = True, False

        # 第一行：标点配对 + 符号校对
        row1_frame = tb.Frame(frame)
        row1_frame.pack(fill="x", pady=(0, 5))

        self.symbol_pairing_var = tk.BooleanVar(value=symbol_pairing)
        pairing_checkbox = tb.Checkbutton(
            row1_frame,
            text=t("settings.text.enable_symbol_pairing"),
            variable=self.symbol_pairing_var,
            command=lambda: self.on_change("enable_symbol_pairing", self.symbol_pairing_var.get()),
            bootstyle="round-toggle",
            width=15,
        )
        pairing_checkbox.pack(side="left", expand=True, anchor="w")

        self.symbol_correction_var = tk.BooleanVar(value=symbol_correction)
        correction_checkbox = tb.Checkbutton(
            row1_frame,
            text=t("settings.text.enable_symbol_correction"),
            variable=self.symbol_correction_var,
            command=lambda: self.on_change("enable_symbol_correction", self.symbol_correction_var.get()),
            bootstyle="round-toggle",
            width=15,
        )
        correction_checkbox.pack(side="left", expand=True, anchor="w")

        # 第二行：错别字校对 + 敏感词匹配
        row2_frame = tb.Frame(frame)
        row2_frame.pack(fill="x", pady=(0, scale(10)))

        self.typos_rule_var = tk.BooleanVar(value=typos_rule)
        typos_checkbox = tb.Checkbutton(
            row2_frame,
            text=t("settings.text.enable_typos_rule"),
            variable=self.typos_rule_var,
            command=lambda: self.on_change("enable_typos_rule", self.typos_rule_var.get()),
            bootstyle="round-toggle",
            width=15,
        )
        typos_checkbox.pack(side="left", expand=True, anchor="w")

        self.sensitive_word_var = tk.BooleanVar(value=sensitive_word)
        sensitive_checkbox = tb.Checkbutton(
            row2_frame,
            text=t("settings.text.enable_sensitive_word"),
            variable=self.sensitive_word_var,
            command=lambda: self.on_change("enable_sensitive_word", self.sensitive_word_var.get()),
            bootstyle="round-toggle",
            width=15,
        )
        sensitive_checkbox.pack(side="left", expand=True, anchor="w")

        # 分隔线
        separator = tb.Separator(frame, orient="horizontal")
        separator.pack(fill="x", pady=(5, 10))

        # 跳过选项行
        skip_frame = tb.Frame(frame)
        skip_frame.pack(fill="x")

        self.skip_code_blocks_var = tk.BooleanVar(value=skip_code_blocks)
        skip_code_checkbox = tb.Checkbutton(
            skip_frame,
            text=t("settings.text.skip_code_blocks"),
            variable=self.skip_code_blocks_var,
            command=lambda: self.on_change("skip_code_blocks", self.skip_code_blocks_var.get()),
            bootstyle="round-toggle",
            width=15,
        )
        skip_code_checkbox.pack(side="left", expand=True, anchor="w")

        self.skip_quote_blocks_var = tk.BooleanVar(value=skip_quote_blocks)
        skip_quote_checkbox = tb.Checkbutton(
            skip_frame,
            text=t("settings.text.skip_quote_blocks"),
            variable=self.skip_quote_blocks_var,
            command=lambda: self.on_change("skip_quote_blocks", self.skip_quote_blocks_var.get()),
            bootstyle="round-toggle",
            width=15,
        )
        skip_quote_checkbox.pack(side="left", expand=True, anchor="w")

    def _create_dictionary_section(self):
        """创建词库配置区域"""
        logger.debug("创建词库配置区域")

        frame = self.create_section_frame(
            self.scrollable_frame, t("settings.text.dictionary_section"), SectionStyle.WARNING
        )

        # 说明文本
        desc_label = tb.Label(
            frame,
            text=t("settings.text.dictionary_desc"),
            bootstyle="secondary",
        )
        desc_label.pack(anchor="w", pady=(0, scale(10)))
        self.bind_label_wraplength(desc_label, frame, min_wraplength=scale(320))

        # 按钮容器
        button_frame = tb.Frame(frame)
        button_frame.pack(fill="x")

        # 三个按钮
        tb.Button(
            button_frame,
            text=t("settings.text.edit_symbol_mapping"),
            command=self._on_edit_symbol_mapping,
            bootstyle="info",
            width=16,
        ).pack(side="left", padx=(0, scale(8)))

        tb.Button(
            button_frame,
            text=t("settings.text.edit_typos"),
            command=self._on_edit_custom_typos,
            bootstyle="primary",
            width=16,
        ).pack(side="left", padx=(0, scale(8)))

        tb.Button(
            button_frame,
            text=t("settings.text.edit_sensitive_words"),
            command=self._on_edit_sensitive_words,
            bootstyle="danger",
            width=16,
        ).pack(side="left")

    # ========== 事件处理方法 ==========

    def _on_md_to_docx_scheme_changed(self):
        """处理MD转DOCX序号方案变更"""
        scheme_name = self.md_to_docx_scheme_var.get()

        # 动态获取名称到ID的映射
        _, scheme_name_to_id = self._get_scheme_mappings()

        scheme_id = scheme_name_to_id.get(scheme_name, "gongwen_standard")
        logger.info(f"MD转DOCX序号方案变更: {scheme_name} ({scheme_id})")
        self.on_change("md_to_docx_default_scheme", scheme_id)

    def _on_template_changed(self, event=None):
        """处理模板类型变更"""
        display = self.template_type_var.get()
        docx_display = t("settings.text.template_docx")
        config_value = "docx" if display == docx_display else "xlsx"
        logger.info(f"模板类型变更: {display} → {config_value}")
        self.on_change("md_default_template", config_value)

    # ========== 序号方案管理方法 ==========

    def _on_edit_numbering_schemes(self):
        """打开序号方案编辑器"""
        logger.info("打开序号方案编辑器")

        try:
            from docwen.gui.settings.editors.numbering_add import HeadingNumberingEditorDialog

            # 创建编辑器对话框
            editor = HeadingNumberingEditorDialog(
                parent=self, config_manager=self.config_manager, on_save=self._on_numbering_schemes_saved
            )

            # 等待对话框关闭
            self.wait_window(editor)

        except Exception as e:
            logger.error(f"打开序号方案编辑器失败: {e}", exc_info=True)
            try:
                from docwen.gui.components.base_dialog import MessageBox

                MessageBox.showerror(
                    t("common.error"), t("settings.text.editor_open_failed", error=str(e)), parent=self
                )
            except Exception:
                pass

    def _on_numbering_schemes_saved(self):
        """序号方案保存后的回调"""
        logger.info("序号方案已更新，刷新下拉框")

        try:
            # 刷新序号方案下拉框的选项
            scheme_names = self.config_manager.get_scheme_names()
            if scheme_names and self.md_to_docx_scheme_combo:
                self.md_to_docx_scheme_combo.configure(values=scheme_names)

                # 更新当前选中的默认方案
                default_scheme = self.config_manager.get_md_to_docx_default_scheme()
                # 动态获取方案映射
                scheme_id_to_name, _ = self._get_scheme_mappings()

                default_name = scheme_id_to_name.get(default_scheme, "公文标准")
                self.md_to_docx_scheme_var.set(default_name)

        except Exception as e:
            logger.warning(f"刷新序号方案下拉框失败: {e}")

    def _on_edit_numbering_patterns(self):
        """打开序号清理规则编辑器"""
        logger.info("打开序号清理规则编辑器")

        try:
            from docwen.gui.settings.editors.numbering_clean import NumberingPatternsEditorDialog

            # 创建编辑器对话框
            editor = NumberingPatternsEditorDialog(
                parent=self,
                config_manager=self.config_manager,
                on_save=None,  # 清理规则保存后不需要特殊回调
            )

            # 等待对话框关闭
            self.wait_window(editor)

        except Exception as e:
            logger.error(f"打开序号清理规则编辑器失败: {e}", exc_info=True)
            try:
                from docwen.gui.components.base_dialog import MessageBox

                MessageBox.showerror(
                    t("common.error"), t("settings.text.editor_open_failed", error=str(e)), parent=self
                )
            except Exception:
                pass

    # ========== 词库配置方法 ==========

    def _on_edit_symbol_mapping(self):
        """编辑符号映射"""
        logger.info("打开符号映射编辑器")
        config_file_path = self.config_manager.get_config_file_path("proofread_symbols")
        from docwen.config.toml_operations import read_toml_file

        config_data = read_toml_file(config_file_path)
        # 直接获取 symbol_map（配置文件中没有 proofread_symbols 顶层表）
        symbol_map = config_data.get("symbol_map", {})
        self._open_editor("symbol", symbol_map, self._on_symbol_mapping_saved, config_file_path)

    def _on_edit_custom_typos(self):
        """编辑错别字"""
        logger.info("打开错别字映射编辑器")
        config_file_path = self.config_manager.get_config_file_path("proofread_typos")
        from docwen.config.toml_operations import read_toml_file

        config_data = read_toml_file(config_file_path)
        # 直接获取 typos（配置文件中没有 proofread_typos 顶层表）
        custom_typos = config_data.get("typos", {})
        self._open_editor("typo", custom_typos, self._on_custom_typos_saved, config_file_path)

    def _on_edit_sensitive_words(self):
        """编辑敏感词"""
        logger.info("打开敏感词映射编辑器")
        config_file_path = self.config_manager.get_config_file_path("proofread_sensitive")
        from docwen.config.toml_operations import read_toml_file

        config_data = read_toml_file(config_file_path)
        # 直接获取 sensitive_words（配置文件中没有 proofread_sensitive 顶层表）
        sensitive_words = config_data.get("sensitive_words", {})
        self._open_editor("sensitive", sensitive_words, self._on_sensitive_words_saved, config_file_path)

    def _open_editor(self, editor_type, data, save_callback, config_file_path=None):
        """打开映射编辑器"""
        try:
            from .mapping_editor import MappingEditorDialog

            editor = MappingEditorDialog(self, editor_type, data, save_callback, config_file_path=config_file_path)
            self.wait_window(editor)
        except ImportError as e:
            logger.error(f"导入映射编辑器失败: {e}")

    def _on_symbol_mapping_saved(self, new_mapping: dict[str, list[str]]):
        """保存符号映射"""
        logger.info("符号映射已保存")
        config_file_path = self.config_manager.get_config_file_path("proofread_symbols")
        self._save_mapping_with_comments(config_file_path, "symbol_map", new_mapping)

    def _on_custom_typos_saved(self, new_mapping: dict[str, list[str]]):
        """保存错别字"""
        logger.info("错别字映射已保存")
        config_file_path = self.config_manager.get_config_file_path("proofread_typos")
        self._save_mapping_with_comments(config_file_path, "typos", new_mapping)

    def _on_sensitive_words_saved(self, new_mapping: dict[str, list[str]]):
        """保存敏感词"""
        logger.info("敏感词映射已保存")
        config_file_path = self.config_manager.get_config_file_path("proofread_sensitive")
        self._save_mapping_with_comments(config_file_path, "sensitive_words", new_mapping)

    def _save_mapping_with_comments(self, filepath, section, mapping_data):
        """
        保存映射数据和备注

        参数:
            filepath: 配置文件路径
            section: 节路径（支持多级，如 "proofread_symbols.symbol_map"）
            mapping_data: 要保存的映射数据
        """
        try:
            from docwen.config.toml_operations import extract_inline_comments, save_mapping_with_comments

            comments_data = extract_inline_comments(filepath, section)
            success = save_mapping_with_comments(filepath, section, mapping_data, comments_data)
            if success:
                self.config_manager.reload_configs()
        except Exception as e:
            logger.error(f"保存映射失败: {e}")

    # ========== 配置获取和应用方法 ==========

    def get_settings(self) -> dict[str, Any]:
        """获取当前设置"""
        # 动态获取方案名称到ID的映射
        _, scheme_name_to_id = self._get_scheme_mappings()

        # 模板类型转换
        template_display = self.template_type_var.get()
        docx_display = t("settings.text.template_docx")
        template_config = "docx" if template_display == docx_display else "xlsx"

        settings = {
            # MD转DOCX序号设置
            "to_docx_remove_numbering": self.md_to_docx_remove_var.get(),
            "to_docx_add_numbering": self.md_to_docx_add_var.get(),
            "to_docx_default_scheme": scheme_name_to_id.get(self.md_to_docx_scheme_var.get(), "gongwen_standard"),
            # 模板设置
            "md_default_template": template_config,
            # 校对设置
            "enable_symbol_pairing": self.symbol_pairing_var.get(),
            "enable_symbol_correction": self.symbol_correction_var.get(),
            "enable_typos_rule": self.typos_rule_var.get(),
            "enable_sensitive_word": self.sensitive_word_var.get(),
            # 跳过设置
            "skip_code_blocks": self.skip_code_blocks_var.get(),
            "skip_quote_blocks": self.skip_quote_blocks_var.get(),
        }

        logger.debug(f"获取文本设置: {settings}")
        return settings

    def apply_settings(self) -> bool:
        """应用设置到配置文件"""
        logger.debug("开始应用文本设置")

        try:
            settings = self.get_settings()
            success = True

            # 保存MD转DOCX序号设置到conversion_defaults.toml的[text]节
            text_keys = ["to_docx_remove_numbering", "to_docx_add_numbering", "to_docx_default_scheme"]
            for key in text_keys:
                if not self.config_manager.update_config_value("conversion_defaults", "text", key, settings[key]):
                    success = False
                    logger.error(f"更新配置失败: {key} = {settings[key]}")

            # 保存模板设置到gui_config.toml的[template]节
            if not self.config_manager.update_config_value(
                "gui_config", "template", "md_default_template", settings["md_default_template"]
            ):
                success = False
                logger.error("更新模板配置失败")

            # 保存校对设置到conversion_defaults.toml的[document]节
            validation_keys = [
                "enable_symbol_pairing",
                "enable_symbol_correction",
                "enable_typos_rule",
                "enable_sensitive_word",
            ]
            for key in validation_keys:
                if not self.config_manager.update_config_value("conversion_defaults", "document", key, settings[key]):
                    success = False
                    logger.error(f"更新校对配置失败: {key} = {settings[key]}")

            # 保存跳过设置到proofread_config.toml的[skip]节
            skip_key_map = {
                "skip_code_blocks": "code_blocks",
                "skip_quote_blocks": "quote_blocks",
            }
            for settings_key, config_key in skip_key_map.items():
                if not self.config_manager.update_config_value(
                    "proofread_config", "skip", config_key, settings[settings_key]
                ):
                    success = False
                    logger.error(f"更新跳过配置失败: {config_key} = {settings[settings_key]}")

            if success:
                logger.info("✓ 文本设置已成功应用")
            else:
                logger.error("✗ 部分文本设置更新失败")

            return success

        except Exception as e:
            logger.error(f"应用文本设置失败: {e}", exc_info=True)
            return False
