"""
配置下拉框组件

提供配置值与显示文本分离的下拉框组件。
内部存储配置值，显示翻译后的文本，避免语言切换时的映射问题。

使用方式：
    from docwen.gui.components.config_combobox import ConfigCombobox

    # 创建下拉框
    combobox = ConfigCombobox(
        parent,
        config_values=["markdown_embed", "wiki_embed"],
        translate_key_prefix="settings.link.image_styles",
        initial_value="wiki_embed"
    )

    # 获取当前选中的配置值
    config_value = combobox.get_config_value()  # 返回 "wiki_embed"

    # 设置配置值
    combobox.set_config_value("markdown_embed")
"""

import logging
import tkinter as tk
from collections.abc import Callable

import ttkbootstrap as tb

from docwen.i18n import t

logger = logging.getLogger(__name__)


class ConfigCombobox(tb.Combobox):
    """
    配置下拉框组件

    实现配置值与显示文本的分离：
    - 内部使用配置值（如 "markdown_embed"）
    - 显示翻译后的文本（如 "Markdown 嵌入显示"）
    - 支持自动根据当前语言更新显示文本

    属性：
        config_values: 配置值列表
        translate_key_prefix: 翻译键前缀（可选，与 translate_keys 二选一）
        translate_keys: 完整翻译键字典（可选，配置值 -> 翻译键）
    """

    def __init__(
        self,
        parent: tk.Widget,
        config_values: list[str],
        translate_key_prefix: str | None = None,
        translate_keys: dict | None = None,
        initial_value: str | None = None,
        on_change: Callable[[str], None] | None = None,
        **kwargs,
    ):
        """
        初始化配置下拉框

        参数：
            parent: 父组件
            config_values: 配置值列表，如 ["markdown_embed", "wiki_embed"]
            translate_key_prefix: 翻译键前缀，如 "settings.link.image_styles"
                                  会自动拼接为 "settings.link.image_styles.markdown_embed"
            translate_keys: 完整翻译键字典，如 {"markdown_embed": "settings.link.image_styles.md_embed"}
                           优先级高于 translate_key_prefix
            initial_value: 初始配置值（默认使用第一个）
            on_change: 值变更回调函数，参数为新的配置值
            **kwargs: 传递给 Combobox 的其他参数
        """
        # 保存配置
        self._config_values = config_values
        self._translate_key_prefix = translate_key_prefix
        self._translate_keys = translate_keys or {}
        self._on_change_callback = on_change

        # 创建内部变量存储配置值
        self._config_var = tk.StringVar()

        # 设置默认的 bootstyle
        if "bootstyle" not in kwargs:
            kwargs["bootstyle"] = "secondary"

        # 设置默认的 state
        if "state" not in kwargs:
            kwargs["state"] = "readonly"

        # 初始化父类
        super().__init__(parent, **kwargs)

        # 设置初始值
        if initial_value and initial_value in config_values:
            self._config_var.set(initial_value)
        elif config_values:
            self._config_var.set(config_values[0])

        # 刷新显示
        self._refresh_display()

        # 绑定选择事件
        self.bind("<<ComboboxSelected>>", self._on_selection_changed)

        logger.debug("ConfigCombobox 初始化完成: config_values=%s, initial=%s", config_values, self._config_var.get())

    def _get_translate_key(self, config_value: str) -> str:
        """
        获取配置值对应的翻译键

        参数：
            config_value: 配置值

        返回：
            str: 翻译键
        """
        # 优先使用完整翻译键字典
        if config_value in self._translate_keys:
            return self._translate_keys[config_value]

        # 使用前缀拼接
        if self._translate_key_prefix:
            return f"{self._translate_key_prefix}.{config_value}"

        # 回退到配置值本身
        return config_value

    def _get_display_text(self, config_value: str) -> str:
        """
        获取配置值对应的显示文本

        参数：
            config_value: 配置值

        返回：
            str: 翻译后的显示文本
        """
        translate_key = self._get_translate_key(config_value)
        return t(translate_key)

    def _refresh_display(self):
        """刷新显示（更新下拉列表和当前选中项的显示文本）"""
        # 更新下拉列表
        display_values = [self._get_display_text(v) for v in self._config_values]
        self.configure(values=display_values)

        # 更新当前选中项的显示
        current_config = self._config_var.get()
        if current_config:
            current_display = self._get_display_text(current_config)
            self.set(current_display)

    def _on_selection_changed(self, event=None):
        """处理选择变更事件"""
        # 获取当前选中的显示文本
        selected_display = self.get()

        # 查找对应的配置值
        for config_value in self._config_values:
            if self._get_display_text(config_value) == selected_display:
                self._config_var.set(config_value)
                logger.debug("ConfigCombobox 选择变更: %s -> %s", selected_display, config_value)

                # 触发回调
                if self._on_change_callback:
                    self._on_change_callback(config_value)
                break

    def get_config_value(self) -> str:
        """
        获取当前选中的配置值

        返回：
            str: 配置值（如 "markdown_embed"）
        """
        return self._config_var.get()

    def set_config_value(self, config_value: str):
        """
        设置配置值

        参数：
            config_value: 配置值
        """
        if config_value not in self._config_values:
            logger.warning("尝试设置无效的配置值: %s", config_value)
            return

        self._config_var.set(config_value)
        self._refresh_display()
        logger.debug("ConfigCombobox 设置配置值: %s", config_value)

    def refresh_translations(self):
        """
        刷新翻译（语言切换后调用）

        重新获取所有配置值的翻译文本并更新显示。
        """
        self._refresh_display()
        logger.debug("ConfigCombobox 翻译已刷新")
