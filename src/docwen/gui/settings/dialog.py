"""
设置主对话框模块

实现设置对话框的主窗口，包含通用、文本、文档、表格、图片、版式、链接、格式、输出、日志等选项卡。

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
from ttkbootstrap.constants import *

from docwen.gui.components.base_dialog import BaseDialog
from docwen.gui.core.theme_manager import get_theme_manager
from docwen.gui.settings.config import DIALOG_CONFIG
from docwen.i18n import t
from docwen.utils.font_utils import get_small_font

# 配置日志记录器
logger = logging.getLogger()


class SettingsDialog(BaseDialog):
    """
    设置主对话框类
    包含通用、转换、校对、日志，四个选项卡
    """

    def __init__(self, main_window, config_manager: Any, on_apply: Callable[[dict[str, Any]], None] | None):
        """
        初始化设置对话框

        参数:
            main_window: MainWindow实例
            config_manager: 配置管理器实例
            on_apply: 应用设置的回调函数
        """
        # 构建对话框标题：设置 - 应用名称
        dialog_title = f"{t('settings.title')} - {t('common.app_name')}"
        super().__init__(main_window.root, title=dialog_title, modal=True)
        logger.info("初始化设置主对话框")

        self.is_closing = False
        self.modified_settings = {}
        self.initial_theme = main_window.get_current_theme()
        self.tabs = {}
        self.main_window = main_window
        self.config_manager = config_manager
        self.on_apply = on_apply

        # 获取字体配置
        self.small_font, self.small_size = get_small_font()

        # 设置对话框属性
        self._setup_dialog_properties()

        # 创建界面
        self._create_interface()

        # 居中对话框（在界面创建完成后）
        self.center_on_parent()

        # 设置焦点和绑定事件
        self._setup_bindings()

        logger.info("设置主对话框初始化完成")

    def _setup_dialog_properties(self):
        """设置对话框基本属性"""
        logger.debug("设置对话框基本属性")

        width, height = self.scale(DIALOG_CONFIG.default_width), self.scale(DIALOG_CONFIG.default_height)
        min_width, min_height = self.scale(DIALOG_CONFIG.min_width), self.scale(DIALOG_CONFIG.min_height)

        self.geometry(f"{width}x{height}")
        self.minsize(min_width, min_height)
        self.resizable(True, True)

        logger.debug("对话框属性设置完成")

    def _create_interface(self):
        """创建对话框界面"""
        logger.debug("创建对话框界面")

        # 创建主容器（padding需要DPI缩放）
        self.main_container = tb.Frame(self, padding=self.scale(DIALOG_CONFIG.padding))
        self.main_container.pack(fill="both", expand=True)

        # 配置主容器网格
        self.main_container.grid_rowconfigure(0, weight=1)  # 选项卡区域
        self.main_container.grid_rowconfigure(1, weight=0)  # 按钮区域
        self.main_container.grid_columnconfigure(0, weight=1)

        # 创建选项卡区域
        self._create_tab_area()

        # 创建按钮区域
        self._create_button_area()

        logger.debug("对话框界面创建完成")

    def _create_tab_area(self):
        """创建选项卡区域"""
        logger.debug("创建选项卡区域")

        # 创建选项卡控件
        self.notebook = tb.Notebook(self.main_container, bootstyle="primary")
        self.notebook.grid(row=0, column=0, sticky="nsew", pady=(0, self.scale(15)))
        try:
            from docwen.utils.dpi_utils import scale

            style = tb.Style.get_instance()
            tab_padding = (scale(14), scale(8))
            tab_margins = (scale(6), scale(4), scale(6), 0)
            notebook_style = self.notebook.cget("style") or ""

            tab_styles = {"TNotebook.Tab", "primary.TNotebook.Tab"}
            notebook_styles = {"TNotebook", "primary.TNotebook"}
            if notebook_style:
                tab_styles.add(f"{notebook_style}.Tab")
                notebook_styles.add(notebook_style)

            for s in tab_styles:
                style.configure(s, padding=tab_padding)
            for s in notebook_styles:
                style.configure(s, tabmargins=tab_margins)
        except Exception:
            pass

        # 创建选项卡
        self._create_general_tab()
        self._create_text_tab()
        self._create_export_tab()
        self._create_document_tab()
        self._create_spreadsheet_tab()
        self._create_image_tab()
        self._create_layout_tab()
        self._create_link_tab()
        self._create_formatting_tab()
        self._create_output_tab()
        self._create_logging_tab()

        logger.debug("选项卡区域创建完成")

    def _create_general_tab(self):
        """创建通用设置选项卡"""
        logger.debug("创建通用设置选项卡")

        try:
            from .general_tab import GeneralTab

            # 创建通用设置选项卡
            general_tab = GeneralTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("gui_config", key, value),
            )

            # 添加到选项卡控件（使用国际化文本）
            self.notebook.add(general_tab, text=t("settings.tabs.general"))

            # 存储选项卡引用
            self.tabs["general"] = general_tab

            logger.debug("通用设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入通用设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.general"), str(e))
        except Exception as e:
            logger.error(f"创建通用设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.general"), str(e))

    def _create_text_tab(self):
        """创建文本设置选项卡"""
        logger.debug("创建文本设置选项卡")

        try:
            from .text_tab import TextTab

            # 创建文本设置选项卡
            text_tab = TextTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("conversion_defaults", key, value),
            )

            # 添加到选项卡控件（使用国际化文本）
            self.notebook.add(text_tab, text=t("settings.tabs.text"))

            # 存储选项卡引用
            self.tabs["text"] = text_tab

            logger.debug("文本设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入文本设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.text"), str(e))

    def _create_export_tab(self):
        logger.debug("创建导出设置选项卡")

        try:
            from .export_tab import ExportTab

            export_tab = ExportTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("conversion_defaults", key, value),
            )

            self.notebook.add(export_tab, text=t("settings.tabs.export"))
            self.tabs["export"] = export_tab
            logger.debug("导出设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入导出设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.export"), str(e))
        except Exception as e:
            logger.error(f"创建导出设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.export"), str(e))

    def _create_link_tab(self):
        """创建链接设置选项卡"""
        logger.debug("创建链接设置选项卡")

        try:
            from .link_tab import LinkTab

            # 创建链接设置选项卡
            link_tab = LinkTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("link_config", key, value),
            )

            # 添加到选项卡控件（使用国际化文本）
            self.notebook.add(link_tab, text=t("settings.tabs.link"))

            # 存储选项卡引用
            self.tabs["link"] = link_tab

            logger.debug("链接设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入链接设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.link"), str(e))
        except Exception as e:
            logger.error(f"创建链接设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.link"), str(e))

    def _create_image_tab(self):
        """创建图片设置选项卡"""
        logger.debug("创建图片设置选项卡")

        try:
            from .image_tab import ImageTab

            # 创建图片设置选项卡
            image_tab = ImageTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("conversion_defaults", key, value),
            )

            # 添加到选项卡控件（使用国际化文本）
            self.notebook.add(image_tab, text=t("settings.tabs.image"))

            # 存储选项卡引用
            self.tabs["image"] = image_tab

            logger.debug("图片设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入图片设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.image"), str(e))
        except Exception as e:
            logger.error(f"创建图片设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.image"), str(e))

    def _create_output_tab(self):
        """创建输出设置选项卡"""
        logger.debug("创建输出设置选项卡")

        try:
            from .output_tab import OutputTab

            # 创建输出设置选项卡
            output_tab = OutputTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("output_config", key, value),
            )

            # 添加到选项卡控件（使用国际化文本）
            self.notebook.add(output_tab, text=t("settings.tabs.output"))

            # 存储选项卡引用
            self.tabs["output"] = output_tab

            logger.debug("输出设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入输出设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.output"), str(e))
        except Exception as e:
            logger.error(f"创建输出设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.output"), str(e))

    def _create_document_tab(self):
        """创建文档设置选项卡"""
        logger.debug("创建文档设置选项卡")

        try:
            from .document_tab import DocumentTab

            document_tab = DocumentTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("conversion_defaults", key, value),
            )

            self.notebook.add(document_tab, text=t("settings.tabs.document"))
            self.tabs["document"] = document_tab

            logger.debug("文档设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入文档设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.document"), str(e))
        except Exception as e:
            logger.error(f"创建文档设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.document"), str(e))

    def _create_spreadsheet_tab(self):
        """创建表格设置选项卡"""
        logger.debug("创建表格设置选项卡")

        try:
            from .spreadsheet_tab import SpreadsheetTab

            spreadsheet_tab = SpreadsheetTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("conversion_defaults", key, value),
            )

            self.notebook.add(spreadsheet_tab, text=t("settings.tabs.spreadsheet"))
            self.tabs["spreadsheet"] = spreadsheet_tab

            logger.debug("表格设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入表格设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.spreadsheet"), str(e))
        except Exception as e:
            logger.error(f"创建表格设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.spreadsheet"), str(e))

    def _create_layout_tab(self):
        """创建版式设置选项卡"""
        logger.debug("创建版式设置选项卡")

        try:
            from .layout_tab import LayoutTab

            layout_tab = LayoutTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("conversion_defaults", key, value),
            )

            self.notebook.add(layout_tab, text=t("settings.tabs.layout"))
            self.tabs["layout"] = layout_tab

            logger.debug("版式设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入版式设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.layout"), str(e))
        except Exception as e:
            logger.error(f"创建版式设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.layout"), str(e))

    def _create_formatting_tab(self):
        """创建格式设置选项卡"""
        logger.debug("创建格式设置选项卡")

        try:
            from .formatting_tab import FormattingTab

            # 创建格式设置选项卡
            formatting_tab = FormattingTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("conversion_config", key, value),
            )

            # 添加到选项卡控件（使用国际化文本）
            self.notebook.add(formatting_tab, text=t("settings.tabs.formatting"))

            # 存储选项卡引用
            self.tabs["formatting"] = formatting_tab

            logger.debug("格式设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入格式设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.formatting"), str(e))
        except Exception as e:
            logger.error(f"创建格式设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.formatting"), str(e))

    def _create_logging_tab(self):
        """创建日志设置选项卡"""
        logger.debug("创建日志设置选项卡")

        try:
            from .logging_tab import LoggingTab

            # 创建日志设置选项卡
            logging_tab = LoggingTab(
                self.notebook,
                self.config_manager,
                lambda key, value: self._on_setting_changed("logger_config", key, value),
            )

            # 添加到选项卡控件（使用国际化文本）
            self.notebook.add(logging_tab, text=t("settings.tabs.logging"))

            # 存储选项卡引用
            self.tabs["logging"] = logging_tab

            logger.debug("日志设置选项卡创建完成")

        except ImportError as e:
            logger.error(f"导入日志设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.logging"), str(e))
        except Exception as e:
            logger.error(f"创建日志设置选项卡失败: {e!s}")
            self._show_tab_error(t("settings.tabs.logging"), str(e))

    def _show_tab_error(self, tab_name: str, error: str):
        """显示选项卡错误信息"""
        logger.error(f"{tab_name}选项卡错误: {error}")

        # 创建错误框架
        error_frame = tb.Frame(self.notebook, padding=20)

        # 添加错误信息
        error_label = tb.Label(
            error_frame,
            text=t("settings.errors.tab_load_failed_message", tab=tab_name, error=error),
            bootstyle="danger",
            justify="center",
        )
        error_label.pack(expand=True)

        # 添加到选项卡控件
        self.notebook.add(error_frame, text=tab_name)

    def _create_button_area(self):
        """创建按钮区域"""
        logger.debug("创建按钮区域")

        # 按钮框架
        button_frame = tb.Frame(self.main_container)
        button_frame.grid(row=1, column=0, sticky="ew")

        # 配置按钮框架网格
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=0)
        button_frame.grid_columnconfigure(2, weight=0)
        button_frame.grid_columnconfigure(3, weight=0)

        # 空白标签（用于左对齐）
        tb.Label(button_frame).grid(row=0, column=0, sticky="ew")

        # 应用按钮
        apply_button = tb.Button(
            button_frame, text=t("common.apply"), command=self._on_apply, bootstyle="success", width=10
        )
        apply_button.grid(row=0, column=1, padx=(0, DIALOG_CONFIG.button_spacing))

        # 确定按钮（应用并关闭）
        ok_button = tb.Button(button_frame, text=t("common.ok"), command=self._on_ok, bootstyle="primary", width=10)
        ok_button.grid(row=0, column=2, padx=(0, DIALOG_CONFIG.button_spacing))

        # 取消按钮
        cancel_button = tb.Button(
            button_frame, text=t("common.cancel"), command=self._on_cancel, bootstyle="secondary", width=10
        )
        cancel_button.grid(row=0, column=3)

        logger.debug("按钮区域创建完成")

    def _setup_bindings(self):
        """设置事件绑定"""
        logger.debug("设置事件绑定")

        # 绑定回车键到确定按钮
        self.bind("<Return>", lambda e: self._on_ok())

        # 绑定ESC键到取消按钮
        self.bind("<Escape>", lambda e: self._on_cancel())

        # 绑定窗口关闭事件
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

        logger.debug("事件绑定设置完成")

    def _on_setting_changed(self, category: str, key: str, value: Any):
        """
        处理设置变更事件

        参数:
            category: 设置类别 ('gui_config', 'typo_settings', 'logger_config')
            key: 设置键
            value: 设置值
        """
        logger.info(f"设置变更: {category}.{key} = {value}")

        # 更新修改的设置
        if category not in self.modified_settings:
            self.modified_settings[category] = {}

        self.modified_settings[category][key] = value

        # 对于某些设置，可能需要立即应用预览效果
        self._apply_preview_settings(category, key, value)

    def _apply_preview_settings(self, category: str, key: str, value: Any):
        """应用预览设置（立即生效但不保存）"""
        # 如果正在关闭，不执行任何操作
        if self.is_closing:
            return

        logger.debug(f"设置变更记录: {category}.{key} = {value}")

        try:
            # 检查对话框是否仍然存在
            if not self.winfo_exists():
                logger.debug("设置对话框已销毁，跳过预览")
                return

            # 处理通用设置的预览
            if (
                category == "gui_config"
                and key == "default_transparency"
                and self.main_window.is_transparency_enabled()
            ):
                # 主题变更不再预览，只在保存时应用
                # 透明度可以安全预览（不涉及style.theme_use()）
                if hasattr(self.main_window, "root") and self.main_window.root and self.main_window.root.winfo_exists():
                    self.main_window.set_transparency(float(value))
                else:
                    logger.debug("主窗口已销毁，跳过透明度预览")

            logger.debug(f"设置记录完成: {category}.{key}")
        except tk.TclError as e:
            error_msg = str(e)
            if "bad window path name" not in error_msg and "has been destroyed" not in error_msg:
                logger.error(f"预览设置失败 (TclError): {error_msg}")
        except Exception as e:
            logger.error(f"预览设置失败: {e!s}")

    def _on_apply(self):
        """处理应用按钮点击事件"""
        logger.info("=" * 60)
        logger.info("应用按钮被点击")
        logger.info("=" * 60)

        try:
            # 1. 保存所有选项卡的设置到配置文件
            logger.info("步骤1: 保存所有设置到配置文件...")
            success = self._apply_all_settings()
            logger.info(f"步骤1完成: 保存结果 = {success}")

            if success:
                # 2. 应用设置到运行时
                logger.info("步骤2: 应用设置到运行时...")
                logger.debug(f"要应用的设置: {self.modified_settings}")

                if self.on_apply:
                    logger.debug("调用 on_apply 回调函数...")
                    self.on_apply(self.modified_settings)
                    logger.info("步骤2完成: 设置已应用到运行时")
                else:
                    logger.warning("on_apply 回调函数为 None，跳过应用")

                # 3. 显示成功消息
                logger.debug("显示成功消息...")
                self._show_status(t("settings.status.applied"), "success")

                # 4. 清空修改的设置
                logger.debug("清空 modified_settings...")
                self.modified_settings = {}
                logger.info("应用按钮处理完成")
            else:
                # 显示错误消息
                logger.error("保存设置失败")
                self._show_status(t("settings.status.apply_failed"), "danger")

            logger.info("=" * 60)

        except Exception as e:
            logger.error("=" * 60)
            logger.error(f"应用按钮处理过程中发生异常: {type(e).__name__}: {e!s}", exc_info=True)
            self._show_status(f"{t('settings.status.apply_error')}: {e!s}", "danger")
            logger.error("=" * 60)

    def _apply_all_settings(self) -> bool:
        """应用所有选项卡的设置到配置文件"""
        logger.debug("应用所有选项卡的设置")

        success = True

        # 应用通用设置
        if "general" in self.tabs:
            general_success = self.tabs["general"].apply_settings()
            if not general_success:
                success = False
                logger.error("应用通用设置失败")

        # 应用文本设置
        if "text" in self.tabs:
            text_success = self.tabs["text"].apply_settings()
            if not text_success:
                success = False
                logger.error("应用文本设置失败")

        # 应用导出设置
        if "export" in self.tabs:
            export_success = self.tabs["export"].apply_settings()
            if not export_success:
                success = False
                logger.error("应用导出设置失败")

        # 应用链接设置
        if "link" in self.tabs:
            link_success = self.tabs["link"].apply_settings()
            if not link_success:
                success = False
                logger.error("应用链接设置失败")

        # 应用输出设置
        if "output" in self.tabs:
            output_success = self.tabs["output"].apply_settings()
            if not output_success:
                success = False
                logger.error("应用输出设置失败")

        # 应用日志设置
        if "logging" in self.tabs:
            logging_success = self.tabs["logging"].apply_settings()
            if not logging_success:
                success = False
                logger.error("应用日志设置失败")

        # 应用文档设置
        if "document" in self.tabs:
            document_success = self.tabs["document"].apply_settings()
            if not document_success:
                success = False
                logger.error("应用文档设置失败")

        # 应用表格设置
        if "spreadsheet" in self.tabs:
            spreadsheet_success = self.tabs["spreadsheet"].apply_settings()
            if not spreadsheet_success:
                success = False
                logger.error("应用表格设置失败")

        # 应用图片设置
        if "image" in self.tabs:
            image_success = self.tabs["image"].apply_settings()
            if not image_success:
                success = False
                logger.error("应用图片设置失败")

        # 应用版式设置
        if "layout" in self.tabs:
            layout_success = self.tabs["layout"].apply_settings()
            if not layout_success:
                success = False
                logger.error("应用版式设置失败")

        # 应用格式设置
        if "formatting" in self.tabs:
            formatting_success = self.tabs["formatting"].apply_settings()
            if not formatting_success:
                success = False
                logger.error("应用格式设置失败")

        logger.debug(f"所有设置应用结果: {success}")
        return success

    def _on_ok(self):
        """处理确定按钮点击事件"""
        logger.info("=" * 60)
        logger.info("确定按钮被点击")
        logger.info("=" * 60)

        if self.is_closing:
            logger.warning("对话框已在关闭过程中，跳过重复处理")
            return

        logger.debug("设置 is_closing = True")
        self.is_closing = True

        try:
            # 1. 保存所有设置到配置文件
            logger.info("步骤1: 保存所有设置到配置文件...")
            success = self._apply_all_settings()
            logger.info(f"步骤1完成: 保存结果 = {success}")

            # 2. 在对话框关闭前应用设置（与"应用"按钮逻辑一致）
            if success and self.modified_settings:
                logger.info("步骤2: 在对话框关闭前应用设置到运行时...")
                logger.debug(f"要应用的设置: {self.modified_settings}")

                if self.on_apply:
                    logger.debug("调用 on_apply 回调函数...")
                    self.on_apply(self.modified_settings)
                    logger.info("步骤2完成: 设置已应用到运行时")
                else:
                    logger.warning("on_apply 回调函数为 None，跳过应用")
            else:
                if not success:
                    logger.warning("步骤1失败，跳过步骤2")
                if not self.modified_settings:
                    logger.info("没有要应用的设置，跳过步骤2")

            # 3. 关闭对话框
            logger.info("步骤3: 关闭对话框...")
            logger.debug("释放焦点抓取...")
            self.grab_release()
            logger.debug("销毁对话框...")
            self.destroy()
            logger.info("步骤3完成: 对话框已关闭")
            logger.info("=" * 60)

        except Exception as e:
            logger.error("=" * 60)
            logger.error(f"确定按钮处理过程中发生异常: {type(e).__name__}: {e!s}", exc_info=True)
            logger.error("尝试强制关闭对话框...")
            # 即使出错也要关闭对话框
            try:
                self.grab_release()
                self.destroy()
                logger.error("对话框已强制关闭")
            except Exception:
                logger.error("强制关闭对话框也失败", exc_info=True)
            logger.error("=" * 60)

    def _on_cancel(self):
        """处理取消按钮点击事件"""
        logger.info("取消按钮被点击")

        if self.is_closing:
            return
        self.is_closing = True

        # 恢复初始主题（如果有变更）
        theme_manager = get_theme_manager()
        current_theme = theme_manager.get_current_theme()
        if current_theme != self.initial_theme:
            logger.debug(f"恢复初始主题: {current_theme} -> {self.initial_theme}")
            self.main_window.refresh_theme(self.initial_theme, preview_only=False)

        # 关闭对话框
        self.grab_release()
        self.destroy()

    def _show_status(self, message: str, style: str):
        """显示状态消息"""
        logger.debug(f"显示状态消息: {message} ({style})")

        # 创建临时状态标签
        if hasattr(self, "_status_label"):
            self._status_label.destroy()

        self._status_label = tb.Label(
            self.main_container,
            text=message,
            bootstyle=style,
            font=(self.small_font, self.small_size),  # 使用获取的小字体配置
        )
        self._status_label.grid(row=2, column=0, sticky="ew", pady=(5, 0))

        # 3秒后自动消失
        self.after(
            DIALOG_CONFIG.status_display_time,
            lambda: self._status_label.destroy() if hasattr(self, "_status_label") else None,
        )
