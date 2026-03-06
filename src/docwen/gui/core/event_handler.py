"""
事件处理模块
处理核心业务逻辑
"""

import logging
import time
from typing import Any

from docwen.i18n import t  # 导入翻译函数

logger = logging.getLogger()


class MainWindowEventHandler:
    """
    主窗口事件处理器类
    处理所有用户交互事件，与界面逻辑分离
    """

    def __init__(self, main_window, config_manager: Any):
        """
        初始化事件处理器

        参数:
            main_window: MainWindow实例引用
            config_manager: 配置管理器实例
        """
        self.main_window = main_window
        self.config_manager = config_manager
        # 创建一个唯一的、贯穿生命周期的逻辑处理器实例
        from .logic import MainWindowLogic

        self.logic = MainWindowLogic(self.main_window)

        # 模板选择防重检查
        self.last_template_selection = None
        self.last_template_selection_time = 0
        self.template_selection_delay_ms = 500  # 500毫秒防重延迟

        logger.info("初始化主窗口事件处理器（并创建唯一的逻辑实例）")

    def on_file_dropped(self, file_paths):
        """
        处理文件拖拽事件

        参数:
            file_paths: 拖拽的文件路径列表
        """
        logger.info(f"文件拖拽事件: {file_paths}")

        try:
            # 获取当前模式
            current_mode = self.main_window.file_drop_area.get_mode()
            logger.debug(f"当前模式: {current_mode}")

            paths = file_paths if isinstance(file_paths, list) else [file_paths]

            if current_mode == "batch":
                logger.info(f"批量模式处理 {len(paths)} 个文件")
                if paths:
                    self.logic.handle_batch_files_added(paths)
                    logger.debug(f"批量模式：已添加 {len(paths)} 个文件")
                return

            if not paths:
                return

            file_path = paths[0]
            logger.info(f"单文件模式处理: {file_path}")
            self.logic.handle_file_dropped(file_path, mode=current_mode)
            logger.debug("文件拖拽事件处理完成")
        except Exception as e:
            logger.error(f"处理文件拖拽失败: {e!s}", exc_info=True)
            self.main_window._show_error(t("messages.file_processing_failed", error=str(e)))

    def on_template_selected(self, template_type: str, template_name: str):
        """
        处理模板选择事件

        参数:
            template_type: 模板类型 ('docx' 或 'xlsx')
            template_name: 模板名称
        """
        current_time = time.time() * 1000  # 转换为毫秒
        current_selection = f"{template_type}:{template_name}"

        # 防重检查：检查是否是相同的模板选择且在防重延迟时间内
        time_diff = current_time - self.last_template_selection_time
        same_selection = current_selection == self.last_template_selection

        if time_diff < self.template_selection_delay_ms and same_selection:
            logger.debug(f"事件处理器阻止重复模板选择: {current_selection} (时间差: {time_diff:.0f}ms)")
            return

        # 更新最后选择记录
        self.last_template_selection = current_selection
        self.last_template_selection_time = current_time

        logger.info(f"模板选择: {template_type}/{template_name}")

        try:
            # 委托给唯一的逻辑实例处理
            self.logic.handle_template_selected(template_type, template_name)
            logger.debug("模板选择事件处理完成")
        except Exception as e:
            logger.error(f"处理模板选择失败: {e!s}", exc_info=True)
            self.main_window._show_error(t("messages.template_selection_failed", error=str(e)))

    def on_action(self, action_type: str, file_path: str, options: dict[str, bool]):
        """
        处理操作事件（转换或校对）

        参数:
            action_type: 操作类型 ('convert_md_to_docx', 'convert_md_to_doc', 'convert' 或 'validate')
            file_path: 文件路径
            options: 校对选项
        """
        logger.info(f"操作事件: {action_type}, 文件: {file_path}, 选项: {options}")

        try:
            # 委托给唯一的逻辑实例处理
            self.logic.handle_action(action_type, file_path, options)
            logger.debug("操作事件处理完成")
        except Exception as e:
            logger.error(f"处理操作事件失败: {e!s}", exc_info=True)
            self.main_window._show_error(t("messages.operation_failed", error=str(e)))

    def on_cancel(self):
        """处理取消按钮点击事件"""
        logger.info("取消事件触发")
        try:
            self.logic.handle_cancel()
        except Exception as e:
            logger.error(f"处理取消事件失败: {e}", exc_info=True)
            self.main_window._show_error(t("messages.cancel_failed", error=str(e)))

    def on_clear_clicked(self):
        """处理清空按钮点击事件"""
        logger.info("清空按钮被点击")
        try:
            self.logic.handle_clear()
        except Exception as e:
            logger.error(f"处理清空事件失败: {e}", exc_info=True)
            self.main_window._show_error(t("messages.clear_failed", error=str(e)))

    def on_settings_clicked(self):
        """
        处理设置按钮点击事件
        """
        logger.info("设置按钮被点击")

        try:
            # 导入新的设置对话框
            from docwen.gui.settings.dialog import SettingsDialog

            # 创建设置对话框 - 使用新的构造函数格式
            settings_dialog = SettingsDialog(
                self.main_window,  # 传递MainWindow实例
                self.config_manager,
                self.on_settings_applied,
            )

            # 设置为模态对话框
            settings_dialog.grab_set()
            settings_dialog.focus_set()

            # 等待对话框关闭
            self.main_window.root.wait_window(settings_dialog)

            # 检查是否需要恢复主题（在对话框完全销毁后）
            theme_to_restore = getattr(settings_dialog, "theme_to_restore", None)
            if theme_to_restore:
                logger.info(f"设置对话框已取消，正在将主题恢复为 '{theme_to_restore}'")
                self.main_window.refresh_theme(theme_to_restore)

        except ImportError as e:
            logger.error(f"导入设置对话框失败: {e!s}")
            self.main_window._show_error(t("messages.settings_unavailable"))
        except Exception as e:
            logger.error(f"打开设置对话框失败: {e!s}")
            self.main_window._show_error(t("messages.open_settings_failed", error=str(e)))

    def on_settings_applied(self, settings: dict[str, Any]):
        """
        处理设置应用事件

        参数:
            settings: 应用的新设置
        """
        logger.info(f"应用新设置: {settings}")

        try:
            # 处理通用设置变更
            if "gui_config" in settings:
                self._apply_gui_settings(settings["gui_config"])

            # 处理校对设置变更
            if "typo_settings" in settings:
                self._apply_typo_settings(settings["typo_settings"])

            # 处理日志设置变更
            if "logger_config" in settings:
                self._apply_logger_settings(settings["logger_config"])

            # 更新状态
            self.main_window.add_status_message("设置已应用", "success")

        except Exception as e:
            logger.error(f"应用设置失败: {e!s}")
            self.main_window._show_error(t("messages.apply_settings_failed", error=str(e)))

    def _apply_gui_settings(self, settings: dict[str, Any]):
        """应用GUI设置变更"""
        logger.info(f"应用GUI设置: {settings}")

        # 应用主题变更
        if "default_theme" in settings:
            self.main_window.refresh_theme(settings["default_theme"])

        # 应用透明度变更
        if "transparency_enabled" in settings:
            self.main_window.transparency_enabled = settings["transparency_enabled"]
            if not self.main_window.transparency_enabled:
                self.main_window.set_transparency(1.0)

        if "default_transparency" in settings and self.main_window.transparency_enabled:
            self.main_window.set_transparency(float(settings["default_transparency"]))

        # 应用其他设置
        if "remember_gui_state" in settings:
            self.main_window.remember_gui_state = settings["remember_gui_state"]

        if "auto_center" in settings:
            self.main_window.auto_center = settings["auto_center"]

        logger.info("GUI设置应用完成")

    def _apply_typo_settings(self, settings: dict[str, Any]):
        """应用校对设置变更"""
        logger.info(f"应用校对设置: {settings}")

        # 这里可以添加对校对设置变更的处理
        # 例如重新加载错别字校对配置等

        # 刷新主界面的操作面板选项状态
        try:
            if hasattr(self.main_window, "action_panel") and self.main_window.action_panel:
                # 重新加载选项设置
                self.main_window.action_panel.refresh_options()
        except Exception as e:
            logger.error(f"刷新操作面板选项失败: {e!s}")

        logger.info("校对设置应用完成")

    def _apply_logger_settings(self, settings: dict[str, Any]):
        """应用日志设置变更"""
        logger.info(f"应用日志设置: {settings}")

        # 这里可以添加对日志设置变更的处理
        # 例如重新配置日志系统等

        logger.info("日志设置应用完成")

    def on_about_clicked(self):
        """
        处理关于按钮点击事件 - 使用独立的关于对话框模块
        """
        logger.info("关于按钮被点击")

        try:
            # 导入并使用独立的关于对话框模块
            from ..components.about_dialog import AboutDialog

            about_dialog = AboutDialog(self.main_window.root)
            about_dialog.show()

        except ImportError as e:
            logger.error(f"导入关于对话框模块失败: {e!s}")
            self.main_window._show_error(t("messages.about_unavailable"))
        except Exception as e:
            logger.error(f"打开关于对话框失败: {e!s}")
            self.main_window._show_error(t("messages.open_about_failed", error=str(e)))

    def on_close(self):
        """处理窗口关闭事件"""
        logger.info("主窗口关闭")

        # 销毁窗口
        self.main_window.root.destroy()
