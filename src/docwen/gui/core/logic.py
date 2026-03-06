"""
主窗口逻辑模块 - 业务处理部分
处理核心业务逻辑
包括文件处理、格式转换和交互逻辑

线程安全说明:
- 使用 threading.Lock 保护多线程访问的共享状态
- 使用 threading.Event 进行线程间通信和取消操作
- 所有UI更新通过 root.after 确保在主线程执行
- 文件操作和资源管理使用 try-finally 确保资源释放

日志标准:
- INFO: 用户操作和重要业务逻辑
- DEBUG: 详细的操作步骤和状态变化
- ERROR: 错误信息包含完整上下文
- WARNING: 潜在问题和非致命错误
"""

import contextlib
import logging
import os
import random
import threading
import time
from pathlib import Path
from queue import Empty, Queue
from typing import Any

from docwen.config.config_manager import config_manager
from docwen.errors import DocWenError
from docwen.gui.core.batch_event_adapter import adapt_batch_event_to_queue_messages
from docwen.i18n import t
from docwen.services.cancellation import CancellationToken
from docwen.services.requests import BatchRequest, ConversionRequest
from docwen.services.use_cases import ConversionService
from docwen.utils.gui_utils import show_error_dialog

# 配置日志
logger = logging.getLogger()


class MainWindowLogic:
    """
    主窗口业务逻辑处理类
    处理文件转换、校对等核心业务逻辑
    """

    def __init__(self, main_window):
        """
        初始化逻辑处理器
        """
        logger.info("初始化主窗口逻辑处理器 - 启用线程安全机制")
        self.main_window = main_window

        # 线程安全机制
        self.processing_lock = threading.Lock()  # 保护共享状态的锁
        self.cancel_event = threading.Event()  # 取消事件，用于优雅终止线程
        self.processing_queue = Queue()  # 线程安全队列，用于线程间通信

        # 动画状态
        self.animation_running = False

        # 操作追踪（用于调试和日志）
        self._current_operation_id = None
        self._active_threads = set()  # 跟踪活跃线程

        logger.info("线程安全机制初始化完成")

    def _format_result_message(self, result, _default_message: str | None = None) -> str:
        from docwen.errors import DocWenError

        if result.success:
            return result.message or t("operation_success", default="操作成功完成！")

        if _default_message:
            default_message = _default_message
        else:
            default_message = t("messages.operation_failed", error="", default="操作失败").rstrip()
            if default_message.endswith(":"):
                default_message = default_message[:-1].rstrip()

        if isinstance(result.error, DocWenError):
            return str(result.error)

        if result.details:
            return result.details

        if result.message:
            return result.message

        return default_message

    def handle_file_dropped(self, file_path: str, mode: str = "single"):
        """
        处理文件拖拽事件
        使用TabbedFileManager统一管理UI状态
        """
        logger.debug("┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓")
        logger.debug("┃ 开始处理文件拖拽                                    ┃")
        logger.debug("┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫")
        logger.debug(f"  文件路径: {file_path}")
        logger.debug(f"  拖拽模式: {mode}")

        try:
            # 步骤1: 验证文件存在
            logger.debug("[步骤1] 验证文件是否存在")
            if not Path(file_path).exists():
                logger.error(f"  ✗ 文件不存在: {file_path}")
                self.main_window._show_error(t("messages.file_not_found", path=file_path))
                return
            logger.debug("  ✓ 文件存在")

            # 步骤2: 检查TabbedFileManager
            logger.debug("[步骤2] 检查TabbedFileManager")
            if not hasattr(self.main_window, "tabbed_file_manager"):
                logger.error("  ✗ TabbedFileManager未初始化")
                self.main_window._show_error(t("messages.system_init_error"))
                return
            logger.debug("  ✓ TabbedFileManager已初始化")

            # 单文件模式处理
            if mode == "single":
                logger.debug("[步骤3] 单文件模式处理")
                logger.debug("  [3.1] 单文件模式：直接添加新文件（不清空现有文件）")

                # 3.2 确定文件类别
                logger.debug("  [3.2] 确定文件的实际类别")
                from docwen.utils.file_type_utils import get_file_info

                file_info = get_file_info(file_path, t_func=t)
                category = file_info.get("actual_category", "unknown")
                logger.debug(f"    文件实际类别: {category}")

                # 3.3 先激活对应的选项卡（确保后续操作在正确的选项卡上下文中执行）
                logger.debug(f"  [3.3] 激活 {category} 选项卡")
                if category not in self.main_window.tabbed_batch_file_list.file_lists:
                    logger.error(f"    ✗ 不支持的文件类别: {category}")
                    self.main_window._show_error(t("messages.unsupported_file_type"))
                    return

                self.main_window.tabbed_batch_file_list.set_current_tab(category)
                logger.debug("    ✓ 选项卡已激活")

                # 3.4 添加文件到对应选项卡
                logger.debug(f"  [3.4] 添加文件到 {category} 选项卡")

                # 获取对应的文件列表对象（在激活选项卡后获取，确保是当前激活的列表）
                file_list = self.main_window.tabbed_batch_file_list.file_lists[category]

                # 添加文件（auto_select=True 会自动选中并触发回调）
                success, error_msg = file_list.add_file(file_path, auto_select=True)

                if not success:
                    # 添加失败（通常是类型不匹配等错误）
                    logger.error(f"    ✗ 添加文件失败: {error_msg}")
                    self.main_window._show_error(error_msg)
                    return

                logger.debug("    ✓ 文件已成功处理")

                # 3.5 显示状态消息（UI更新已由on_file_selected完成）
                logger.debug("  [3.5] 显示状态消息")
                self.main_window.tabbed_file_manager.on_files_added([file_path], [])
                logger.debug("    ✓ 状态消息已显示")

            else:
                # 批量模式处理
                logger.debug("[步骤3] 批量模式处理")
                logger.debug("  注意：文件已在FileDropArea的_on_drop方法中添加")

                current_files = self.main_window.tabbed_batch_file_list.get_current_files()
                logger.debug(f"  当前文件列表数量: {len(current_files)}")

                if current_files:
                    logger.debug("  显示状态消息")
                    self.main_window.tabbed_file_manager.on_files_added([file_path], [])
                    logger.debug("  ✓ 批量模式处理完成")
                else:
                    logger.warning("  ⚠ 批量模式下无文件")

            logger.debug("┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛")
            logger.debug("✓ 文件拖拽处理完成\n")

        except Exception as e:
            logger.debug("┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛")
            logger.error(f"✗ 处理文件拖拽时发生异常: {e!s}", exc_info=True)
            self.main_window._show_error(t("messages.process_file_failed", error=str(e)))

    def _handle_md_file(self, file_path: str):
        """处理MD或TXT文件，并根据配置自动选择模板。"""
        logger.info(f"处理MD/TXT文件: {file_path}")
        self.main_window.current_file_path = file_path
        self.main_window.selected_template = None
        self.main_window.add_status_message(t("components.status_bar.selected_text_file"), "secondary", False)
        self.main_window.show_template_selector()
        self.main_window.show_action_panel()
        # 模板选择已由window.py和selector_manager.py处理，此处无需重复
        logger.info("MD/TXT文件处理完成")

    def _handle_docx_file(self, file_path: str):
        """处理DOCX文件"""
        logger.info(f"处理DOCX文件: {file_path}")
        self.main_window.current_file_path = file_path
        self.main_window.selected_template = None
        current_format = Path(file_path).suffix.lstrip(".").lower()
        self.main_window.show_conversion_panel("document", current_format, file_path)
        self.main_window.show_action_panel()
        self.main_window.setup_docx_action_panel(file_path)
        self.main_window.add_status_message(t("components.status_bar.selected_document_file"), "secondary", False)
        logger.info("DOCX文件处理完成")

    def _handle_table_file(self, file_path: str):
        """处理表格文件 (XLSX, CSV)"""
        logger.info(f"处理表格文件: {file_path}")
        self.main_window.current_file_path = file_path
        self.main_window.selected_template = None
        current_format = Path(file_path).suffix.lstrip(".").lower()
        self.main_window.show_conversion_panel("spreadsheet", current_format, file_path)
        file_list = None
        if hasattr(self.main_window, "tabbed_batch_file_list") and self.main_window.tabbed_batch_file_list:
            file_list = self.main_window.tabbed_batch_file_list.get_current_files()
        self.main_window.show_action_panel()
        self.main_window.setup_table_action_panel(file_path, file_list)
        self.main_window.add_status_message(t("components.status_bar.selected_spreadsheet_file"), "secondary", False)
        logger.info("表格文件处理完成")

    def _handle_image_file(self, file_path: str):
        """处理图片文件"""
        logger.info(f"处理图片文件: {file_path}")
        self.main_window.current_file_path = file_path
        self.main_window.selected_template = None
        current_format = Path(file_path).suffix.lstrip(".").lower()
        self.main_window.show_conversion_panel("image", current_format, file_path)
        self.main_window.show_action_panel()
        self.main_window.setup_image_action_panel(file_path)
        self.main_window.add_status_message(t("components.status_bar.selected_image_file"), "secondary", False)
        logger.info("图片文件处理完成")

    def _handle_layout_file(self, file_path: str):
        """处理版式文件"""
        logger.info(f"处理版式文件: {file_path}")
        self.main_window.current_file_path = file_path
        self.main_window.selected_template = None
        current_format = Path(file_path).suffix.lstrip(".").lower()
        self.main_window.show_conversion_panel("layout", current_format, file_path)
        self.main_window.show_action_panel()
        self.main_window.setup_layout_action_panel(file_path)
        self.main_window.add_status_message(t("components.status_bar.selected_layout_file"), "secondary", False)
        logger.info("版式文件处理完成")

    def handle_template_selected(self, template_type: str, template_name: str):
        """处理模板选择事件（简化版 - 事件驱动方案）"""
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        logger.info("📋 主窗口处理模板选择（简化版）")
        logger.info(f"   模板类型: {template_type}")
        logger.info(f"   模板名称: {template_name}")
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        self.main_window.selected_template = (template_type, template_name)
        logger.info("✓ 已更新主窗口的 selected_template 状态")
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        logger.info("✓ 主窗口模板选择处理完成")
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")

    def _try_auto_select_template(self) -> None:
        try:
            selector_tabbed = getattr(self.main_window, "template_selector_tabbed", None)
            if selector_tabbed and hasattr(selector_tabbed, "activate_and_select"):
                selector_tabbed.activate_and_select("docx")
                if hasattr(selector_tabbed, "get_selected_template"):
                    selected = selector_tabbed.get_selected_template()
                    if selected:
                        self.main_window.selected_template = selected
                return

            selector = getattr(self.main_window, "template_selector", None)
            if selector and hasattr(selector, "reset"):
                selector.reset()
                selected_name = selector.get_selected() if hasattr(selector, "get_selected") else None
                if isinstance(selected_name, str) and selected_name:
                    self.main_window.selected_template = ("docx", selected_name)
        except Exception as e:
            logger.debug(f"自动选择模板失败: {e}")

    def handle_cancel(self):
        """处理取消操作的请求。"""
        logger.info("接收到取消请求，开始取消操作")
        with self.processing_lock:
            logger.info("设置取消事件，通知所有后台线程停止处理")
            self.cancel_event.set()
            active_count = len(self._active_threads)
            logger.info(f"取消操作：当前有 {active_count} 个活跃线程将被通知停止")
        logger.info("取消操作处理完成")

    def handle_clear(self):
        """处理清空操作的请求。"""
        logger.info("接收到清空请求，重置界面状态。")
        self.main_window.clear_status_bar()
        self.main_window.current_file_path = None
        self.main_window.selected_template = None
        if hasattr(self.main_window, "tabbed_batch_file_list") and self.main_window.tabbed_batch_file_list:
            self.main_window.tabbed_batch_file_list.clear_all()
            logger.info("已清空选项卡式批量文件列表")
        self.main_window.hide_action_panel()
        self.main_window.hide_template_selector()
        if hasattr(self.main_window, "template_selector") and self.main_window.template_selector:
            self.main_window.template_selector.reset()
        self.main_window.add_status_message(t("components.status_bar.files_cleared"), "secondary", False)
        logger.info("界面状态已彻底重置")

    def _generate_operation_id(self) -> str:
        """生成唯一的操作ID"""
        timestamp = int(time.time())
        random_num = random.randint(1000, 9999)
        operation_id = f"op_{timestamp}_{random_num}"
        logger.debug(f"生成操作ID: {operation_id}")
        return operation_id

    def handle_action(self, action_type: str, file_path: str, options: dict[str, Any]):
        """
        处理操作事件（转换或校对）

        参数:
            action_type: 操作类型 (e.g. 'convert_document_to_pdf', 'validate')
            file_path: 文件路径
            options: 选项
        """
        operation_id = self._generate_operation_id()
        logger.info(f"[{operation_id}] 开始处理操作事件: {action_type}, 文件: {file_path}, 选项: {options}")

        # 特殊处理：合并PDF功能需要文件列表
        if action_type == "merge_pdfs":
            # ... (保持原有逻辑)
            if hasattr(self.main_window, "tabbed_batch_file_list"):
                file_list = self.main_window.tabbed_batch_file_list.get_current_files()
                if len(file_list) < 2:
                    self.main_window._show_error(t("messages.batch_merge_need_two_layout"))
                    return
                options["file_list"] = file_list
                actual_formats = []
                for fp in file_list:
                    from docwen.utils.file_type_utils import detect_actual_file_format

                    fmt = detect_actual_file_format(fp)
                    actual_formats.append(fmt)
                options["actual_formats"] = actual_formats
            else:
                self.main_window._show_error(t("messages.batch_file_list_not_init"))
                return

        # 特殊处理：合并图片为TIFF
        if action_type == "merge_images_to_tiff":
            # ... (保持原有逻辑)
            if hasattr(self.main_window, "tabbed_batch_file_list"):
                file_list = self.main_window.tabbed_batch_file_list.get_current_files()
                if len(file_list) < 2:
                    self.main_window._show_error(t("messages.batch_merge_need_two_image"))
                    return
                options["file_list"] = file_list
            else:
                self.main_window._show_error(t("messages.batch_file_list_not_init"))
                return

        # 特殊处理：拆分PDF
        if action_type == "split_pdf":
            # ... (保持原有逻辑)
            if hasattr(self.main_window, "tabbed_file_manager") and self.main_window.tabbed_file_manager:
                selected_file = self.main_window.tabbed_file_manager.selected_file
                if selected_file:
                    file_path = selected_file.file_path
                    if not options.get("actual_format"):
                        options["actual_format"] = selected_file.actual_format
            options.pop("file_list", None)

        # 特殊处理：汇总表格
        if action_type == "merge_tables":
            # ... (保持原有逻辑)
            if not hasattr(self.main_window, "tabbed_batch_file_list"):
                self.main_window._show_error(t("messages.batch_file_list_not_init"))
                return
            file_list = self.main_window.tabbed_batch_file_list.get_current_files()
            if len(file_list) < 2:
                self.main_window._show_error(t("messages.batch_merge_need_two_table"))
                return
            selected_file = self.main_window.tabbed_batch_file_list.get_selected_file()
            if not selected_file:
                self.main_window._show_error(t("messages.select_base_table_first"))
                return
            base_file = selected_file.file_path
            mode = options.get("mode", 0)
            if mode not in [1, 2, 3]:
                self.main_window._show_error(t("messages.select_merge_mode"))
                return
            options["file_list"] = file_list
            options["mode"] = mode
            file_path = base_file

        # 判断是否批量模式
        current_mode = self.main_window.file_drop_area.get_mode()
        logger.info(f"[{operation_id}] 当前模式: {current_mode}")

        # 聚合类操作（如合并PDF、汇总表格）

        if action_type == "split_pdf":
            logger.info(f"[{operation_id}] 拆分PDF功能：使用单文件模式")
        elif current_mode == "batch" and action_type not in ["merge_tables", "merge_pdfs", "merge_images_to_tiff"]:
            logger.info(f"[{operation_id}] 批量模式：开始批量处理")
            self.handle_batch_processing(action_type, options)
            return

        # 聚合操作专用处理流程
        if action_type == "merge_tables":
            logger.info(f"[{operation_id}] 表格汇总操作：使用专用流程")
            try:
                self.main_window.action_panel.show_cancel_button()
                self._start_processing_animation()
                self.cancel_event.clear()

                # 提取参数
                file_list = options.get("file_list", [])
                mode = options.get("mode", 0)
                base_file = file_path  # handle_action头部已将base_file赋值给file_path

                processing_thread = threading.Thread(
                    target=self._process_merge_tables, args=(base_file, file_list, mode, self.cancel_event)
                )
                processing_thread.start()
                self.main_window.root.after(100, self._check_queue)
                return
            except Exception as e:
                logger.error(f"[{operation_id}] 启动表格汇总失败: {e!s}")
                self.main_window._show_error(t("messages.process_failed", error=str(e)))
                self._stop_processing_animation()
                self.main_window.action_panel.hide_cancel_button()
                return

        if action_type == "merge_pdfs":
            logger.info(f"[{operation_id}] PDF合并操作：使用专用流程")
            try:
                self.main_window.action_panel.show_cancel_button()
                self._start_processing_animation()
                self.cancel_event.clear()

                file_list = options.get("file_list", [])

                processing_thread = threading.Thread(
                    target=self._process_merge_pdfs, args=(file_list, options, self.cancel_event)
                )
                processing_thread.start()
                self.main_window.root.after(100, self._check_queue)
                return
            except Exception as e:
                logger.error(f"[{operation_id}] 启动PDF合并失败: {e!s}")
                self.main_window._show_error(t("messages.process_failed", error=str(e)))
                self._stop_processing_animation()
                self.main_window.action_panel.hide_cancel_button()
                return

        if action_type == "merge_images_to_tiff":
            logger.info(f"[{operation_id}] 图片合并操作：使用专用流程")
            try:
                self.main_window.action_panel.show_cancel_button()
                self._start_processing_animation()
                self.cancel_event.clear()

                file_list = options.get("file_list", [])

                processing_thread = threading.Thread(
                    target=self._process_merge_images_to_tiff, args=(file_list, options, self.cancel_event)
                )
                processing_thread.start()
                self.main_window.root.after(100, self._check_queue)
                return
            except Exception as e:
                logger.error(f"[{operation_id}] 启动图片合并失败: {e!s}")
                self.main_window._show_error(t("messages.process_failed", error=str(e)))
                self._stop_processing_animation()
                self.main_window.action_panel.hide_cancel_button()
                return

        # 单文件模式处理 (或不属于上述聚合操作的其他单线程任务)
        if not Path(file_path).exists():
            error_msg = f"文件 '{Path(file_path).name}' 已被移动或删除，请重新拖入文件。"
            logger.warning(f"[{operation_id}] {error_msg}")
            self.main_window._show_error(error_msg)
            return

        try:
            with self.processing_lock:
                logger.debug(f"[{operation_id}] 重置取消事件")
                self.cancel_event.clear()
                self._current_operation_id = operation_id

            self.main_window.action_panel.show_cancel_button()
            self._start_processing_animation()

            processing_thread = threading.Thread(
                target=self._run_in_background,
                args=(operation_id, action_type, file_path, options, self.cancel_event),
                name=f"ProcessingThread-{operation_id}",
            )
            processing_thread.start()
            logger.info(f"[{operation_id}] 后台处理线程已启动")

            with self.processing_lock:
                if processing_thread.ident:
                    self._active_threads.add(processing_thread.ident)

            self.main_window.root.after(100, self._check_queue)

        except Exception as e:
            logger.error(f"[{operation_id}] 启动处理线程失败: {e!s}")
            self.main_window._show_error(f"处理失败: {e!s}")
            self._stop_processing_animation()
            self.main_window.action_panel.hide_cancel_button()
            self.main_window.action_panel.enable_all_buttons()
            with self.processing_lock:
                self._current_operation_id = None

    def _run_in_background(
        self,
        operation_id: str,
        action_type: str,
        file_path: str,
        options: dict[str, Any],
        cancel_event: threading.Event,
    ):
        """
        在后台线程中运行任务。
        """
        current_thread = threading.current_thread()
        logger.info(f"[{operation_id}] 后台线程开始处理: {action_type}, 文件: {file_path}")

        try:
            if cancel_event.is_set():
                self.processing_queue.put((False, t("conversion.messages.operation_cancelled"), None, action_type))
                return

            # 1. 自动获取文件信息（优先从缓存，否则检测）
            file_category = None
            actual_format = None

            if (
                hasattr(self.main_window, "tabbed_file_manager")
                and self.main_window.tabbed_file_manager
                and self.main_window.tabbed_file_manager.selected_file
            ):
                file_category = self.main_window.tabbed_file_manager.selected_file.actual_category
                actual_format = self.main_window.tabbed_file_manager.selected_file.actual_format

            if not file_category or not actual_format:
                from docwen.utils.file_type_utils import detect_actual_file_format, get_actual_file_category

                if not actual_format:
                    actual_format = detect_actual_file_format(file_path)
                if not file_category:
                    file_category = get_actual_file_category(file_path)

            # 2. 解析目标格式
            # 这里的逻辑是：如果options里有 target_format，就用它
            # 如果没有，且 action_type 是转换类型（convert_A_to_B），则尝试解析
            # 如果 action_type 是命名动作（如 validate），则 target_format 为 None
            raw_target_format = options.get("target_format")
            target_format = raw_target_format if isinstance(raw_target_format, str) else None
            if not target_format and "convert_" in action_type and "_to_" in action_type:
                with contextlib.suppress(Exception):
                    target_format = action_type.split("_to_")[-1]

            logger.debug(
                f"[{operation_id}] 策略解析: action={action_type}, source={actual_format}, target={target_format}"
            )

            # 3. 执行用例层
            # 准备参数
            strategy_options = options.copy()
            strategy_options["cancel_event"] = cancel_event
            strategy_options["actual_format"] = actual_format

            # 对于Markdown文件，注入模板名称
            if file_category == "text" and not strategy_options.get("template_name"):
                if self.main_window.selected_template:
                    strategy_options["template_name"] = self.main_window.selected_template[1]
                else:
                    self.main_window.root.after(0, lambda: self._try_auto_select_template())
                    # 给一点时间让UI线程更新（非阻塞方式可能拿不到，这里主要依赖预先选择）
                    time.sleep(0.1)
                    if self.main_window.selected_template:
                        strategy_options["template_name"] = self.main_window.selected_template[1]

            def thread_safe_progress_update(message):
                self.main_window.set_transient_status("progress", message, "secondary", ttl_ms=15000)

            def progress_callback(message):
                self.main_window.root.after(0, lambda: thread_safe_progress_update(message))

            service = ConversionService()
            cancel_token = CancellationToken(cancel_event)
            req = ConversionRequest(file_path=file_path, action_type=action_type, options=strategy_options)
            result = service.execute(req, progress_callback=progress_callback, cancel_token=cancel_token)

            # 处理结果并放入队列
            if cancel_event.is_set():
                result_tuple = (False, t("conversion.messages.operation_cancelled"), None, action_type)
            else:
                message_to_return = self._format_result_message(result)
                result_tuple = (result.success, message_to_return, result.output_path, action_type)

            logger.info(f"[{operation_id}] 处理完成: {result_tuple[0]}, {result_tuple[1]}")
            self.processing_queue.put(result_tuple)

        except Exception as e:
            logger.error(f"[{operation_id}] 处理异常: {e!s}", exc_info=True)
            if isinstance(e, DocWenError):
                self.processing_queue.put((False, str(e), None, action_type))
            else:
                self.processing_queue.put((False, f"发生意外错误: {e!s}", None, action_type))
        finally:
            with self.processing_lock:
                if current_thread.ident in self._active_threads:
                    self._active_threads.remove(current_thread.ident)

    def _check_queue(self):
        """检查后台线程的结果队列并更新UI。"""
        try:
            while True:
                item = self.processing_queue.get_nowait()

                if isinstance(item, tuple) and len(item) == 4 and isinstance(item[0], bool):
                    success, message, output_path, action_type = item

                    self._stop_processing_animation()
                    self.main_window.action_panel.hide_cancel_button()
                    self.main_window.action_panel.enable_all_buttons()

                    if message == t("conversion.messages.operation_cancelled"):
                        self.main_window.add_status_message(
                            t("components.status_bar.operation_cancelled"), "warning", False
                        )
                    elif success:
                        show_location = not self.main_window.batch_panel_visible
                        self.main_window.add_status_message(
                            message or t("components.status_bar.operation_success"),
                            "success",
                            show_location,
                            output_path,
                        )
                        if output_path and self._should_auto_open_folder():
                            self._open_and_select_file(output_path)

                        if (
                            action_type == "split_pdf"
                            and self.main_window.batch_panel_visible
                            and hasattr(self.main_window, "tabbed_file_manager")
                            and self.main_window.tabbed_file_manager
                        ):
                            selected_file = self.main_window.tabbed_file_manager.selected_file
                            if selected_file and hasattr(self.main_window, "tabbed_batch_file_list"):
                                file_path = selected_file.file_path
                                current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()
                                if current_file_list:
                                    current_file_list.update_file_status(file_path, "completed", output_path)
                    else:
                        show_error_dialog(
                            t("components.status_bar.conversion_failed"),
                            message or t("conversion.messages.conversion_failed_check_log"),
                        )
                        self.main_window.add_status_message(
                            t("components.status_bar.conversion_failed"), "danger", False
                        )
                    return

                match item:
                    case ("set_progress", message):
                        self.main_window.set_transient_status("progress", message, "secondary", ttl_ms=15000)
                        continue
                    case ("set_file_status", file_path, status, output_path, skip_reason, error_message):
                        if hasattr(self.main_window, "tabbed_batch_file_list"):
                            current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()
                            if current_file_list:
                                current_file_list.update_file_status(
                                    file_path,
                                    status,
                                    output_path,
                                    skip_reason=skip_reason,
                                    error_message=error_message,
                                )
                        continue
        except Empty:
            self.main_window.root.after(100, self._check_queue)

    def _start_processing_animation(self):
        """启动状态栏处理动画。"""
        if not self.animation_running:
            self.animation_running = True
            self.main_window.root.after(0, self._update_animation)

    def _stop_processing_animation(self):
        """停止状态栏处理动画。"""
        self.animation_running = False
        self.main_window.clear_transient_status("processing")
        self.main_window.clear_transient_status("progress")

    def _update_animation(self, dot_count=1):
        """更新动画文本。"""
        if not self.animation_running:
            return

        dots = "." * dot_count
        self.main_window.set_transient_status(
            "processing", t("components.status_bar.processing_wait", dots=dots), "secondary"
        )

        next_dot_count = (dot_count % 3) + 1
        self.main_window.root.after(300, self._update_animation, next_dot_count)

    def handle_batch_files_added(self, file_list: list):
        """批量模式下添加多个文件"""
        logger.info(f"批量添加文件: {len(file_list)} 个")
        if not hasattr(self.main_window, "tabbed_batch_file_list"):
            return
        tabbed_list = self.main_window.tabbed_batch_file_list
        current_files = tabbed_list.get_current_files()
        if current_files:
            first_file = current_files[0]
            ext = Path(first_file).suffix.lower()
            if ext in [".md", ".txt"]:
                self._handle_md_file(first_file)
            elif ext == ".docx":
                self._handle_docx_file(first_file)
            elif ext in [".xlsx", ".csv"]:
                self._handle_table_file(first_file)
            total_files = len(current_files)
            message = f"已添加 {total_files} 个文件到批量列表"
            self.main_window.add_status_message(message, "success", False)
        else:
            logger.warning("批量列表为空，无法设置UI")

    def handle_batch_processing(self, action_type: str, options: dict[str, Any]):
        """执行批量处理"""
        logger.info(f"开始批量处理: {action_type}")

        if not hasattr(self.main_window, "tabbed_batch_file_list"):
            return

        tabbed_list = self.main_window.tabbed_batch_file_list
        files = tabbed_list.get_current_files()

        if not files:
            self.main_window._show_error("当前激活选项卡无文件")
            return

        logger.info(f"批量处理文件数量: {len(files)}")

        try:
            self.cancel_event.clear()
            self.main_window.action_panel.show_cancel_button()
            self._start_processing_animation()

            processing_thread = threading.Thread(
                target=self._batch_process_files, args=(files, action_type, options, self.cancel_event)
            )
            processing_thread.start()
            self.main_window.root.after(100, self._check_queue)

        except Exception as e:
            logger.error(f"启动批量处理失败: {e!s}")
            self.main_window._show_error(f"批量处理启动失败: {e!s}")
            self._stop_processing_animation()
            self.main_window.action_panel.hide_cancel_button()
            self.main_window.action_panel.enable_all_buttons()

    def _batch_process_files(
        self, files: list, action_type: str, options: dict[str, Any], cancel_event: threading.Event
    ):
        """后台线程：批量处理文件"""
        try:
            logger.info(f"批量处理线程启动，文件数: {len(files)}")

            if not hasattr(self.main_window, "tabbed_batch_file_list"):
                return

            tabbed_list = self.main_window.tabbed_batch_file_list
            current_file_list = tabbed_list.get_current_file_list()
            if not current_file_list:
                return

            # 解析目标格式（批量处理中目标格式通常是统一的）
            raw_target_format = options.get("target_format")
            target_format = raw_target_format if isinstance(raw_target_format, str) else None
            if not target_format and "convert_" in action_type and "_to_" in action_type:
                with contextlib.suppress(Exception):
                    target_format = action_type.split("_to_")[-1]
            logger.info(f"批量处理目标格式: {target_format}")

            # 1. 预处理：重置所有文件状态为 pending (序号)
            logger.info("重置所有文件状态为 pending")
            for file_path in files:
                self.processing_queue.put(("set_file_status", file_path, "pending", None, None, None))

            # 给一点时间让UI刷新（可选，避免闪烁过快）
            time.sleep(0.1)

            def thread_safe_progress_update(message: str):
                self.processing_queue.put(("set_progress", message))

            def event_sink(event):
                if cancel_event.is_set():
                    return
                for msg in adapt_batch_event_to_queue_messages(event):
                    self.processing_queue.put(msg)

            max_workers = min(4, max(1, (os.cpu_count() or 2)))
            cancel_token = CancellationToken(cancel_event)
            service = ConversionService()
            requests = []
            for file_path in files:
                file_category = None
                actual_format = None
                file_info = current_file_list.get_file_by_path(file_path)
                if file_info:
                    file_category = file_info.actual_category
                    actual_format = file_info.actual_format
                per_file_options = options.copy()
                if target_format and "target_format" not in per_file_options:
                    per_file_options["target_format"] = target_format
                requests.append(
                    ConversionRequest(
                        file_path=file_path,
                        action_type=action_type,
                        target_format=target_format,
                        options=per_file_options,
                        actual_format=actual_format,
                        category=file_category,
                    )
                )

            batch_request = BatchRequest(requests=requests, continue_on_error=True, max_workers=max_workers)
            batch_result = service.execute_batch(
                batch_request,
                progress_callback=thread_safe_progress_update,
                cancel_token=cancel_token,
                event_sink=event_sink,
            )

            if cancel_event.is_set() or batch_result.cancelled:
                self.processing_queue.put((False, t("conversion.messages.operation_cancelled"), None, action_type))
                return

            message = t(
                "components.status_bar.batch_completed",
                success=batch_result.success_count,
                failed=batch_result.failed_count,
            )
            self.processing_queue.put((True, message, None, action_type))

        except Exception as e:
            error_msg = f"批量处理发生异常: {e!s}"
            logger.error(error_msg, exc_info=True)
            self.processing_queue.put((False, error_msg, None, action_type))

    def _process_merge_tables(self, base_file: str, file_list: list, mode: int, cancel_event: threading.Event):
        """后台线程: 处理汇总表格"""
        try:
            if cancel_event.is_set():
                self.processing_queue.put((False, t("conversion.messages.operation_cancelled"), None, "merge_tables"))
                return

            # 获取当前文件列表引用用于更新状态
            current_file_list = None
            if hasattr(self.main_window, "tabbed_batch_file_list"):
                current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()

            # 1. 重置所有文件状态
            if current_file_list:
                for fp in file_list:
                    self.processing_queue.put(("set_file_status", fp, "pending", None, None, None))

            strategy_options = {"file_list": file_list, "mode": mode, "cancel_event": cancel_event}

            def thread_safe_update(msg):
                self.processing_queue.put(("set_progress", msg))

            service = ConversionService()
            cancel_token = CancellationToken(cancel_event)
            req = ConversionRequest(
                file_path=base_file, action_type="merge_tables", file_list=list(file_list), options=strategy_options
            )
            result = service.execute(req, progress_callback=thread_safe_update, cancel_token=cancel_token)

            if result.success:
                # 更新所有文件状态为 completed
                if current_file_list:
                    for fp in file_list:
                        self.processing_queue.put(("set_file_status", fp, "completed", result.output_path, None, None))
                self.processing_queue.put((True, result.message, result.output_path, "merge_tables"))
            else:
                message_to_return = self._format_result_message(result)
                self.processing_queue.put((False, message_to_return, None, "merge_tables"))

        except DocWenError as e:
            self.processing_queue.put((False, str(e), None, "merge_tables"))
        except Exception as e:
            self.processing_queue.put((False, f"汇总异常: {e!s}", None, "merge_tables"))

    def _process_merge_pdfs(self, file_list: list, options: dict[str, Any], cancel_event: threading.Event):
        """后台线程: 处理PDF合并 (新增)"""
        try:
            if cancel_event.is_set():
                self.processing_queue.put((False, t("conversion.messages.operation_cancelled"), None, "merge_pdfs"))
                return

            current_file_list = None
            if hasattr(self.main_window, "tabbed_batch_file_list"):
                current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()

            # 1. 重置状态
            if current_file_list:
                for fp in file_list:
                    self.processing_queue.put(("set_file_status", fp, "pending", None, None, None))

            # 获取当前选中的文件路径（用于确定输出目录）
            selected_file_path = file_list[0]  # 默认使用第一个文件
            if hasattr(self.main_window, "tabbed_batch_file_list"):
                selected_file = self.main_window.tabbed_batch_file_list.get_selected_file()
                if selected_file:
                    selected_file_path = selected_file.file_path
                    logger.info(f"PDF合并：使用选中文件确定输出目录: {Path(selected_file_path).name}")

            strategy_options = options.copy()
            strategy_options["cancel_event"] = cancel_event
            strategy_options["selected_file"] = selected_file_path  # 传递选中文件路径
            strategy_options.setdefault("file_list", list(file_list))

            def thread_safe_update(msg):
                self.processing_queue.put(("set_progress", msg))

            service = ConversionService()
            cancel_token = CancellationToken(cancel_event)
            req = ConversionRequest(
                file_path=selected_file_path,
                action_type="merge_pdfs",
                file_list=list(file_list),
                options=strategy_options,
            )
            result = service.execute(req, progress_callback=thread_safe_update, cancel_token=cancel_token)

            if result.success:
                if current_file_list:
                    for fp in file_list:
                        self.processing_queue.put(("set_file_status", fp, "completed", result.output_path, None, None))
                self.processing_queue.put((True, result.message, result.output_path, "merge_pdfs"))
            else:
                message_to_return = self._format_result_message(result)
                self.processing_queue.put((False, message_to_return, None, "merge_pdfs"))

        except DocWenError as e:
            logger.error(f"PDF合并失败: {e}", exc_info=True)
            self.processing_queue.put((False, str(e), None, "merge_pdfs"))
        except Exception as e:
            logger.error(f"PDF合并异常: {e}", exc_info=True)
            self.processing_queue.put((False, f"PDF合并异常: {e!s}", None, "merge_pdfs"))

    def _process_merge_images_to_tiff(self, file_list: list, options: dict[str, Any], cancel_event: threading.Event):
        """后台线程: 处理图片合并 (新增)"""
        try:
            if cancel_event.is_set():
                self.processing_queue.put(
                    (False, t("conversion.messages.operation_cancelled"), None, "merge_images_to_tiff")
                )
                return

            current_file_list = None
            if hasattr(self.main_window, "tabbed_batch_file_list"):
                current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()

            # 1. 重置状态
            if current_file_list:
                for fp in file_list:
                    self.processing_queue.put(("set_file_status", fp, "pending", None, None, None))

            # 获取当前选中的文件路径（用于确定输出目录）
            selected_file_path = file_list[0]  # 默认使用第一个文件
            if hasattr(self.main_window, "tabbed_batch_file_list"):
                selected_file = self.main_window.tabbed_batch_file_list.get_selected_file()
                if selected_file:
                    selected_file_path = selected_file.file_path
                    logger.info(f"图片合并：使用选中文件确定输出目录: {Path(selected_file_path).name}")

            strategy_options = options.copy()
            strategy_options["cancel_event"] = cancel_event
            strategy_options["selected_file"] = selected_file_path  # 传递选中文件路径
            strategy_options.setdefault("file_list", list(file_list))

            def thread_safe_update(msg):
                self.processing_queue.put(("set_progress", msg))

            service = ConversionService()
            cancel_token = CancellationToken(cancel_event)
            req = ConversionRequest(
                file_path=selected_file_path,
                action_type="merge_images_to_tiff",
                file_list=list(file_list),
                options=strategy_options,
            )
            result = service.execute(req, progress_callback=thread_safe_update, cancel_token=cancel_token)

            if result.success:
                if current_file_list:
                    for fp in file_list:
                        self.processing_queue.put(("set_file_status", fp, "completed", result.output_path, None, None))
                self.processing_queue.put((True, result.message, result.output_path, "merge_images_to_tiff"))
            else:
                message_to_return = self._format_result_message(result)
                self.processing_queue.put((False, message_to_return, None, "merge_images_to_tiff"))

        except DocWenError as e:
            logger.error(f"图片合并失败: {e}", exc_info=True)
            self.processing_queue.put((False, str(e), None, "merge_images_to_tiff"))
        except Exception as e:
            logger.error(f"图片合并异常: {e}", exc_info=True)
            self.processing_queue.put((False, f"图片合并异常: {e!s}", None, "merge_images_to_tiff"))

    def _should_auto_open_folder(self) -> bool:
        """检查是否应该自动打开输出文件夹"""
        try:
            return config_manager.get_auto_open_folder()
        except Exception as e:
            logger.warning(f"读取自动打开文件夹配置失败: {e}")
            return False

    def _open_and_select_file(self, file_path: str):
        """在资源管理器中打开并选中文件"""
        if not file_path or not Path(file_path).exists():
            return

        import subprocess
        import sys

        try:
            if sys.platform == "win32":
                # Windows (使用 normpath 确保路径分隔符正确)
                file_path = os.path.normpath(file_path)
                subprocess.run(["explorer", "/select,", file_path], check=False)
            elif sys.platform == "darwin":
                # macOS
                subprocess.run(["open", "-R", file_path], check=False)
            else:
                # Linux (xdg-open 通常只打开文件夹，不一定选中文件，尝试打开所在文件夹)
                folder_path = str(Path(file_path).parent)
                subprocess.run(["xdg-open", folder_path], check=False)
        except Exception as e:
            logger.error(f"打开文件位置失败: {e}")

    def open_and_select_file(self, file_path: str):
        self._open_and_select_file(file_path)
