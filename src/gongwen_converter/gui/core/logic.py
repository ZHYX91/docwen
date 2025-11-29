"""
主窗口逻辑模块 - 业务处理部分
处理公文转换器的核心业务逻辑
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

import os
import logging
import time
import random
from typing import Dict, Optional, Tuple, Any
import threading
from queue import Queue, Empty

from gongwen_converter.converter.formats.office import office_to_docx, office_to_xlsx, OfficeSoftwareNotFoundError
from gongwen_converter.utils.gui_utils import show_error_dialog

# 导入策略注册表获取函数
from gongwen_converter.services.strategies import get_strategy
from gongwen_converter.config.config_manager import config_manager

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
        self.cancel_event = threading.Event()    # 取消事件，用于优雅终止线程
        self.processing_queue = Queue()          # 线程安全队列，用于线程间通信
        
        # 动画状态
        self.animation_running = False
        
        # 操作追踪（用于调试和日志）
        self._current_operation_id = None
        self._active_threads = set()  # 跟踪活跃线程
        
        logger.info("线程安全机制初始化完成")

    
    def handle_file_dropped(self, file_path: str, mode: str = "single"):
        """
        处理文件拖拽事件
        使用TabbedFileManager统一管理UI状态
        """
        logger.info(f"┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓")
        logger.info(f"┃ 开始处理文件拖拽                                    ┃")
        logger.info(f"┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┫")
        logger.info(f"  文件路径: {file_path}")
        logger.info(f"  拖拽模式: {mode}")
        
        try:
            # 步骤1: 验证文件存在
            logger.info(f"[步骤1] 验证文件是否存在")
            if not os.path.exists(file_path):
                logger.error(f"  ✗ 文件不存在: {file_path}")
                self.main_window._show_error(f"文件不存在: {file_path}")
                return
            logger.info(f"  ✓ 文件存在")
            
            # 步骤2: 检查TabbedFileManager
            logger.info(f"[步骤2] 检查TabbedFileManager")
            if not hasattr(self.main_window, 'tabbed_file_manager'):
                logger.error(f"  ✗ TabbedFileManager未初始化")
                self.main_window._show_error("系统初始化错误")
                return
            logger.info(f"  ✓ TabbedFileManager已初始化")
            
            # 单文件模式处理
            if mode == "single":
                logger.info(f"[步骤3] 单文件模式处理")
                logger.info(f"  [3.1] 单文件模式：直接添加新文件（不清空现有文件）")
                
                # 3.2 确定文件类别
                logger.info(f"  [3.2] 确定文件的实际类别")
                from gongwen_converter.utils.file_type_utils import get_actual_file_category
                category = get_actual_file_category(file_path)
                logger.info(f"    文件实际类别: {category}")
                
                # 3.3 先激活对应的选项卡（确保后续操作在正确的选项卡上下文中执行）
                logger.info(f"  [3.3] 激活 {category} 选项卡")
                if category not in self.main_window.tabbed_batch_file_list.file_lists:
                    logger.error(f"    ✗ 不支持的文件类别: {category}")
                    self.main_window._show_error("不支持的文件类型")
                    return
                
                self.main_window.tabbed_batch_file_list.set_current_tab(category)
                logger.info(f"    ✓ 选项卡已激活")
                
                # 3.4 添加文件到对应选项卡
                logger.info(f"  [3.4] 添加文件到 {category} 选项卡")
                
                # 获取对应的文件列表对象（在激活选项卡后获取，确保是当前激活的列表）
                file_list = self.main_window.tabbed_batch_file_list.file_lists[category]
                
                # 添加文件（auto_select=True 会自动选中并触发回调）
                success, error_msg = file_list.add_file(file_path, auto_select=True)
                
                if not success:
                    # 添加失败（通常是类型不匹配等错误）
                    logger.error(f"    ✗ 添加文件失败: {error_msg}")
                    self.main_window._show_error(error_msg)
                    return
                
                logger.info(f"    ✓ 文件已成功处理")
                
                # 3.5 显示状态消息（UI更新已由on_file_selected完成）
                logger.info(f"  [3.5] 显示状态消息")
                self.main_window.tabbed_file_manager.on_files_added([file_path], [])
                logger.info(f"    ✓ 状态消息已显示")
                
            else:
                # 批量模式处理
                logger.info(f"[步骤3] 批量模式处理")
                logger.info(f"  注意：文件已在FileDropArea的_on_drop方法中添加")
                
                current_files = self.main_window.tabbed_batch_file_list.get_current_files()
                logger.info(f"  当前文件列表数量: {len(current_files)}")
                
                if current_files:
                    logger.info(f"  显示状态消息")
                    self.main_window.tabbed_file_manager.on_files_added([file_path], [])
                    logger.info(f"  ✓ 批量模式处理完成")
                else:
                    logger.warning(f"  ⚠ 批量模式下无文件")
            
            logger.info(f"┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛")
            logger.info(f"✓ 文件拖拽处理完成\n")
                
        except Exception as e:
            logger.error(f"┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛")
            logger.error(f"✗ 处理文件拖拽时发生异常: {str(e)}", exc_info=True)
            self.main_window._show_error(f"处理文件失败: {str(e)}")
    
    def _handle_md_file(self, file_path: str):
        """处理MD或TXT文件，并根据配置自动选择模板。"""
        logger.info(f"处理MD/TXT文件: {file_path}")
        self.main_window.current_file_path = file_path
        file_ext = os.path.splitext(file_path)[1].lower()
        self.main_window.selected_template = None
        status_message = "已选择TXT文件" if file_ext == ".txt" else "已选择MD文件"
        self.main_window.add_status_message(status_message, "secondary", False)
        self.main_window.show_template_selector()
        self.main_window.show_action_panel()
        try:
            conversion_config = config_manager.get_conversion_config_block()
            template_config = conversion_config.get("template", {})
            default_template_type = template_config.get("md_default_template", "docx")
        except Exception as e:
            logger.warning(f"读取MD默认模板类型配置失败: {e}，回退到 'docx'")
            default_template_type = "docx"
        if self.main_window.template_selector:
            self.main_window.template_selector.activate_and_select(default_template_type)
        logger.info("MD/TXT文件处理完成，已自动选择模板")
    
    def _handle_docx_file(self, file_path: str):
        """处理DOCX文件"""
        logger.info(f"处理DOCX文件: {file_path}")
        self.main_window.current_file_path = file_path
        self.main_window.selected_template = None
        _, ext = os.path.splitext(file_path)
        current_format = ext.lstrip('.').lower()
        self.main_window.show_conversion_panel('document', current_format, file_path)
        self.main_window.show_action_panel()
        self.main_window.setup_docx_action_panel(file_path)
        self.main_window.add_status_message("已选择DOCX文件", "secondary", False)
        logger.info("DOCX文件处理完成")

    def _handle_table_file(self, file_path: str):
        """处理表格文件 (XLSX, CSV)"""
        logger.info(f"处理表格文件: {file_path}")
        self.main_window.current_file_path = file_path
        self.main_window.selected_template = None
        _, ext = os.path.splitext(file_path)
        current_format = ext.lstrip('.').lower()
        self.main_window.show_conversion_panel('spreadsheet', current_format, file_path)
        file_list = None
        if hasattr(self.main_window, 'tabbed_batch_file_list') and self.main_window.tabbed_batch_file_list:
            file_list = self.main_window.tabbed_batch_file_list.get_current_files()
        self.main_window.show_action_panel()
        self.main_window.setup_table_action_panel(file_path, file_list)
        self.main_window.add_status_message("已选择表格文件", "secondary", False)
        logger.info("表格文件处理完成")
    
    def _handle_image_file(self, file_path: str):
        """处理图片文件"""
        logger.info(f"处理图片文件: {file_path}")
        self.main_window.current_file_path = file_path
        self.main_window.selected_template = None
        _, ext = os.path.splitext(file_path)
        current_format = ext.lstrip('.').lower()
        self.main_window.show_conversion_panel('image', current_format, file_path)
        self.main_window.show_action_panel()
        self.main_window.setup_image_action_panel(file_path)
        self.main_window.add_status_message("已选择图片文件", "secondary", False)
        logger.info("图片文件处理完成")
    
    def _handle_layout_file(self, file_path: str):
        """处理版式文件"""
        logger.info(f"处理版式文件: {file_path}")
        self.main_window.current_file_path = file_path
        self.main_window.selected_template = None
        _, ext = os.path.splitext(file_path)
        current_format = ext.lstrip('.').lower()
        self.main_window.show_conversion_panel('layout', current_format, file_path)
        self.main_window.show_action_panel()
        self.main_window.setup_layout_action_panel(file_path)
        self.main_window.add_status_message("已选择版式文件", "secondary", False)
        logger.info("版式文件处理完成")
    
    def handle_template_selected(self, template_type: str, template_name: str):
        """处理模板选择事件（简化版 - 事件驱动方案）"""
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        logger.info(f"📋 主窗口处理模板选择（简化版）")
        logger.info(f"   模板类型: {template_type}")
        logger.info(f"   模板名称: {template_name}")
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        self.main_window.selected_template = (template_type, template_name)
        logger.info(f"✓ 已更新主窗口的 selected_template 状态")
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        logger.info("✓ 主窗口模板选择处理完成")
        logger.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")

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
        if hasattr(self.main_window, 'tabbed_batch_file_list') and self.main_window.tabbed_batch_file_list:
            self.main_window.tabbed_batch_file_list.clear_all()
            logger.info("已清空选项卡式批量文件列表")
        self.main_window.hide_action_panel()
        self.main_window.hide_template_selector()
        if hasattr(self.main_window, 'template_selector') and self.main_window.template_selector:
            self.main_window.template_selector.reset()
        self.main_window.add_status_message("已清空文件选择", "secondary", False)
        logger.info("界面状态已彻底重置")

    def _generate_operation_id(self) -> str:
        """生成唯一的操作ID"""
        timestamp = int(time.time())
        random_num = random.randint(1000, 9999)
        operation_id = f"op_{timestamp}_{random_num}"
        logger.debug(f"生成操作ID: {operation_id}")
        return operation_id

    def handle_action(self, action_type: str, file_path: str, options: Dict[str, Any]):
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
            if hasattr(self.main_window, 'tabbed_batch_file_list'):
                file_list = self.main_window.tabbed_batch_file_list.get_current_files()
                if len(file_list) < 2:
                    self.main_window._show_error("需要批量模式，且至少2个版式文件才能合并")
                    return
                options['file_list'] = file_list
                actual_formats = []
                for fp in file_list:
                    from gongwen_converter.utils.file_type_utils import detect_actual_file_format
                    fmt = detect_actual_file_format(fp)
                    actual_formats.append(fmt)
                options['actual_formats'] = actual_formats
            else:
                self.main_window._show_error("批量文件列表组件未初始化")
                return
        
        # 特殊处理：合并图片为TIFF
        if action_type == "merge_images_to_tiff":
            # ... (保持原有逻辑)
            if hasattr(self.main_window, 'tabbed_batch_file_list'):
                file_list = self.main_window.tabbed_batch_file_list.get_current_files()
                if len(file_list) < 2:
                    self.main_window._show_error("需要批量模式，且至少2个图片文件才能合并")
                    return
                options['file_list'] = file_list
            else:
                self.main_window._show_error("批量文件列表组件未初始化")
                return
        
        # 特殊处理：拆分PDF
        if action_type == "split_pdf":
            # ... (保持原有逻辑)
            if hasattr(self.main_window, 'tabbed_file_manager') and self.main_window.tabbed_file_manager:
                selected_file = self.main_window.tabbed_file_manager.selected_file
                if selected_file:
                    file_path = selected_file.file_path
                    if not options.get('actual_format'):
                        options['actual_format'] = selected_file.actual_format
            if 'file_list' in options:
                del options['file_list']
        
        # 特殊处理：汇总表格
        if action_type == "merge_tables":
            # ... (保持原有逻辑)
            if not hasattr(self.main_window, 'tabbed_batch_file_list'):
                self.main_window._show_error("批量文件列表组件未初始化")
                return
            file_list = self.main_window.tabbed_batch_file_list.get_current_files()
            if len(file_list) < 2:
                self.main_window._show_error("需要批量模式，且至少2个表格文件才能汇总")
                return
            selected_file = self.main_window.tabbed_batch_file_list.get_selected_file()
            if not selected_file:
                self.main_window._show_error("请先点击选择一个基准表格")
                return
            base_file = selected_file.file_path
            mode = options.get("mode", 0)
            if mode not in [1, 2, 3]:
                self.main_window._show_error("请选择汇总模式（按行/按列/按单元格）")
                return
            options['file_list'] = file_list
            options['mode'] = mode
            file_path = base_file

        # 新增：在每次转换前检查有效期
        from gongwen_converter.security.expiration_check import get_expiration_status, ExpirationStatus
        from gongwen_converter.utils.gui_utils import show_info_dialog

        status_info = get_expiration_status()
        if status_info.status == ExpirationStatus.EXPIRED:
            show_info_dialog(
                "试用期已到期",
                "您仍可以完成本次操作，但此软件的测试版本已过期。\n请联系开发者获取正式版本。",
                alert=True
            )

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
                file_list = options.get('file_list', [])
                mode = options.get('mode', 0)
                base_file = file_path # handle_action头部已将base_file赋值给file_path
                
                processing_thread = threading.Thread(
                    target=self._process_merge_tables,
                    args=(base_file, file_list, mode, self.cancel_event)
                )
                processing_thread.start()
                self.main_window.root.after(100, self._check_queue)
                return
            except Exception as e:
                logger.error(f"[{operation_id}] 启动表格汇总失败: {str(e)}")
                self.main_window._show_error(f"处理失败: {str(e)}")
                self._stop_processing_animation()
                self.main_window.action_panel.hide_cancel_button()
                return

        if action_type == "merge_pdfs":
            logger.info(f"[{operation_id}] PDF合并操作：使用专用流程")
            try:
                self.main_window.action_panel.show_cancel_button()
                self._start_processing_animation()
                self.cancel_event.clear()
                
                file_list = options.get('file_list', [])
                
                processing_thread = threading.Thread(
                    target=self._process_merge_pdfs,
                    args=(file_list, options, self.cancel_event)
                )
                processing_thread.start()
                self.main_window.root.after(100, self._check_queue)
                return
            except Exception as e:
                logger.error(f"[{operation_id}] 启动PDF合并失败: {str(e)}")
                self.main_window._show_error(f"处理失败: {str(e)}")
                self._stop_processing_animation()
                self.main_window.action_panel.hide_cancel_button()
                return

        if action_type == "merge_images_to_tiff":
            logger.info(f"[{operation_id}] 图片合并操作：使用专用流程")
            try:
                self.main_window.action_panel.show_cancel_button()
                self._start_processing_animation()
                self.cancel_event.clear()
                
                file_list = options.get('file_list', [])
                
                processing_thread = threading.Thread(
                    target=self._process_merge_images_to_tiff,
                    args=(file_list, options, self.cancel_event)
                )
                processing_thread.start()
                self.main_window.root.after(100, self._check_queue)
                return
            except Exception as e:
                logger.error(f"[{operation_id}] 启动图片合并失败: {str(e)}")
                self.main_window._show_error(f"处理失败: {str(e)}")
                self._stop_processing_animation()
                self.main_window.action_panel.hide_cancel_button()
                return
        
        # 单文件模式处理 (或不属于上述聚合操作的其他单线程任务)
        if not os.path.exists(file_path):
            error_msg = f"文件 '{os.path.basename(file_path)}' 已被移动或删除，请重新拖入文件。"
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
                name=f"ProcessingThread-{operation_id}"
            )
            processing_thread.start()
            logger.info(f"[{operation_id}] 后台处理线程已启动")
            
            with self.processing_lock:
                if processing_thread.ident:
                    self._active_threads.add(processing_thread.ident)
            
            self.main_window.root.after(100, self._check_queue)

        except Exception as e:
            logger.error(f"[{operation_id}] 启动处理线程失败: {str(e)}")
            self.main_window._show_error(f"处理失败: {str(e)}")
            self._stop_processing_animation()
            self.main_window.action_panel.hide_cancel_button()
            self.main_window.action_panel.enable_all_buttons()
            with self.processing_lock:
                self._current_operation_id = None

    def _run_in_background(self, operation_id: str, action_type: str, file_path: str, 
                          options: Dict[str, Any], cancel_event: threading.Event):
        """
        在后台线程中运行任务。
        """
        current_thread = threading.current_thread()
        logger.info(f"[{operation_id}] 后台线程开始处理: {action_type}, 文件: {file_path}")
        
        try:
            if cancel_event.is_set():
                self.processing_queue.put((False, "操作已取消", None, action_type))
                return
            
            # 1. 自动获取文件信息（优先从缓存，否则检测）
            file_category = None
            actual_format = None
            
            if (hasattr(self.main_window, 'tabbed_file_manager') and 
                self.main_window.tabbed_file_manager and 
                self.main_window.tabbed_file_manager.selected_file):
                file_category = self.main_window.tabbed_file_manager.selected_file.actual_category
                actual_format = self.main_window.tabbed_file_manager.selected_file.actual_format
            
            if not file_category or not actual_format:
                from gongwen_converter.utils.file_type_utils import get_actual_file_category, detect_actual_file_format
                if not actual_format:
                    actual_format = detect_actual_file_format(file_path)
                if not file_category:
                    file_category = get_actual_file_category(file_path)
            
            # 2. 解析目标格式
            # 这里的逻辑是：如果options里有 target_format，就用它
            # 如果没有，且 action_type 是转换类型（convert_A_to_B），则尝试解析
            # 如果 action_type 是命名动作（如 validate），则 target_format 为 None
            target_format = options.get('target_format')
            if not target_format and 'convert_' in action_type and '_to_' in action_type:
                try:
                    target_format = action_type.split('_to_')[-1]
                except:
                    pass
            
            logger.debug(f"[{operation_id}] 策略解析: action={action_type}, source={actual_format}, target={target_format}")

            # 3. 查找策略类 (使用 refactor 后的 get_strategy)
            try:
                strategy_class = get_strategy(
                    action_type=action_type,
                    source_format=actual_format,
                    target_format=target_format
                )
            except ValueError as e:
                logger.error(f"[{operation_id}] 策略查找失败: {e}")
                self.processing_queue.put((False, str(e), None, action_type))
                return

            # 4. 执行策略
            # 准备参数
            strategy_options = options.copy()
            strategy_options['cancel_event'] = cancel_event
            strategy_options['actual_format'] = actual_format
            
            # 对于Markdown文件，注入模板名称
            if file_category == 'text' and not strategy_options.get('template_name'):
                 if self.main_window.selected_template:
                     strategy_options['template_name'] = self.main_window.selected_template[1]
                 else:
                     self.main_window.root.after(0, lambda: self._try_auto_select_template())
                     # 给一点时间让UI线程更新（非阻塞方式可能拿不到，这里主要依赖预先选择）
                     time.sleep(0.1)
                     if self.main_window.selected_template:
                         strategy_options['template_name'] = self.main_window.selected_template[1]
            
            # 对于Markdown文件，添加校对选项
            if file_category == 'text':
                strategy_options['spell_check_options'] = self._convert_options_to_flag(options)
            
            # 对于文档校对，添加spell_check_options
            if file_category == 'document':
                strategy_options['spell_check_options'] = self._convert_options_to_flag(options)

            def thread_safe_progress_update(message):
                self._stop_processing_animation()
                self.main_window.add_status_message(message, "secondary", False)
            
            progress_callback = lambda message: self.main_window.root.after(0, lambda: thread_safe_progress_update(message))

            strategy = strategy_class()
            logger.info(f"[{operation_id}] 执行策略: {strategy.__class__.__name__}")
            
            result = strategy.execute(
                file_path=file_path,
                options=strategy_options,
                progress_callback=progress_callback
            )
            
            # 处理结果并放入队列
            if cancel_event.is_set():
                result_tuple = (False, "操作已取消", None, action_type)
            else:
                message_to_return = result.message or (str(result.error) if result.error else "操作失败")
                result_tuple = (result.success, message_to_return, result.output_path, action_type)
            
            logger.info(f"[{operation_id}] 处理完成: {result_tuple[0]}, {result_tuple[1]}")
            self.processing_queue.put(result_tuple)
            
        except Exception as e:
            logger.error(f"[{operation_id}] 处理异常: {str(e)}", exc_info=True)
            self.processing_queue.put((False, f"发生意外错误: {str(e)}", None, action_type))
        finally:
            with self.processing_lock:
                if current_thread.ident in self._active_threads:
                    self._active_threads.remove(current_thread.ident)

    def _check_queue(self):
        """检查后台线程的结果队列并更新UI。"""
        try:
            success, message, output_path, action_type = self.processing_queue.get_nowait()

            self._stop_processing_animation()
            self.main_window.action_panel.hide_cancel_button()
            self.main_window.action_panel.enable_all_buttons()
            
            if message == "操作已取消":
                self.main_window.add_status_message("操作已取消", "warning", False)
            elif success:
                show_location = not self.main_window.batch_panel_visible
                self.main_window.add_status_message(message or "操作成功完成！", "success", show_location, output_path)
                if output_path:
                    self.main_window.final_output_path = output_path
                    if self._should_auto_open_folder():
                        self._open_and_select_file(output_path)
                
                # 拆分PDF后续处理
                if action_type == "split_pdf" and self.main_window.batch_panel_visible:
                    if hasattr(self.main_window, 'tabbed_file_manager') and self.main_window.tabbed_file_manager:
                        selected_file = self.main_window.tabbed_file_manager.selected_file
                        if selected_file and hasattr(self.main_window, 'tabbed_batch_file_list'):
                            file_path = selected_file.file_path
                            current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()
                            if current_file_list:
                                self.main_window.root.after(0,
                                    lambda: current_file_list.update_file_status(file_path, 'completed', output_path)
                                )
            else:
                show_error_dialog("转换失败", message or "操作失败，请查看日志。")
                self.main_window.add_status_message("转换失败", "danger", False)

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

    def _update_animation(self, dot_count=1):
        """更新动画文本。"""
        if not self.animation_running:
            return
        
        dots = "." * dot_count
        self.main_window.add_status_message(f"处理中，请稍候{dots}", "secondary", False)
        
        next_dot_count = (dot_count % 3) + 1
        self.main_window.root.after(300, self._update_animation, next_dot_count)
    
    def _convert_options_to_flag(self, options: Dict[str, bool]) -> int:
        """将选项字典转换为标志位"""
        flag = 0
        if options.get("symbol_pairing", False): flag |= 1
        if options.get("symbol_correction", False): flag |= 2
        if options.get("typos_rule", False): flag |= 4
        if options.get("sensitive_word", False): flag |= 8
        return flag
    
    def handle_batch_files_added(self, file_list: list):
        """批量模式下添加多个文件"""
        logger.info(f"批量添加文件: {len(file_list)} 个")
        if not hasattr(self.main_window, 'tabbed_batch_file_list'):
            return
        tabbed_list = self.main_window.tabbed_batch_file_list
        current_files = tabbed_list.get_current_files()
        if current_files:
            first_file = current_files[0]
            _, ext = os.path.splitext(first_file)
            ext = ext.lower()
            if ext in ['.md', '.txt']:
                self._handle_md_file(first_file)
            elif ext == '.docx':
                self._handle_docx_file(first_file)
            elif ext in ['.xlsx', '.csv']:
                self._handle_table_file(first_file)
            total_files = len(current_files)
            message = f"已添加 {total_files} 个文件到批量列表"
            self.main_window.add_status_message(message, "success", False)
        else:
            logger.warning("批量列表为空，无法设置UI")
    
    def handle_batch_processing(self, action_type: str, options: Dict[str, Any]):
        """执行批量处理"""
        logger.info(f"开始批量处理: {action_type}")
        
        if not hasattr(self.main_window, 'tabbed_batch_file_list'):
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
                target=self._batch_process_files,
                args=(files, action_type, options, self.cancel_event)
            )
            processing_thread.start()
            self.main_window.root.after(100, self._check_queue)
        
        except Exception as e:
            logger.error(f"启动批量处理失败: {str(e)}")
            self.main_window._show_error(f"批量处理启动失败: {str(e)}")
            self._stop_processing_animation()
            self.main_window.action_panel.hide_cancel_button()
            self.main_window.action_panel.enable_all_buttons()
    
    def _batch_process_files(self, files: list, action_type: str, 
                            options: Dict[str, Any], cancel_event: threading.Event):
        """后台线程：批量处理文件"""
        try:
            logger.info(f"批量处理线程启动，文件数: {len(files)}")
            total = len(files)
            success_count = 0
            failed_count = 0
            
            if not hasattr(self.main_window, 'tabbed_batch_file_list'):
                return
            
            tabbed_list = self.main_window.tabbed_batch_file_list
            current_file_list = tabbed_list.get_current_file_list()
            if not current_file_list:
                return
            
            # 解析目标格式（批量处理中目标格式通常是统一的）
            target_format = options.get('target_format')
            if not target_format and 'convert_' in action_type and '_to_' in action_type:
                try:
                    target_format = action_type.split('_to_')[-1]
                except:
                    pass
            logger.info(f"批量处理目标格式: {target_format}")

            # 1. 预处理：重置所有文件状态为 pending (序号)
            logger.info("重置所有文件状态为 pending")
            for file_path in files:
                self.main_window.root.after(0, 
                    lambda fp=file_path: current_file_list.update_file_status(fp, 'pending')
                )
            
            # 给一点时间让UI刷新（可选，避免闪烁过快）
            time.sleep(0.1)

            for i, file_path in enumerate(files, 1):
                if cancel_event.is_set():
                    self.processing_queue.put((False, "操作已取消", None, action_type))
                    return
                
                self.main_window.root.after(0, 
                    lambda fp=file_path: current_file_list.update_file_status(fp, 'processing')
                )
                
                self.main_window.root.after(0,
                    lambda c=i, t=total, fn=os.path.basename(file_path): 
                        self.main_window.add_status_message(f"正在处理 {c}/{t}: {fn}", "secondary", False)
                )
                
                # 处理单个文件
                result = self._process_single_file_for_batch_new(
                    file_path, action_type, options, cancel_event, target_format
                )
                
                if result[0]: # Success
                    success_count += 1
                    output_path = result[2]
                    message = result[1]
                    if message and message.startswith("已跳过:"):
                        self.main_window.root.after(0,
                            lambda fp=file_path, msg=message: 
                                current_file_list.update_file_status(fp, 'skipped', skip_reason=msg)
                        )
                    else:
                        self.main_window.root.after(0,
                            lambda fp=file_path, op=output_path: current_file_list.update_file_status(fp, 'completed', op)
                        )
                else: # Failed
                    failed_count += 1
                    error_message = result[1]
                    self.main_window.root.after(0,
                        lambda fp=file_path, err=error_message: 
                            current_file_list.update_file_status(fp, 'failed', error_message=err)
                    )
            
            message = f"批量处理完成：成功 {success_count} 个，失败 {failed_count} 个"
            self.processing_queue.put((True, message, None, action_type))
            
        except Exception as e:
            error_msg = f"批量处理发生异常: {str(e)}"
            logger.error(error_msg, exc_info=True)
            self.processing_queue.put((False, error_msg, None, action_type))

    def _process_single_file_for_batch_new(self, file_path: str, action_type: str,
                                       options: Dict[str, Any], 
                                       cancel_event: threading.Event,
                                       target_format: str = None) -> Tuple[bool, str, Optional[str], str]:
        """批量处理单文件（新逻辑）"""
        try:
            # 1. 获取文件信息
            file_category = None
            actual_format = None
            if hasattr(self.main_window, 'tabbed_batch_file_list'):
                current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()
                if current_file_list:
                    file_info = current_file_list.get_file_by_path(file_path)
                    if file_info:
                        file_category = file_info.actual_category
                        actual_format = file_info.actual_format
            
            if not actual_format:
                from gongwen_converter.utils.file_type_utils import detect_actual_file_format
                actual_format = detect_actual_file_format(file_path)
            
            # 2. 判断是否跳过
            if target_format and actual_format:
                # 简单判断：如果相同且不是图片压缩（limit_size），则跳过
                is_same = actual_format.lower() == target_format.lower()
                is_image_compress = (file_category == 'image' and options.get('limit_size'))
                if is_same and not is_image_compress:
                    return (True, f"已跳过: 已是{actual_format.upper()}格式", None, action_type)

            # 3. 查找策略
            strategy_class = get_strategy(
                action_type=action_type,
                source_format=actual_format,
                target_format=target_format
            )
            
            # 4. 准备参数并执行
            strategy_options = options.copy()
            strategy_options['cancel_event'] = cancel_event
            strategy_options['actual_format'] = actual_format
            
            # 注入模板名称（如果是Markdown并且还没选）
            if file_category == 'text' and not strategy_options.get('template_name'):
                 if self.main_window.selected_template:
                     strategy_options['template_name'] = self.main_window.selected_template[1]
                 else:
                     # 尝试自动选择
                     self._try_auto_select_template()
                     if self.main_window.selected_template:
                         strategy_options['template_name'] = self.main_window.selected_template[1]
            
            # 对于Markdown文件，添加校对选项
            if file_category == 'text':
                strategy_options['spell_check_options'] = self._convert_options_to_flag(options)
            
            if file_category == 'document':
                strategy_options['spell_check_options'] = self._convert_options_to_flag(options)

            strategy = strategy_class()
            
            # 进度回调（更新状态栏，注意多线程安全）
            def thread_safe_progress_update(message):
                # 批量模式状态栏更新可以是"正在处理 X/Y: 文件名"
                # 但具体策略内部会发送"转换中..."等消息
                # 这里简单忽略策略内部的详细进度，或者显示在状态栏
                self.main_window.root.after(0, 
                    lambda: self.main_window.add_status_message(message, "secondary", False)
                )
            
            result = strategy.execute(
                file_path=file_path,
                options=strategy_options,
                progress_callback=thread_safe_progress_update
            )
            
            if cancel_event.is_set():
                return (False, "操作已取消", None, action_type)
            
            message_to_return = result.message or (str(result.error) if result.error else "操作失败")
            return (result.success, message_to_return, result.output_path, action_type)

        except Exception as e:
            return (False, f"处理异常: {str(e)}", None, action_type)

    def _process_merge_tables(self, base_file: str, file_list: list, mode: int, cancel_event: threading.Event):
        """后台线程: 处理汇总表格"""
        try:
            if cancel_event.is_set():
                self.processing_queue.put((False, "操作已取消", None, "merge_tables"))
                return
            
            # 获取当前文件列表引用用于更新状态
            current_file_list = None
            if hasattr(self.main_window, 'tabbed_batch_file_list'):
                current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()
            
            # 1. 重置所有文件状态
            if current_file_list:
                for fp in file_list:
                    self.main_window.root.after(0, 
                        lambda path=fp: current_file_list.update_file_status(path, 'pending')
                    )
            
            strategy_class = get_strategy(action_type="merge_tables")
            strategy_options = {
                'file_list': file_list,
                'mode': mode,
                'cancel_event': cancel_event
            }
            
            def thread_safe_update(msg):
                self.main_window.root.after(0, lambda: self.main_window.add_status_message(msg, "secondary", False))
            
            strategy = strategy_class()
            result = strategy.execute(base_file, strategy_options, thread_safe_update)
            
            if result.success:
                # 更新所有文件状态为 completed
                if current_file_list:
                    for fp in file_list:
                        self.main_window.root.after(0,
                            lambda path=fp, output=result.output_path: 
                                current_file_list.update_file_status(path, 'completed', output)
                        )
                self.processing_queue.put((True, result.message, result.output_path, "merge_tables"))
            else:
                self.processing_queue.put((False, result.message, None, "merge_tables"))
        
        except Exception as e:
            self.processing_queue.put((False, f"汇总异常: {str(e)}", None, "merge_tables"))

    def _process_merge_pdfs(self, file_list: list, options: Dict[str, Any], cancel_event: threading.Event):
        """后台线程: 处理PDF合并 (新增)"""
        try:
            if cancel_event.is_set():
                self.processing_queue.put((False, "操作已取消", None, "merge_pdfs"))
                return
            
            current_file_list = None
            if hasattr(self.main_window, 'tabbed_batch_file_list'):
                current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()
            
            # 1. 重置状态
            if current_file_list:
                for fp in file_list:
                    self.main_window.root.after(0, lambda path=fp: current_file_list.update_file_status(path, 'pending'))
            
            # 获取当前选中的文件路径（用于确定输出目录）
            selected_file_path = file_list[0]  # 默认使用第一个文件
            if hasattr(self.main_window, 'tabbed_batch_file_list'):
                selected_file = self.main_window.tabbed_batch_file_list.get_selected_file()
                if selected_file:
                    selected_file_path = selected_file.file_path
                    logger.info(f"PDF合并：使用选中文件确定输出目录: {os.path.basename(selected_file_path)}")
            
            strategy_class = get_strategy(action_type="merge_pdfs")
            strategy_options = options.copy()
            strategy_options['cancel_event'] = cancel_event
            strategy_options['selected_file'] = selected_file_path  # 传递选中文件路径
            
            def thread_safe_update(msg):
                self.main_window.root.after(0, lambda: self.main_window.add_status_message(msg, "secondary", False))
            
            # 使用选中文件路径确定输出目录，file_list保持原顺序用于合并
            strategy = strategy_class()
            result = strategy.execute(selected_file_path, strategy_options, thread_safe_update)
            
            if result.success:
                if current_file_list:
                    for fp in file_list:
                        self.main_window.root.after(0,
                            lambda path=fp, output=result.output_path: 
                                current_file_list.update_file_status(path, 'completed', output)
                        )
                self.processing_queue.put((True, result.message, result.output_path, "merge_pdfs"))
            else:
                self.processing_queue.put((False, result.message, None, "merge_pdfs"))
        
        except Exception as e:
            logger.error(f"PDF合并异常: {e}", exc_info=True)
            self.processing_queue.put((False, f"PDF合并异常: {str(e)}", None, "merge_pdfs"))

    def _process_merge_images_to_tiff(self, file_list: list, options: Dict[str, Any], cancel_event: threading.Event):
        """后台线程: 处理图片合并 (新增)"""
        try:
            if cancel_event.is_set():
                self.processing_queue.put((False, "操作已取消", None, "merge_images_to_tiff"))
                return
            
            current_file_list = None
            if hasattr(self.main_window, 'tabbed_batch_file_list'):
                current_file_list = self.main_window.tabbed_batch_file_list.get_current_file_list()
            
            # 1. 重置状态
            if current_file_list:
                for fp in file_list:
                    self.main_window.root.after(0, lambda path=fp: current_file_list.update_file_status(path, 'pending'))
            
            # 获取当前选中的文件路径（用于确定输出目录）
            selected_file_path = file_list[0]  # 默认使用第一个文件
            if hasattr(self.main_window, 'tabbed_batch_file_list'):
                selected_file = self.main_window.tabbed_batch_file_list.get_selected_file()
                if selected_file:
                    selected_file_path = selected_file.file_path
                    logger.info(f"图片合并：使用选中文件确定输出目录: {os.path.basename(selected_file_path)}")
            
            strategy_class = get_strategy(action_type="merge_images_to_tiff")
            strategy_options = options.copy()
            strategy_options['cancel_event'] = cancel_event
            strategy_options['selected_file'] = selected_file_path  # 传递选中文件路径
            
            def thread_safe_update(msg):
                self.main_window.root.after(0, lambda: self.main_window.add_status_message(msg, "secondary", False))
            
            # 使用选中文件路径确定输出目录，file_list保持原顺序用于合并
            strategy = strategy_class()
            result = strategy.execute(selected_file_path, strategy_options, thread_safe_update)
            
            if result.success:
                if current_file_list:
                    for fp in file_list:
                        self.main_window.root.after(0,
                            lambda path=fp, output=result.output_path: 
                                current_file_list.update_file_status(path, 'completed', output)
                        )
                self.processing_queue.put((True, result.message, result.output_path, "merge_images_to_tiff"))
            else:
                self.processing_queue.put((False, result.message, None, "merge_images_to_tiff"))
        
        except Exception as e:
            logger.error(f"图片合并异常: {e}", exc_info=True)
            self.processing_queue.put((False, f"图片合并异常: {str(e)}", None, "merge_images_to_tiff"))

    def _should_auto_open_folder(self) -> bool:
        """检查是否应该自动打开输出文件夹"""
        try:
            return config_manager.get_auto_open_folder()
        except Exception as e:
            logger.warning(f"读取自动打开文件夹配置失败: {e}")
            return False

    def _open_and_select_file(self, file_path: str):
        """在资源管理器中打开并选中文件"""
        if not file_path or not os.path.exists(file_path):
            return
            
        import sys
        import subprocess
        
        try:
            if sys.platform == 'win32':
                # Windows (使用 normpath 确保路径分隔符正确)
                file_path = os.path.normpath(file_path)
                subprocess.run(['explorer', '/select,', file_path], check=False)
            elif sys.platform == 'darwin':
                # macOS
                subprocess.run(['open', '-R', file_path], check=False)
            else:
                # Linux (xdg-open 通常只打开文件夹，不一定选中文件，尝试打开所在文件夹)
                folder_path = os.path.dirname(file_path)
                subprocess.run(['xdg-open', folder_path], check=False)
        except Exception as e:
            logger.error(f"打开文件位置失败: {e}")
