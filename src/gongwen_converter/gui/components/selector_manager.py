"""
选项卡式文件管理器
提供统一的UI状态管理和选项卡切换处理
"""

import logging
import tkinter as tk
from typing import Dict, List, Optional, Callable, Any
from gongwen_converter.utils.file_type_utils import get_file_category, get_category_name, get_actual_file_category
from gongwen_converter.gui.components.file_selector import FileInfo

# 导入ttkbootstrap用于界面美化和样式管理
import ttkbootstrap as tb
from ttkbootstrap.constants import *

logger = logging.getLogger(__name__)


class SelectorManager:
    """
    选项卡式文件管理器
    统一管理选项卡切换、UI状态更新和文件分类
    """
    
    def __init__(self, 
                 tabbed_file_list: Any,
                 action_panel: Any,
                 template_selector: Any,
                 file_drop_area: Any,
                 status_bar: Any,
                 main_window: Any = None,
                 config_manager: Any = None):
        """
        初始化选项卡式文件管理器
        
        参数:
            tabbed_file_list: 选项卡式文件列表组件
            action_panel: 操作面板组件
            template_selector: 模板选择器组件
            file_drop_area: 文件拖拽区域组件
            status_bar: 状态栏组件
            main_window: 主窗口引用（用于显示/隐藏面板框架）
            config_manager: 配置管理器实例
        """
        logger.debug("初始化选项卡式文件管理器")
        
        # 存储组件引用
        self.tabbed_file_list = tabbed_file_list
        self.action_panel = action_panel
        self.template_selector = template_selector
        self.file_drop_area = file_drop_area
        self.status_bar = status_bar
        self.main_window = main_window  # 主窗口引用，用于控制面板框架的显示/隐藏
        self.config_manager = config_manager  # 配置管理器引用
        
        # 当前状态
        self.current_category: Optional[str] = None
        self.current_files: List[str] = []
        self.current_mode: str = "single"  # "single" 或 "batch"
        self.selected_file: Optional[FileInfo] = None  # 当前选中的文件
        
        # 绑定所有文件列表的选中回调
        self._bind_file_list_callbacks()
        
        # 绑定模板选择器的回调
        self._bind_template_selector_callback()
        
        logger.info("选项卡式文件管理器初始化完成")
    
    def _bind_file_list_callbacks(self):
        """绑定所有文件列表的回调函数"""
        logger.debug("绑定文件列表回调函数")
        
        # 绑定文件选中回调
        if hasattr(self.tabbed_file_list, 'on_file_selected'):
            if self.tabbed_file_list.on_file_selected != self.on_file_selected:
                logger.debug("绑定文件选中回调")
                self.tabbed_file_list.on_file_selected = self.on_file_selected
            else:
                logger.debug("文件选中回调已正确设置")
        else:
            logger.warning("tabbed_file_list没有on_file_selected属性")
        
        # 绑定文件移除回调
        if hasattr(self.tabbed_file_list, 'on_file_removed'):
            self.tabbed_file_list.on_file_removed = self.on_files_removed
            logger.debug("已绑定文件移除回调")
        else:
            logger.warning("tabbed_file_list没有on_file_removed属性")
    
    def _bind_template_selector_callback(self):
        """
        绑定模板选择器的回调函数
        当用户切换模板选项卡或选择模板时，模板选择器会触发回调
        从而动态更新操作面板
        """
        logger.debug("绑定模板选择器回调函数")
        
        # 检查模板选择器是否存在
        if not self.template_selector:
            logger.warning("模板选择器不存在，无法绑定回调")
            return
        
        # 尝试多种绑定方式，确保兼容性
        if hasattr(self.template_selector, 'set_selection_callback'):
            logger.debug("使用 set_selection_callback 方法绑定")
            self.template_selector.set_selection_callback(self.on_template_selected)
            logger.info("模板选择回调已通过 set_selection_callback 绑定")
        
        elif hasattr(self.template_selector, 'on_template_selected'):
            logger.debug("直接设置 on_template_selected 属性")
            self.template_selector.on_template_selected = self.on_template_selected
            logger.info("模板选择回调已通过属性赋值绑定")
        
        else:
            logger.warning("模板选择器不支持回调绑定，模板切换功能可能受限")
    
    def on_tab_changed(self, new_category: str, old_category: Optional[str] = None):
        """
        处理选项卡切换事件
        
        参数:
            new_category: 新的选项卡类别
            old_category: 旧的选项卡类别
        """
        logger.info(f"选项卡切换: {old_category} -> {new_category}")
        
        # 更新当前状态
        self.current_category = new_category
        
        # 只在真正切换到不同选项卡时才清空选中文件
        if old_category and old_category != new_category:
            if self.selected_file:
                logger.debug(f"清空选中文件: {self.selected_file.file_path} (因为切换了选项卡)")
            self.selected_file = None
        else:
            if old_category == new_category:
                logger.debug(f"同选项卡切换({new_category}→{new_category})，保持选中文件状态")
        
        # 统一刷新UI
        self.refresh_ui(reason="tab_changed")
        
        # 批量模式下，通知拖拽区更新显示
        if self.current_mode == "batch" and self.file_drop_area:
            logger.debug("批量模式：通知拖拽区更新当前选项卡信息")
            self.file_drop_area.update_display(None)
        
        logger.debug(f"选项卡切换处理完成: {new_category}")
    
    def on_files_added(self, added_files: List[str], failed_files: List[tuple]):
        """
        处理文件添加事件
        
        参数:
            added_files: 成功添加的文件列表
            failed_files: 失败的文件和错误消息列表
        """
        logger.info(f"文件添加事件: {len(added_files)} 成功, {len(failed_files)} 失败")
        
        # 显示状态消息
        if added_files:
            self.current_category = self.tabbed_file_list.get_current_category()
            category_name = get_category_name(self.current_category)
            status_msg = f"已添加 {len(added_files)} 个文件到 {category_name}"
            self.status_bar.add_message(status_msg, "success", False)
        
        if failed_files:
            error_msg = f"{len(failed_files)} 个文件添加失败"
            self.status_bar.add_message(error_msg, "danger", False)
        
        # 批量模式下，添加文件后需要刷新UI以更新格式转换面板
        if added_files and self.current_mode == "batch":
            logger.info("批量模式：添加文件后刷新UI以更新格式转换面板")
            self.refresh_ui(reason="files_added")
    
    def on_files_removed(self, file_path: str):
        """
        处理文件移除事件
        
        参数:
            file_path: 被移除的文件路径
        """
        logger.info(f"文件移除事件: {file_path}")
        
        # 如果移除的是选中文件，清空选中状态
        if self.selected_file and self.selected_file.file_path == file_path:
            self.selected_file = None
        
        # 统一刷新UI
        self.refresh_ui(reason="file_removed")
    
    def on_mode_changed(self, new_mode: str):
        """
        处理模式切换事件
        
        参数:
            new_mode: 新的模式 ("single" 或 "batch")
        """
        logger.info(f"模式切换: {self.current_mode} -> {new_mode}")
        
        self.current_mode = new_mode
        
        # 根据模式处理选中文件
        if new_mode == "single":
            # 单文件模式：确保有选中文件
            selected_file = self.tabbed_file_list.get_selected_file()
            if not selected_file and self.tabbed_file_list.has_files():
                # 没有选中文件但有文件，自动选中第一个
                logger.debug("单文件模式下无选中文件，自动选中第一个")
                self.current_category = self.tabbed_file_list.get_current_category()
                if self.current_category in self.tabbed_file_list.file_lists:
                    file_list = self.tabbed_file_list.file_lists[self.current_category]
                    file_objects = file_list.get_file_objects()
                    if file_objects:
                        file_list._select_file(file_objects[0])
                        logger.debug(f"已自动选中第一个文件: {file_objects[0].file_path}")
            else:
                self.selected_file = selected_file
        elif new_mode == "batch":
            # 批量模式：更新所有选项卡标题以显示文件数量
            if hasattr(self.tabbed_file_list, 'update_all_tab_titles'):
                logger.debug("批量模式：更新所有选项卡标题")
                self.tabbed_file_list.update_all_tab_titles()
        
        # 通知 action_panel 更新按钮状态
        if hasattr(self.action_panel, 'set_mode'):
            logger.debug(f"通知操作面板更新模式: {new_mode}")
            self.action_panel.set_mode(new_mode)
        
        # 统一刷新UI
        self.refresh_ui(reason="mode_changed")
        
        # 在批量模式下，所有UI更新完成后，让Canvas获取焦点
        if new_mode == "batch":
            # 获取当前选项卡的文件列表组件
            current_file_list = self.tabbed_file_list.get_current_file_list()
            if current_file_list and current_file_list.has_files():
                # 让Canvas获取焦点，使用户可以直接使用方向键
                current_file_list.canvas.focus_set()
                logger.info("批量模式切换完成：Canvas已获取焦点，可使用方向键调整文件顺序")
    
    def on_file_selected(self, file_info: FileInfo):
        """
        处理文件选中事件
        基于实际文件格式更新UI状态
        
        参数:
            file_info: 选中的文件信息对象
        """
        logger.info(f"文件选中事件: {file_info.file_path}")
        
        # 记录之前选中的文件
        previous_file = self.selected_file
        
        # 更新选中文件
        self.selected_file = file_info
        
        # 基于选中文件的实际类别更新UI状态
        if file_info.actual_category and file_info.actual_category != self.current_category:
            logger.debug(f"文件实际类别与当前选项卡不一致: {file_info.actual_category} != {self.current_category}")
            # 切换到正确的选项卡
            self.tabbed_file_list.set_current_tab(file_info.actual_category)
            self.current_category = file_info.actual_category
        
        # 判断是否需要完整UI刷新
        need_full_refresh = True
        
        if (self.current_mode == "batch" and 
            previous_file and 
            previous_file.actual_category == file_info.actual_category and
            previous_file.actual_format == file_info.actual_format):
            # 批量模式下，同类别、同格式的文件切换，跳过完整刷新
            logger.info("批量模式同选项卡内文件切换，跳过UI刷新")
            need_full_refresh = False
        
        # 根据判断结果决定是否刷新UI
        if need_full_refresh:
            logger.debug("执行完整UI刷新")
            self.refresh_ui(reason="file_selected")
        else:
            logger.debug("跳过完整UI刷新，仅更新轻量级状态")
            
            # 根据文件类别更新关键信息
            if file_info.actual_category == 'layout':
                # 版式类文件：更新PDF信息
                self._update_layout_info(file_info)
            elif file_info.actual_category == 'spreadsheet':
                # 表格类文件：更新基准表格名称
                self._update_spreadsheet_info(file_info)
        
        # 在单文件模式下，通知拖拽区更新显示
        if self.current_mode == "single" and self.file_drop_area:
            logger.debug(f"单文件模式：通知拖拽区更新显示 - {file_info.file_path}")
            self.file_drop_area.update_display(file_info)
        
        # 更新操作面板的基准表格显示
        if self.action_panel and hasattr(self.action_panel, 'set_reference_table'):
            import os
            file_name = os.path.basename(file_info.file_path)
            self.action_panel.set_reference_table(file_name)
            logger.debug(f"已更新基准表格显示: {file_name}")
        
        # 显示选中文件状态
        category_name = get_category_name(file_info.actual_category)
        status_msg = f"已选中: {file_info.file_name} ({category_name})"
        if file_info.warning_message:
            status_msg += f" {file_info.warning_message}"
        self.status_bar.add_message(status_msg, "info", False)
    
    def on_template_selected(self, template_type: str, template_name: str):
        """
        处理模板选择事件
        根据模板类型动态更新操作面板
        
        参数:
            template_type: 模板类型
            template_name: 模板名称
        """
        logger.info("----------------------------------------")
        logger.info(f"模板选择事件触发")
        logger.info(f"模板类型: {template_type}")
        logger.info(f"模板名称: {template_name}")
        logger.info("----------------------------------------")
        
        # 验证前提条件：必须有选中的文件
        if not self.selected_file:
            logger.warning("无选中文件，暂不更新操作面板")
            return
        
        # 验证文件类别：只有文本类文件需要模板
        if self.selected_file.actual_category != "text":
            logger.warning(f"当前文件类别 '{self.selected_file.actual_category}' 不需要模板")
            return
        
        # 确保每次模板选择都会更新主窗口的状态
        if self.main_window:
            self.main_window.selected_template = (template_type, template_name)
            logger.info(f"已同步更新主窗口的selected_template状态: ({template_type}, {template_name})")
        else:
            logger.warning("main_window引用不存在，无法同步状态")
        
        # 获取文件路径
        file_path = self.selected_file.file_path
        
        # 根据模板类型更新操作面板
        logger.info(f"开始更新操作面板...")
        
        if template_type == "xlsx":
            # Excel表格模板
            logger.info("检测到Excel模板，切换到'MD转表格'模式")
            
            if hasattr(self.action_panel, 'setup_for_md_to_spreadsheet'):
                self.action_panel.setup_for_md_to_spreadsheet(file_path)
            else:
                logger.error("操作面板不支持 setup_for_md_to_spreadsheet 方法")
        
        else:
            # Word文档模板
            logger.info(f"检测到文档模板({template_type})，切换到'MD转文档'模式")
            
            if hasattr(self.action_panel, 'setup_for_md_to_document'):
                self.action_panel.setup_for_md_to_document(file_path)
            else:
                logger.error("操作面板不支持 setup_for_md_to_document 方法")
        
        # 更新状态栏
        status_msg = f"已选择模板: {template_name}"
        self.status_bar.add_message(status_msg, "secondary", False)
        logger.info(f"状态栏已更新: {status_msg}")
        
        logger.info("----------------------------------------")
        logger.info("模板选择事件处理完成")
        logger.info("----------------------------------------")
    
    def refresh_ui(self, reason: str = None):
        """
        统一的UI刷新方法
        根据当前完整状态决定UI显示
        
        参数:
            reason: 触发刷新的原因
        """
        logger.info(f"===== 开始刷新UI =====")
        logger.info(f"触发原因: {reason}")
        logger.info(f"当前模式: {self.current_mode}")
        
        # 1. 获取当前状态
        self.current_category = self.tabbed_file_list.get_current_category()
        current_files = self.tabbed_file_list.get_current_files()
        has_files = len(current_files) > 0
        
        logger.info(f"当前选项卡类别: {self.current_category}")
        logger.info(f"当前文件列表: {current_files}")
        logger.info(f"有文件: {has_files}")
        logger.info(f"选中文件: {self.selected_file.file_path if self.selected_file else None}")
        
        # 2. 确定要显示的文件信息
        display_file = None
        
        if self.current_mode == "single":
            # 单文件模式：使用选中文件
            display_file = self.selected_file
            logger.info(f"单文件模式 - 使用选中文件: {display_file.file_path if display_file else None}")
        else:
            # 批量模式：使用当前选项卡的第一个文件对象
            if current_files:
                # 获取第一个文件对象用于确定类别
                if self.current_category in self.tabbed_file_list.file_lists:
                    file_list = self.tabbed_file_list.file_lists[self.current_category]
                    file_objects = file_list.get_file_objects()
                    if file_objects:
                        display_file = file_objects[0]
                        logger.info(f"批量模式 - 使用第一个文件对象: {display_file.file_path}")
                        
                        # 批量模式下，如果没有选中文件，自动将第一个文件设为选中
                        # 确保 on_template_selected 等回调能获取到有效的文件上下文
                        if not self.selected_file:
                            self.selected_file = display_file
                            logger.debug(f"批量模式 - 自动设置选中文件为第一个文件: {display_file.file_path}")
                    else:
                        logger.warning("批量模式 - 无法获取文件对象列表")
                else:
                    logger.warning(f"批量模式 - 当前类别 {self.current_category} 不在文件列表中")
            else:
                logger.info("批量模式 - 无文件")
        
        # 3. 根据状态决定UI
        logger.info(f"判断UI状态: has_files={has_files}, display_file={display_file is not None}")
        
        if not has_files or not display_file:
            logger.info("显示无文件状态")
            self._show_no_files_state()
        else:
            logger.info(f"显示文件状态: {display_file.file_path}, 实际类别: {display_file.actual_category}")
            self._show_file_state(display_file, current_files)
        
        # 批量模式下，同步更新FileDropArea的统计显示
        if self.current_mode == "batch" and self.file_drop_area:
            logger.debug("批量模式：更新FileDropArea统计显示")
            self.file_drop_area.update_display(None)
        
        logger.info(f"===== UI刷新完成 =====\n")
    
    def _show_no_files_state(self):
        """显示无文件状态"""
        logger.debug("显示无文件状态")
        
        # 通过主窗口方法隐藏面板框架
        if self.main_window:
            if hasattr(self.main_window, 'hide_action_panel'):
                self.main_window.hide_action_panel()
            if hasattr(self.main_window, 'hide_template_selector'):
                self.main_window.hide_template_selector()
        
        # 如果文件拖拽区域在批量模式，恢复无文件状态
        if (self.current_mode == "batch" and 
            hasattr(self.file_drop_area, '_switch_to_empty_state')):
            self.file_drop_area._switch_to_empty_state()
        
        # 更新状态栏
        category_name = get_category_name(self.current_category) if self.current_category else "当前"
        status_msg = f"{category_name}选项卡无文件"
        self.status_bar.add_message(status_msg, "secondary", False)
    
    def _show_file_state(self, file_info: FileInfo, current_files: List[str]):
        """
        显示有文件状态的UI
        基于文件的实际格式显示UI
        
        参数:
            file_info: 要显示的文件信息
            current_files: 当前文件列表
        """
        logger.debug(f"显示文件状态: {file_info.file_path}, 实际类别: {file_info.actual_category}")
        
        # 根据文件的实际类别显示相应的UI组件
        category = file_info.actual_category
        
        if category == "text":
            self._setup_text_ui(file_info, current_files)
        elif category == "document":
            self._setup_document_ui(file_info, current_files)
        elif category == "spreadsheet":
            self._setup_spreadsheet_ui(file_info, current_files)
        elif category == "layout":
            self._setup_layout_ui(file_info, current_files)
        elif category == "image":
            self._setup_image_ui(file_info, current_files)
        else:
            logger.warning(f"未知的文件类别: {category}")
            self._setup_unknown_ui()
    
    def _setup_text_ui(self, file_info: FileInfo, current_files: List[str]):
        """
        设置文本类文件的UI
        
        参数:
            file_info: 文件信息对象
            current_files: 当前文件列表
        """
        logger.info("设置文本类文件UI")
        
        # 1. 显示模板选择器框架（右栏）
        if self.main_window and hasattr(self.main_window, 'show_template_selector'):
            logger.info("显示模板选择器框架")
            self.main_window.show_template_selector()
        else:
            logger.warning("main_window不可用，回退到只显示组件内容")
            if hasattr(self.template_selector, 'show'):
                self.template_selector.show()
        
        # 2. 显示操作面板框架（下栏）
        if self.main_window and hasattr(self.main_window, 'show_action_panel'):
            logger.info("显示操作面板框架")
            self.main_window.show_action_panel()
        else:
            logger.warning("main_window不可用，回退到只显示组件内容")
            if hasattr(self.action_panel, 'show'):
                self.action_panel.show()
        
        # 3. 激活对应的模板选项卡并自动选中第一个模板
        default_template_type = "docx"  # 默认值
        
        if self.config_manager:
            try:
                default_template_type = self.config_manager.get_default_md_template_type()
                logger.info(f"从配置读取到默认模板类型: {default_template_type}")
            except Exception as e:
                logger.error(f"读取配置失败: {e}")
                logger.debug("使用默认值: docx")
        else:
            logger.warning("config_manager未初始化，使用默认值: docx")
        
        logger.info(f"文件拖入，激活{default_template_type.upper()}模板选项卡")
        
        if hasattr(self.template_selector, 'activate_and_select'):
            self.template_selector.activate_and_select(default_template_type)
            logger.info(f"已激活并选中第一个{default_template_type.upper()}模板")
        else:
            logger.error("模板选择器不支持 activate_and_select 方法")
        
        logger.info("文本类文件UI设置完成\n")
    
    def _setup_document_ui(self, file_info: FileInfo, current_files: List[str]):
        """
        设置文档类文件的UI
        
        参数:
            file_info: 文件信息对象
            current_files: 当前文件列表
        """
        logger.info("设置文档类文件UI")
        
        # 1. 显示格式转换面板（右栏）
        if self.main_window and hasattr(self.main_window, 'show_conversion_panel'):
            current_format = file_info.actual_format or 'docx'
            logger.info(f"显示格式转换面板")
            self.main_window.show_conversion_panel('document', current_format, file_info.file_path)
        else:
            logger.warning("main_window不支持转换面板，回退到隐藏模板选择器")
            if self.main_window and hasattr(self.main_window, 'hide_template_selector'):
                self.main_window.hide_template_selector()
            elif hasattr(self.template_selector, 'hide'):
                self.template_selector.hide()
        
        # 2. 显示操作面板框架（下栏）
        if self.main_window and hasattr(self.main_window, 'show_action_panel'):
            logger.info("显示下栏框架")
            self.main_window.show_action_panel()
        else:
            logger.warning("main_window不可用，回退到只显示组件内容")
            if hasattr(self.action_panel, 'show'):
                self.action_panel.show()
        
        # 3. 设置操作面板为文档文件模式
        if hasattr(self.action_panel, 'setup_for_document_file'):
            logger.info("设置操作面板为文档文件模式")
            self.action_panel.setup_for_document_file(
                file_info.file_path, 
                current_files if self.current_mode == "batch" else None
            )
        
        logger.info("文档类文件UI设置完成\n")
    
    def _setup_spreadsheet_ui(self, file_info: FileInfo, current_files: List[str]):
        """
        设置表格类文件的UI
        
        参数:
            file_info: 文件信息对象
            current_files: 当前文件列表
        """
        logger.info("设置表格类文件UI")
        
        # 1. 显示格式转换面板（右栏）
        conversion_panel = None
        if self.main_window and hasattr(self.main_window, 'show_conversion_panel'):
            current_format = file_info.actual_format or 'xlsx'
            logger.info(f"显示格式转换面板")
            self.main_window.show_conversion_panel('spreadsheet', current_format, file_info.file_path)
            
            # 获取conversion_panel引用
            if hasattr(self.main_window, 'conversion_panel'):
                conversion_panel = self.main_window.conversion_panel
        else:
            logger.warning("main_window不支持转换面板，回退到隐藏模板选择器")
            if self.main_window and hasattr(self.main_window, 'hide_template_selector'):
                self.main_window.hide_template_selector()
            elif hasattr(self.template_selector, 'hide'):
                self.template_selector.hide()
        
        # 2. 显示操作面板框架（下栏）
        if self.main_window and hasattr(self.main_window, 'show_action_panel'):
            logger.info("显示下栏框架")
            self.main_window.show_action_panel()
        else:
            logger.warning("main_window不可用，回退到只显示组件内容")
            if hasattr(self.action_panel, 'show'):
                self.action_panel.show()
        
        # 3. 设置操作面板为表格文件模式
        if hasattr(self.action_panel, 'setup_for_spreadsheet_file'):
            logger.info("设置操作面板为表格文件模式")
            self.action_panel.setup_for_spreadsheet_file(
                file_info.file_path,
                current_files if self.current_mode == "batch" else None
            )
        
        # 4. 设置基准表格名称（单文件和批量模式都显示）
        if conversion_panel and hasattr(conversion_panel, 'set_reference_table'):
            import os
            file_name = os.path.basename(file_info.file_path)
            conversion_panel.set_reference_table(file_name)
            logger.info(f"已设置基准表格: {file_name}")
        
        logger.info("表格类文件UI设置完成\n")
    
    def _setup_layout_ui(self, file_info: FileInfo, current_files: List[str]):
        """
        设置版式类文件的UI
        
        参数:
            file_info: 文件信息对象
            current_files: 当前文件列表
        """
        logger.info("设置版式类文件UI")
        
        # 1. 显示格式转换面板（右栏）
        conversion_panel = None
        if self.main_window and hasattr(self.main_window, 'show_conversion_panel'):
            current_format = file_info.actual_format or 'pdf'
            logger.info(f"显示格式转换面板")
            self.main_window.show_conversion_panel('layout', current_format, file_info.file_path)
            
            # 获取conversion_panel引用
            if hasattr(self.main_window, 'conversion_panel'):
                conversion_panel = self.main_window.conversion_panel
        else:
            logger.warning("main_window不支持转换面板，回退到隐藏模板选择器")
            if self.main_window and hasattr(self.main_window, 'hide_template_selector'):
                self.main_window.hide_template_selector()
            elif hasattr(self.template_selector, 'hide'):
                self.template_selector.hide()
        
        # 2. 显示操作面板框架（下栏）
        if self.main_window and hasattr(self.main_window, 'show_action_panel'):
            logger.info("显示下栏框架")
            self.main_window.show_action_panel()
        else:
            logger.warning("main_window不可用，回退到只显示组件内容")
            if hasattr(self.action_panel, 'show'):
                self.action_panel.show()
        
        # 3. 设置操作面板为版式文件模式
        if hasattr(self.action_panel, 'setup_for_layout_file'):
            logger.info("设置操作面板为版式文件模式")
            self.action_panel.setup_for_layout_file(file_info.file_path)
        
        # 4. 设置PDF信息（页数+文件名）
        if conversion_panel and hasattr(conversion_panel, 'set_pdf_info'):
            try:
                import fitz  # PyMuPDF
                import os
                
                file_path = file_info.file_path
                file_name = os.path.basename(file_path)
                
                # 打开PDF获取页数
                with fitz.open(file_path) as pdf:
                    total_pages = len(pdf)
                
                logger.info(f"PDF总页数: {total_pages}")
                conversion_panel.set_pdf_info(total_pages, file_name)
            except Exception as e:
                logger.warning(f"设置PDF信息失败: {e}")
        
        logger.info("版式类文件UI设置完成\n")
    
    def _setup_image_ui(self, file_info: FileInfo, current_files: List[str]):
        """
        设置图片类文件的UI
        
        参数:
            file_info: 文件信息对象
            current_files: 当前文件列表
        """
        logger.info("设置图片类文件UI")
        
        # 1. 显示格式转换面板（右栏）
        if self.main_window and hasattr(self.main_window, 'show_conversion_panel'):
            current_format = file_info.actual_format or 'png'
            logger.info(f"显示格式转换面板")
            self.main_window.show_conversion_panel('image', current_format, file_info.file_path)
        else:
            logger.warning("main_window不支持转换面板，回退到隐藏模板选择器")
            if self.main_window and hasattr(self.main_window, 'hide_template_selector'):
                self.main_window.hide_template_selector()
            elif hasattr(self.template_selector, 'hide'):
                self.template_selector.hide()
        
        # 2. 显示操作面板框架（下栏）
        if self.main_window and hasattr(self.main_window, 'show_action_panel'):
            logger.info("显示下栏框架")
            self.main_window.show_action_panel()
        else:
            logger.warning("main_window不可用，回退到只显示组件内容")
            if hasattr(self.action_panel, 'show'):
                self.action_panel.show()
        
        # 3. 设置操作面板为图片文件模式
        if hasattr(self.action_panel, 'setup_for_image_file'):
            logger.info("设置操作面板为图片文件模式")
            self.action_panel.setup_for_image_file(file_info.file_path)
        
        logger.info("图片类文件UI设置完成\n")
    
    def _setup_unknown_ui(self):
        """设置未知文件类型的UI"""
        logger.debug("设置未知文件类型UI")
        
        # 隐藏所有面板
        if hasattr(self.template_selector, 'hide'):
            self.template_selector.hide()
        if hasattr(self.action_panel, 'hide'):
            self.action_panel.hide()
        
        # 显示错误状态
        self.status_bar.add_message("未知文件类型", "danger", False)
    
    def _update_layout_info(self, file_info: FileInfo):
        """
        轻量级更新版式文件信息
        
        参数:
            file_info: 文件信息对象
        """
        if self.main_window and hasattr(self.main_window, 'conversion_panel'):
            conversion_panel = self.main_window.conversion_panel
            if conversion_panel:
                try:
                    import fitz  # PyMuPDF
                    import os
                    
                    # 打开PDF获取页数
                    with fitz.open(file_info.file_path) as pdf:
                        total_pages = len(pdf)
                    
                    file_name = os.path.basename(file_info.file_path)
                    
                    # 更新PDF信息
                    if hasattr(conversion_panel, 'set_pdf_info'):
                        conversion_panel.set_pdf_info(total_pages, file_name)
                        logger.debug(f"轻量级更新PDF信息: {total_pages}页, {file_name}")
                    
                    # 清空拆分输入框
                    if hasattr(conversion_panel, 'clear_split_input'):
                        conversion_panel.clear_split_input()
                        
                except Exception as e:
                    logger.warning(f"轻量级更新PDF信息失败: {e}")
    
    def _update_spreadsheet_info(self, file_info: FileInfo):
        """
        轻量级更新表格文件信息
        
        参数:
            file_info: 文件信息对象
        """
        if self.main_window and hasattr(self.main_window, 'conversion_panel'):
            conversion_panel = self.main_window.conversion_panel
            if conversion_panel and hasattr(conversion_panel, 'set_reference_table'):
                import os
                file_name = os.path.basename(file_info.file_path)
                conversion_panel.set_reference_table(file_name)
                logger.debug(f"轻量级更新基准表格: {file_name}")
    
    
    def get_current_category(self) -> Optional[str]:
        """获取当前激活的选项卡类别"""
        return self.current_category
    
    def get_current_files(self) -> List[str]:
        """获取当前激活选项卡中的文件列表"""
        return self.tabbed_file_list.get_current_files()
    
    def get_current_mode(self) -> str:
        """获取当前模式"""
        return self.current_mode
    
    def has_files(self) -> bool:
        """检查当前选项卡是否有文件"""
        return len(self.current_files) > 0


# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 创建模拟组件
    class MockTabbedFileList:
        def get_current_files(self):
            return ["/fake/path/test1.docx", "/fake/path/test2.docx"]
        
        def get_current_category(self):
            return "document"
    
    class MockActionPanel:
        def show(self):
            logger.info("操作面板显示")
        
        def hide(self):
            logger.info("操作面板隐藏")
        
        def setup_for_document_file(self, file_path, file_list=None):
            logger.info(f"设置文档文件模式: {file_path}, 文件列表: {file_list}")
    
    class MockTemplateSelector:
        def show(self):
            logger.info("模板选择器显示")
        
        def hide(self):
            logger.info("模板选择器隐藏")
    
    class MockFileDropArea:
        def _switch_to_empty_state(self):
            logger.info("切换到无文件状态")
    
    class MockStatusBar:
        def add_message(self, message, message_type, show_location):
            logger.info(f"状态栏: {message} ({message_type})")
    
    # 创建模拟组件实例
    tabbed_file_list = MockTabbedFileList()
    action_panel = MockActionPanel()
    template_selector = MockTemplateSelector()
    file_drop_area = MockFileDropArea()
    status_bar = MockStatusBar()
    
    # 创建选项卡式文件管理器
    file_manager = SelectorManager(
        tabbed_file_list=tabbed_file_list,
        action_panel=action_panel,
        template_selector=template_selector,
        file_drop_area=file_drop_area,
        status_bar=status_bar
    )
    
    # 测试选项卡切换
    logger.info("=== 测试选项卡切换 ===")
    file_manager.on_tab_changed("document", "text")
    
    # 测试文件添加
    logger.info("=== 测试文件添加 ===")
    file_manager.on_files_added(
        ["/fake/path/test1.docx", "/fake/path/test2.docx"],
        [("/fake/path/unsupported.xyz", "不支持的文件类型")]
    )
    
    # 测试文件移除
    logger.info("=== 测试文件移除 ===")
    file_manager.on_files_removed("/fake/path/test1.docx")
    
    # 测试模式切换
    logger.info("=== 测试模式切换 ===")
    file_manager.on_mode_changed("batch")
    
    logger.info("选项卡式文件管理器测试完成")
