"""
文件选择器拖拽功能混入类

提供文件列表的拖拽支持功能：
- 设置拖拽目标
- 处理拖拽进入/离开/放置事件
- 解析拖拽文件数据
- 支持文件夹递归收集

依赖 tkinterdnd2 库实现拖拽功能，如果库不可用则拖拽功能将被禁用。
"""

import os
import re
import logging
from typing import List, Optional, TYPE_CHECKING

from gongwen_converter.utils.path_utils import collect_files_from_folder
from gongwen_converter.utils.file_type_utils import is_supported_file

logger = logging.getLogger(__name__)

# 尝试导入拖拽支持库
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKINTERDND2_AVAILABLE = True
    logger.debug("tkinterdnd2库已成功导入")
except ImportError:
    TkinterDnD = None
    DND_FILES = None
    TKINTERDND2_AVAILABLE = False
    logger.warning("tkinterdnd2未安装，拖拽功能将不可用")


class DragDropMixin:
    """
    拖拽功能混入类
    
    为文件选择器提供拖拽支持，包括：
    - 将Canvas注册为拖拽目标
    - 处理拖拽进入/离开/放置事件
    - 智能解析混合格式的拖拽文件数据
    - 支持文件夹递归收集支持的文件
    
    依赖属性（由主类提供）：
        canvas: Canvas组件引用
        master: 父组件引用
        drag_enabled: 拖拽是否启用
        is_dragging: 是否正在拖拽
    
    依赖方法（由主类提供）：
        add_files(): 批量添加文件方法
    """
    
    # 声明依赖的属性（由主类提供）
    canvas: 'Optional[object]'
    master: 'Optional[object]'
    drag_enabled: bool
    is_dragging: bool
    
    def _setup_drag_drop(self):
        """
        设置拖拽支持
        
        将Canvas注册为拖拽目标，并绑定拖拽事件处理器。
        如果 tkinterdnd2 库不可用，则禁用拖拽功能。
        """
        logger.debug("设置文件列表拖拽支持")
        
        if not TKINTERDND2_AVAILABLE:
            logger.warning("tkinterdnd2未安装，文件列表拖拽功能不可用")
            self.drag_enabled = False
            return
        
        try:
            # 将Canvas注册为拖拽目标
            self.canvas.drop_target_register(DND_FILES)
            self.canvas.dnd_bind('<<DragEnter>>', self._on_drag_enter)
            self.canvas.dnd_bind('<<DragLeave>>', self._on_drag_leave)
            self.canvas.dnd_bind('<<Drop>>', self._on_drop)
            logger.info("文件列表拖拽支持设置成功")
        except Exception as e:
            logger.error(f"设置文件列表拖拽支持失败: {str(e)}")
            self.drag_enabled = False
    
    def _on_drag_enter(self, event):
        """
        处理拖拽进入事件
        
        当文件被拖入文件列表区域时触发。
        """
        if not self.drag_enabled:
            return
        
        logger.debug("文件拖拽进入文件列表区域")
        self.is_dragging = True
        
        return "copy"
    
    def _on_drag_leave(self, event):
        """
        处理拖拽离开事件
        
        当拖拽的文件离开文件列表区域时触发。
        """
        if not self.drag_enabled:
            return
        
        logger.debug("文件拖拽离开文件列表区域")
        self.is_dragging = False
    
    def _on_drop(self, event):
        """
        处理文件放置事件
        
        当文件被放置到文件列表区域时触发。
        使用智能分类和自动切换选项卡功能。
        """
        if not self.drag_enabled:
            return
        
        logger.debug("文件放置到文件列表事件触发")
        self.is_dragging = False
        
        try:
            file_data = event.data.strip()
            logger.debug(f"文件列表接收到拖拽数据: {file_data}")
            
            # 解析拖拽的文件
            files = self._parse_dropped_files(file_data)
            
            if len(files) > 0:
                logger.info(f"文件列表拖拽文件: {len(files)} 个文件")
                
                # 处理文件（支持文件夹递归）
                processed_files = self._process_batch_files(files)
                logger.info(f"处理后得到 {len(processed_files)} 个支持文件")
                
                if processed_files:
                    # 获取 TabbedFileSelector 的引用
                    tabbed_list = self._get_tabbed_batch_file_list()
                    
                    if tabbed_list:
                        # 使用 TabbedFileSelector 的 add_files 方法
                        # 这样可以实现智能分类和自动切换选项卡
                        added_files, failed_files = tabbed_list.add_files(processed_files)
                        
                        if added_files:
                            logger.info(f"成功添加 {len(added_files)} 个文件，已自动分类和切换选项卡")
                        
                        if failed_files:
                            logger.warning(f"有 {len(failed_files)} 个文件添加失败")
                            for file_path, error_msg in failed_files:
                                logger.debug(f"失败文件: {file_path}, 原因: {error_msg}")
                    else:
                        # 如果无法获取 TabbedFileSelector，回退到本地处理
                        logger.warning("无法获取 TabbedFileSelector，使用本地处理")
                        added_files, failed_files = self.add_files(processed_files, auto_select_first=True)
                        
                        if added_files:
                            logger.info(f"成功添加 {len(added_files)} 个文件到当前列表")
                else:
                    logger.warning("未找到任何支持的文件")
            else:
                logger.warning("未检测到有效文件")
        
        except Exception as e:
            logger.error(f"处理文件列表拖拽失败: {str(e)}")
    
    def _get_tabbed_batch_file_list(self):
        """
        获取 TabbedFileSelector 的引用
        
        尝试通过组件层级向上查找主窗口，然后获取其中的 TabbedFileSelector。
        
        返回:
            TabbedFileSelector实例或None
        """
        try:
            # 向上查找主窗口
            current = self.master
            while current and not hasattr(current, '_main_window'):
                current = current.master if hasattr(current, 'master') else None
            
            if current and hasattr(current, '_main_window'):
                main_window = current._main_window
                if hasattr(main_window, 'tabbed_batch_file_list'):
                    return main_window.tabbed_batch_file_list
            
            # 如果上面的方法失败，尝试直接从父组件获取
            # FileSelector 的父组件应该是 TabbedFileSelector 中的一个 tab
            current = self.master
            while current:
                # 检查父组件的父组件是否是 TabbedFileSelector
                if hasattr(current, 'master') and current.master:
                    parent = current.master
                    if hasattr(parent, 'master') and parent.master:
                        grandparent = parent.master
                        if grandparent.__class__.__name__ == 'TabbedFileSelector':
                            return grandparent
                current = current.master if hasattr(current, 'master') else None
                
        except Exception as e:
            logger.error(f"获取 TabbedFileSelector 失败: {e}")
        
        return None
    
    def _parse_dropped_files(self, file_data: str) -> List[str]:
        """
        解析拖拽的文件数据
        
        能够正确处理以下格式：
        - {file with spaces.jpg} file2.jpg
        - {file1} {file2}
        - file1 file2
        - 单个文件
        
        参数:
            file_data: 拖拽事件携带的文件数据字符串
            
        返回:
            List[str]: 解析出的文件路径列表
        """
        files = []
        
        # 策略1: 使用正则表达式提取所有花括号包裹的路径
        # 匹配模式: {路径内容}
        brace_pattern = r'\{([^}]+)\}'
        brace_matches = re.findall(brace_pattern, file_data)
        
        if brace_matches:
            # 找到了花括号包裹的路径
            files.extend(brace_matches)
            
            # 移除已提取的花括号部分，处理剩余内容
            remaining = re.sub(brace_pattern, '', file_data).strip()
            
            if remaining:
                # 剩余部分可能包含没有花括号的路径（空格分隔）
                potential_paths = remaining.split()
                for path in potential_paths:
                    path = path.strip()
                    if path and os.path.exists(path):
                        files.append(path)
        else:
            # 没有花括号，使用传统的空格分隔方式
            if ' ' in file_data:
                # 先尝试整个字符串是否是有效路径
                if os.path.exists(file_data):
                    files = [file_data]
                else:
                    # 按空格分割并验证每个路径
                    potential_files = file_data.split(' ')
                    files = [f for f in potential_files if f.strip() and os.path.exists(f.strip())]
                    if not files:
                        # 如果都无效，保留原始数据
                        files = [file_data]
            else:
                # 单个文件
                files = [file_data]
        
        # 清理并去重
        cleaned_files = []
        seen = set()
        for f in files:
            f = f.strip()
            if f and f not in seen:
                cleaned_files.append(f)
                seen.add(f)
        
        return cleaned_files
    
    def _process_batch_files(self, files: List[str]) -> List[str]:
        """
        处理批量模式下的拖拽文件
        
        支持文件夹递归收集所有支持的文件。
        
        参数:
            files: 原始文件/文件夹路径列表
            
        返回:
            List[str]: 处理后的支持文件路径列表
        """
        processed_files = []
        
        for file_path in files:
            if os.path.isdir(file_path):
                # 如果是文件夹，递归收集所有支持的文件
                folder_files = collect_files_from_folder(file_path)
                logger.info(f"从文件夹 {file_path} 收集到 {len(folder_files)} 个支持文件")
                processed_files.extend(folder_files)
            else:
                # 如果是文件，检查是否支持
                if is_supported_file(file_path):
                    processed_files.append(file_path)
                else:
                    logger.debug(f"跳过不支持的文件: {file_path}")
        
        return processed_files
