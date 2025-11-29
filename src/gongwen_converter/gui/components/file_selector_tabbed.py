"""
选项卡式批量文件列表组件
将批量文件列表分为5个选项卡：文本类、文档类、表格类、版式类、图片类
"""

import os
import logging
import tkinter as tk
from typing import List, Dict, Callable, Optional, Tuple
from gongwen_converter.utils.dpi_utils import scale
from gongwen_converter.utils.file_type_utils import get_file_category, get_category_name, CATEGORY_NAMES, activate_optimal_tab, get_actual_file_category
from gongwen_converter.gui.components.file_selector import FileSelector

# 导入ttkbootstrap用于界面美化和样式管理
import ttkbootstrap as tb
from ttkbootstrap.constants import *

logger = logging.getLogger(__name__)


class TabbedFileSelector(tb.Frame):
    """
    选项卡式批量文件列表组件
    使用 ttk.Notebook 实现5个选项卡，每个选项卡对应一个文件类别
    """
    
    def __init__(self, master, 
                 on_file_removed: Optional[Callable] = None,
                 on_file_opened: Optional[Callable] = None,
                 on_tab_changed: Optional[Callable] = None,
                 on_file_selected: Optional[Callable] = None,
                 on_files_added: Optional[Callable] = None,
                 **kwargs):
        """
        初始化选项卡式批量文件列表组件
        
        参数:
            master: 父组件
            on_file_removed: 文件移除回调函数
            on_file_opened: 文件打开回调函数
            on_tab_changed: 选项卡切换回调函数
            on_file_selected: 文件选中回调函数
            on_files_added: 文件添加完成回调函数
        """
        logger.debug("初始化选项卡式批量文件列表组件")
        
        # 调用父类构造函数
        super().__init__(master, **kwargs)
        
        # 存储回调函数
        self.on_file_removed = on_file_removed
        self.on_file_opened = on_file_opened
        self.on_tab_changed = on_tab_changed
        self.on_file_selected = on_file_selected
        self.on_files_added = on_files_added
        
        # 存储所有选项卡和对应的文件列表组件
        self.tabs: Dict[str, tb.Frame] = {}
        self.file_lists: Dict[str, FileSelector] = {}
        
        # 当前激活的选项卡
        self.current_tab: Optional[str] = None
        
        # 创建界面元素
        self._create_widgets()
        
        logger.info("选项卡式批量文件列表组件初始化完成")
    
    def _create_widgets(self):
        """创建界面元素"""
        logger.debug("创建选项卡式批量文件列表界面元素")
        
        # 配置grid布局
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # 创建选项卡控件
        self.notebook = tb.Notebook(self, bootstyle="warning")
        self.notebook.grid(row=0, column=0, sticky="nsew")
        
        # 绑定选项卡切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
        # 创建5个选项卡
        self._create_tabs()
        
        logger.debug("选项卡式批量文件列表界面元素创建完成")
    
    def _create_tabs(self):
        """创建5个选项卡"""
        logger.debug("创建5个文件类别选项卡")
        
        # 5个文件类别
        categories = ['text', 'document', 'spreadsheet', 'layout', 'image']
        
        for category in categories:
            # 创建选项卡框架
            tab_frame = tb.Frame(self.notebook, bootstyle="default")
            
            # 创建对应的文件列表组件
            file_list = FileSelector(
                tab_frame,
                on_file_removed=self._on_file_removed,
                on_file_opened=self.on_file_opened,
                on_list_cleared=lambda cat=category: self._on_list_cleared(cat),
                on_file_selected=self._on_file_selected  # 添加文件选中回调
            )
            
            # 存储选项卡和文件列表
            self.tabs[category] = tab_frame
            self.file_lists[category] = file_list
            
            # 添加选项卡到notebook
            display_name = get_category_name(category)
            self.notebook.add(tab_frame, text=f"  {display_name}  ")
        
        # 设置默认激活的选项卡（第一个）
        if categories:
            self.current_tab = categories[0]
            self.notebook.select(0)
    
    def _on_tab_changed(self, event):
        """处理选项卡切换事件"""
        selected_index = self.notebook.index("current")
        categories = list(self.tabs.keys())
        
        if 0 <= selected_index < len(categories):
            new_tab = categories[selected_index]
            old_tab = self.current_tab
            self.current_tab = new_tab
            
            logger.debug(f"选项卡切换: {old_tab} -> {new_tab}")
            
            # 调用选项卡切换回调
            if self.on_tab_changed:
                self.on_tab_changed(new_tab, old_tab)
    
    def _on_file_removed(self, file_path: str):
        """处理文件移除事件"""
        logger.debug(f"文件被移除: {file_path}")
        
        # 找到文件所属的类别并更新选项卡标题
        for category, file_list in self.file_lists.items():
            files = file_list.get_files()
            if file_path in files:
                # 文件移除后（由FileSelector处理），更新选项卡标题
                # 需要注意：此时文件可能已被FileSelector移除，所以需要在下一个事件循环中更新
                self.after(10, lambda cat=category: self._update_tab_title(cat))
                break
        
        # 调用外部回调
        if self.on_file_removed:
            self.on_file_removed(file_path)
    
    def _update_tab_title(self, category: str):
        """
        更新选项卡标题
        - 有文件时：显示 " 类别名 (数量) " (两侧各1个空格)
        - 无文件时：显示 "  类别名  " (两侧各2个空格)
        
        参数:
            category: 文件类别
        """
        if category not in self.tabs:
            return
        
        # 获取文件数量
        file_count = self.get_file_count(category)
        display_name = get_category_name(category)
        
        # 根据文件数量决定显示格式
        if file_count > 0:
            # 有文件时使用1个空格，更紧凑以避免显示不全
            title = f" {display_name} ({file_count}) "
        else:
            # 无文件时使用2个空格，保持美观
            title = f"  {display_name}  "
        
        # 更新选项卡标题
        categories = list(self.tabs.keys())
        tab_index = categories.index(category)
        self.notebook.tab(tab_index, text=title)
        
        logger.debug(f"更新选项卡标题: {category} -> {title.strip()}")
    
    def update_all_tab_titles(self):
        """
        更新所有选项卡的标题
        用于模式切换等需要批量更新标题的场景
        """
        logger.debug("批量更新所有选项卡标题")
        for category in self.file_lists.keys():
            self._update_tab_title(category)
        logger.debug("所有选项卡标题更新完成")
    
    def _on_list_cleared(self, category: str):
        """处理列表清空事件"""
        logger.debug(f"选项卡 {category} 列表已清空")
        
        # 更新选项卡标题
        self._update_tab_title(category)
        
        # 如果当前激活选项卡无文件，触发UI重置
        if self.current_tab == category:
            # 这里可以添加UI重置逻辑
            pass
    
    def _on_file_selected(self, file_info):
        """处理文件选中事件，转发到外部回调"""
        logger.debug(f"转发文件选中事件: {file_info.file_path}")
        
        # 如果有外部文件选中回调，则调用
        if hasattr(self, 'on_file_selected') and self.on_file_selected:
            self.on_file_selected(file_info)
    
    def add_file(self, file_path: str) -> Tuple[bool, str]:
        """
        添加文件到对应的选项卡
        
        参数:
            file_path: 文件路径
            
        返回:
            Tuple[bool, str]: (是否成功添加, 错误消息)
        """
        logger.debug(f"添加文件到选项卡式列表: {file_path}")
        
        # 获取文件类别
        category = get_file_category(file_path)
        if category == 'unknown':
            return False, "不支持的文件类型"
        
        # 添加到对应选项卡的文件列表
        if category in self.file_lists:
            success, error_msg = self.file_lists[category].add_file(file_path)
            if success:
                logger.info(f"文件已添加到 {category} 选项卡: {file_path}")
                # 更新选项卡标题
                self._update_tab_title(category)
            return success, error_msg
        else:
            logger.error(f"未知的文件类别: {category}")
            return False, f"未知的文件类别: {category}"
    
    def add_files(self, file_paths: List[str]) -> Tuple[List[str], List[Tuple[str, str]]]:
        """
        批量添加文件到对应的选项卡
        基于实际文件格式进行分类，避免伪装后缀名问题
        
        参数:
            file_paths: 文件路径列表
            
        返回:
            Tuple[List[str], List[Tuple[str, str]]]: (成功添加的文件列表, 失败的文件和错误消息列表)
        """
        logger.debug(f"批量添加文件到选项卡式列表: {len(file_paths)} 个文件")
        
        added_files = []
        failed_files = []
        
        # 统计各类别文件数量（基于实际文件格式）
        category_count = {'text': 0, 'document': 0, 'spreadsheet': 0, 'layout': 0, 'image': 0}
        categorized_files = {'text': [], 'document': [], 'spreadsheet': [], 'layout': [], 'image': []}
        
        # 基于实际文件格式分类文件
        for file_path in file_paths:
            category = get_actual_file_category(file_path)
            if category in categorized_files:
                category_count[category] += 1
                categorized_files[category].append(file_path)
            else:
                failed_files.append((file_path, "不支持的文件类型"))
        
        # 添加文件到对应选项卡，确保第一个文件被自动选中
        is_first_batch = True  # 标记是否是第一批添加的文件
        affected_categories = []  # 记录受影响的类别
        
        for category, files in categorized_files.items():
            if files:
                # 只在第一批文件时自动选中第一个
                auto_select_first = is_first_batch
                success_files, failed_list = self.file_lists[category].add_files(files, auto_select_first=auto_select_first)
                added_files.extend(success_files)
                failed_files.extend(failed_list)
                
                # 如果成功添加了文件，标记不再是第一批，并记录受影响的类别
                if success_files:
                    is_first_batch = False
                    affected_categories.append(category)
        
        # 更新所有受影响类别的选项卡标题
        for category in affected_categories:
            self._update_tab_title(category)
        
        # 使用统一算法激活最优选项卡
        if added_files:
            self._activate_optimal_tab(added_files)
        
        logger.info(f"批量添加完成: {len(added_files)}/{len(file_paths)} 个文件成功")
        
        # 触发文件添加完成回调
        if added_files and self.on_files_added:
            self.on_files_added(added_files, failed_files)
            logger.debug(f"触发文件添加完成回调: {len(added_files)} 个文件")
        
        return added_files, failed_files
    
    def _activate_optimal_tab(self, file_paths: List[str]):
        """
        使用统一算法激活最优选项卡
        基于实际文件格式检测，避免伪装后缀名问题
        
        参数:
            file_paths: 文件路径列表
        """
        if not file_paths:
            return
        
        # 使用统一的选项卡激活算法
        optimal_category = activate_optimal_tab(file_paths)
        
        # 激活对应选项卡
        if optimal_category in self.tabs:
            categories = list(self.tabs.keys())
            tab_index = categories.index(optimal_category)
            self.notebook.select(tab_index)
            self.current_tab = optimal_category
            logger.info(f"自动激活最优选项卡: {optimal_category} (文件数: {len(file_paths)})")
        else:
            logger.warning(f"无法激活未知的选项卡类别: {optimal_category}")
    
    def get_current_category(self) -> Optional[str]:
        """
        获取当前激活的选项卡类别
        
        返回:
            Optional[str]: 当前激活的选项卡类别
        """
        return self.current_tab
    
    def get_current_file_list(self) -> Optional[FileSelector]:
        """
        获取当前激活选项卡对应的文件列表对象
        
        此方法用于批量处理时获取文件列表对象，以便更新文件状态。
        
        返回:
            Optional[FileSelector]: 当前激活选项卡的文件列表对象，如果没有激活选项卡则返回None
        """
        if self.current_tab and self.current_tab in self.file_lists:
            return self.file_lists[self.current_tab]
        return None
    
    def get_current_files(self) -> List[str]:
        """
        获取当前激活选项卡中的文件列表
        
        返回:
            List[str]: 当前激活选项卡中的文件路径列表
        """
        if self.current_tab and self.current_tab in self.file_lists:
            return self.file_lists[self.current_tab].get_files()
        return []
    
    def get_files_by_category(self, category: str) -> List[str]:
        """
        获取指定类别的文件列表
        
        参数:
            category: 文件类别
            
        返回:
            List[str]: 指定类别的文件路径列表
        """
        if category in self.file_lists:
            return self.file_lists[category].get_files()
        return []
    
    def get_all_files(self) -> List[str]:
        """
        获取所有选项卡中的文件列表
        
        返回:
            List[str]: 所有文件路径列表
        """
        all_files = []
        for file_list in self.file_lists.values():
            all_files.extend(file_list.get_files())
        return all_files
    
    def get_file_count(self, category: Optional[str] = None) -> int:
        """
        获取文件数量
        
        参数:
            category: 指定类别，如果为None则获取所有文件数量
            
        返回:
            int: 文件数量
        """
        if category:
            if category in self.file_lists:
                return self.file_lists[category].get_file_count()
            return 0
        else:
            total = 0
            for file_list in self.file_lists.values():
                total += file_list.get_file_count()
            return total
    
    def has_files(self, category: Optional[str] = None) -> bool:
        """
        检查是否有文件
        
        参数:
            category: 指定类别，如果为None则检查所有选项卡
            
        返回:
            bool: 是否有文件
        """
        if category:
            if category in self.file_lists:
                return self.file_lists[category].has_files()
            return False
        else:
            for file_list in self.file_lists.values():
                if file_list.has_files():
                    return True
            return False
    
    def clear_all(self):
        """清空所有选项卡的文件"""
        logger.debug("清空所有选项卡的文件")
        
        for file_list in self.file_lists.values():
            file_list.clear_all_files()
        
        # 更新所有选项卡标题
        for category in self.file_lists.keys():
            self._update_tab_title(category)
        
        logger.info("所有选项卡的文件已清空")
    
    def clear_category(self, category: str):
        """
        清空指定类别的文件
        
        参数:
            category: 文件类别
        """
        logger.debug(f"清空 {category} 选项卡的文件")
        
        if category in self.file_lists:
            self.file_lists[category].clear_all_files()
            # 更新选项卡标题
            self._update_tab_title(category)
            logger.info(f"{category} 选项卡的文件已清空")
        else:
            logger.warning(f"未知的文件类别: {category}")
    
    def update_file_status(self, file_path: str, status: str, output_path: Optional[str] = None):
        """
        更新文件状态
        
        参数:
            file_path: 文件路径
            status: 新状态
            output_path: 输出文件路径（可选）
        """
        logger.debug(f"更新文件状态: {file_path} -> {status}")
        
        # 找到文件所在的类别并更新状态
        for category, file_list in self.file_lists.items():
            files = file_list.get_files()
            if file_path in files:
                file_list.update_file_status(file_path, status, output_path)
                logger.debug(f"在 {category} 选项卡中更新文件状态")
                return
        
        logger.warning(f"未找到文件: {file_path}")
    
    def set_current_tab(self, category: str):
        """
        设置当前激活的选项卡
        
        参数:
            category: 文件类别
        """
        if category in self.tabs:
            categories = list(self.tabs.keys())
            tab_index = categories.index(category)
            self.notebook.select(tab_index)
            self.current_tab = category
            logger.debug(f"手动设置当前选项卡为: {category}")
        else:
            logger.warning(f"未知的文件类别: {category}")
    
    def get_selected_file(self):
        """
        获取当前激活选项卡中的选中文件
        
        返回:
            Optional[FileInfo]: 选中的文件信息，如果没有选中则返回None
        """
        if self.current_tab and self.current_tab in self.file_lists:
            return self.file_lists[self.current_tab].get_selected_file()
        return None
    
    def get_selected_file_path(self) -> Optional[str]:
        """
        获取当前激活选项卡中选中文件的路径
        
        返回:
            Optional[str]: 选中文件的路径，如果没有选中则返回None
        """
        selected_file = self.get_selected_file()
        return selected_file.file_path if selected_file else None


# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 创建测试窗口
    root = tb.Window(title="选项卡式批量文件列表组件测试", themename="morph")
    root.geometry("600x400")
    
    def on_file_removed(file_path):
        logger.info(f"文件被移除: {file_path}")
    
    def on_file_opened(file_path):
        logger.info(f"文件被打开: {file_path}")
    
    def on_tab_changed(new_tab, old_tab):
        logger.info(f"选项卡切换: {old_tab} -> {new_tab}")
    
    # 创建选项卡式批量文件列表组件
    tabbed_file_list = TabbedFileSelector(
        root,
        on_file_removed=on_file_removed,
        on_file_opened=on_file_opened,
        on_tab_changed=on_tab_changed
    )
    tabbed_file_list.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
    
    # 添加测试按钮框架
    test_frame = tb.Frame(root, bootstyle="default")
    test_frame.pack(fill=tk.X, padx=10, pady=10)
    
    # 测试按钮回调函数
    def test_add_files():
        files = [
            "/fake/path/test1.docx",  # 文档类
            "/fake/path/test2.md",    # 文本类
            "/fake/path/test3.xlsx",  # 表格类
            "/fake/path/test4.pdf",   # 版式类
            "/fake/path/test5.jpg",   # 图片类
            "/fake/path/test6.docx",  # 文档类
            "/fake/path/test7.png"    # 图片类
        ]
        added, failed = tabbed_file_list.add_files(files)
        logger.info(f"添加结果: {len(added)} 成功, {len(failed)} 失败")
    
    def test_clear_all():
        tabbed_file_list.clear_all()
    
    def test_get_current_files():
        files = tabbed_file_list.get_current_files()
        current_category = tabbed_file_list.get_current_category()
        logger.info(f"当前选项卡 ({current_category}) 文件: {files}")
    
    def test_get_all_files():
        files = tabbed_file_list.get_all_files()
        logger.info(f"所有文件: {files}")
    
    # 创建测试按钮
    buttons = [
        ("批量添加", test_add_files, "primary"),
        ("清空所有", test_clear_all, "danger"),
        ("当前文件", test_get_current_files, "info"),
        ("所有文件", test_get_all_files, "success")
    ]
    
    for text, command, style in buttons:
        btn = tb.Button(
            test_frame,
            text=text,
            command=command,
            bootstyle=style,
            width=12
        )
        btn.pack(side=tk.LEFT, padx=5)
    
    # 运行测试
    logger.info("启动选项卡式批量文件列表组件测试")
    root.mainloop()
