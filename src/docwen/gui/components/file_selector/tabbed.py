"""
选项卡式批量文件列表组件

将批量文件列表分为5个选项卡：文本类、文档类、表格类、版式类、图片类。
每个选项卡对应一个 FileSelector 实例，支持按文件类别管理批量文件。

主要功能：
- 5个文件类别选项卡（text/document/spreadsheet/layout/image）
- 智能文件分类（基于实际文件格式检测）
- 选项卡切换回调
- 批量添加时自动激活最优选项卡
- 文件数量显示在选项卡标题
"""

import contextlib
import logging
import tkinter as tk
import tkinter.font as tkfont
from collections.abc import Callable
from tkinter import ttk
from typing import TYPE_CHECKING, Optional

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.i18n import t
from docwen.utils.dpi_utils import scale
from docwen.utils.file_type_utils import activate_optimal_tab, get_actual_file_category, get_category_name

# 类型检查时导入（避免循环依赖）
if TYPE_CHECKING:
    from . import FileSelector

logger = logging.getLogger(__name__)


class TabbedFileSelector(tb.Frame):
    """
    选项卡式批量文件列表组件

    使用 ttk.Notebook 实现5个选项卡，每个选项卡对应一个文件类别。

    属性:
        notebook: ttk.Notebook 控件
        tabs: 选项卡框架字典 {category: Frame}
        file_lists: 文件列表字典 {category: FileSelector}
        current_tab: 当前激活的选项卡类别

    回调函数:
        on_file_removed: 文件移除回调
        on_file_opened: 文件打开回调
        on_tab_changed: 选项卡切换回调
        on_file_selected: 文件选中回调
        on_files_added: 文件添加完成回调
    """

    def __init__(
        self,
        master,
        on_file_removed: Callable | None = None,
        on_file_opened: Callable | None = None,
        on_tab_changed: Callable | None = None,
        on_file_selected: Callable | None = None,
        on_files_added: Callable | None = None,
        on_list_cleared: Callable | None = None,
        **kwargs,
    ):
        """
        初始化选项卡式批量文件列表组件

        参数:
            master: 父组件
            on_file_removed: 文件移除回调函数(file_path)
            on_file_opened: 文件打开回调函数(file_path)
            on_tab_changed: 选项卡切换回调函数(new_tab, old_tab)
            on_file_selected: 文件选中回调函数(file_info)
            on_files_added: 文件添加完成回调函数(added_files, failed_files)
            on_list_cleared: 列表清空回调函数()
        """
        logger.debug("初始化选项卡式批量文件列表组件")

        super().__init__(master, **kwargs)

        # 存储回调函数
        self.on_file_removed = on_file_removed
        self.on_file_opened = on_file_opened
        self.on_tab_changed = on_tab_changed
        self.on_file_selected = on_file_selected
        self.on_files_added = on_files_added
        self.on_list_cleared = on_list_cleared

        # 存储所有选项卡和对应的文件列表组件
        self.tabs: dict[str, tb.Frame] = {}
        self.file_lists: dict[str, FileSelector] = {}

        # 当前激活的选项卡
        self.current_tab: str | None = None
        self._hovered_tab_index: int | None = None
        self._tab_tooltip: tk.Toplevel | None = None
        self._tab_tooltip_label: tb.Label | None = None

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
        try:
            style = tb.Style.get_instance()
            tab_padding = (scale(16), scale(10))
            tab_margins = (scale(6), scale(4), scale(6), 0)
            notebook_style = self.notebook.cget("style") or ""

            tab_styles = {"TNotebook.Tab", "warning.TNotebook.Tab"}
            notebook_styles = {"TNotebook", "warning.TNotebook"}
            if notebook_style:
                tab_styles.add(f"{notebook_style}.Tab")
                notebook_styles.add(notebook_style)

            for s in tab_styles:
                style.configure(s, padding=tab_padding)
            for s in notebook_styles:
                style.configure(s, tabmargins=tab_margins)
        except Exception:
            pass

        # 绑定选项卡切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self.notebook.bind("<Motion>", self._on_notebook_motion, add="+")
        self.notebook.bind("<Leave>", self._on_notebook_leave, add="+")
        self.notebook.bind("<Configure>", self._on_notebook_configure, add="+")

        # 创建5个选项卡
        self._create_tabs()

        logger.debug("选项卡式批量文件列表界面元素创建完成")

    def _create_tabs(self):
        """创建5个选项卡"""
        logger.debug("创建5个文件类别选项卡")

        # 延迟导入 FileSelector 避免循环依赖
        from . import FileSelector

        # 5个文件类别
        categories = ["text", "document", "spreadsheet", "layout", "image"]

        for category in categories:
            # 创建选项卡框架
            tab_frame = tb.Frame(self.notebook, bootstyle="default")

            # 创建对应的文件列表组件
            file_list = FileSelector(
                tab_frame,
                on_file_removed=self._on_file_removed,
                on_file_opened=self.on_file_opened,
                on_list_cleared=lambda cat=category: self._on_list_cleared(cat),
                on_file_selected=self._on_file_selected,
            )

            # 存储选项卡和文件列表
            self.tabs[category] = tab_frame
            self.file_lists[category] = file_list

            # 添加选项卡到notebook
            display_name = get_category_name(category, t_func=t)
            self.notebook.add(tab_frame, text=display_name)

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

            if self.on_tab_changed:
                self.on_tab_changed(new_tab, old_tab)

    def _on_file_removed(self, file_path: str):
        """处理文件移除事件"""
        logger.debug(f"文件被移除: {file_path}")

        # 找到文件所属的类别并更新选项卡标题
        for category, file_list in self.file_lists.items():
            files = file_list.get_files()
            if file_path in files:
                self.after(10, lambda cat=category: self._update_tab_title(cat))
                break

        if self.on_file_removed:
            self.on_file_removed(file_path)

    def _update_tab_title(self, category: str):
        """
        更新选项卡标题

        - 有文件时：显示 " 类别名 (数量) "
        - 无文件时：显示 "  类别名  "

        参数:
            category: 文件类别
        """
        if category not in self.tabs:
            return

        file_count = self.get_file_count(category)
        display_name = get_category_name(category, t_func=t)
        notebook_width = 0
        try:
            notebook_width = int(self.notebook.winfo_width())
        except Exception:
            notebook_width = 0

        narrow = notebook_width > 0 and notebook_width < scale(360)
        short_display_name = t(f"file_category.{category}_short", default=display_name)

        title = short_display_name if narrow else f"{display_name} ({file_count})" if file_count > 0 else display_name

        categories = list(self.tabs.keys())
        tab_index = categories.index(category)
        self.notebook.tab(tab_index, text=title)

        logger.debug(f"更新选项卡标题: {category} -> {title.strip()}")

    def _get_tab_full_text(self, category: str) -> str:
        display_name = get_category_name(category, t_func=t)
        file_count = self.get_file_count(category)
        if file_count > 0:
            return f"{display_name} ({file_count})"
        return display_name

    def _get_tab_label_font(self) -> tkfont.Font:
        try:
            font = ttk.Style().lookup("TNotebook.Tab", "font")
            if font:
                return tkfont.Font(font=font)
        except Exception:
            pass
        try:
            return tkfont.nametofont("TkDefaultFont")
        except Exception:
            return tkfont.Font()

    def _show_tab_tooltip(self, x_root: int, y_root: int, text: str):
        if not text:
            self._hide_tab_tooltip()
            return
        if self._tab_tooltip and self._tab_tooltip.winfo_exists():
            try:
                if self._tab_tooltip_label:
                    self._tab_tooltip_label.configure(text=text)
                self._tab_tooltip.wm_geometry(f"+{x_root}+{y_root}")
                return
            except Exception:
                self._hide_tab_tooltip()

        tw = tk.Toplevel(self.notebook)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x_root}+{y_root}")
        label = tb.Label(tw, text=text, justify="left", padding=(scale(8), scale(6)), bootstyle="secondary")
        label.pack()
        self._tab_tooltip = tw
        self._tab_tooltip_label = label

    def _hide_tab_tooltip(self):
        if self._tab_tooltip and self._tab_tooltip.winfo_exists():
            with contextlib.suppress(Exception):
                self._tab_tooltip.destroy()
        self._tab_tooltip = None
        self._tab_tooltip_label = None
        self._hovered_tab_index = None

    def _on_notebook_motion(self, event):
        try:
            tab_index = self.notebook.index(f"@{event.x},{event.y}")
        except Exception:
            self._hide_tab_tooltip()
            return

        if tab_index != self._hovered_tab_index:
            self._hovered_tab_index = tab_index

        try:
            bbox = self.notebook.bbox(tab_index)
        except Exception:
            self._hide_tab_tooltip()
            return

        if not bbox:
            self._hide_tab_tooltip()
            return

        categories = list(self.tabs.keys())
        if not (0 <= tab_index < len(categories)):
            self._hide_tab_tooltip()
            return

        category = categories[tab_index]
        full_text = self._get_tab_full_text(category)
        font = self._get_tab_label_font()
        text_width = font.measure(full_text)
        available_width = max(1, int(bbox[2]) - scale(24))
        clipped = text_width > available_width
        narrow = int(self.notebook.winfo_width()) < scale(360)

        if narrow and self.get_file_count(category) > 0:
            clipped = True

        if clipped:
            self._show_tab_tooltip(event.x_root + scale(12), event.y_root + scale(18), full_text)
        else:
            self._hide_tab_tooltip()

    def _on_notebook_leave(self, _event):
        self._hide_tab_tooltip()

    def _on_notebook_configure(self, _event):
        self.update_all_tab_titles()
        self._hide_tab_tooltip()

    def update_all_tab_titles(self):
        """
        更新所有选项卡的标题

        用于模式切换等需要批量更新标题的场景。
        """
        logger.debug("批量更新所有选项卡标题")
        for category in self.file_lists:
            self._update_tab_title(category)
        logger.debug("所有选项卡标题更新完成")

    def _on_list_cleared(self, category: str):
        """处理列表清空事件"""
        logger.debug(f"选项卡 {category} 列表已清空")
        self._update_tab_title(category)

    def _on_file_selected(self, file_info):
        """处理文件选中事件，转发到外部回调"""
        logger.debug(f"转发文件选中事件: {file_info.file_path}")

        if hasattr(self, "on_file_selected") and self.on_file_selected:
            self.on_file_selected(file_info)

    def add_file(self, file_path: str) -> tuple[bool, str]:
        """
        添加文件到对应的选项卡

        参数:
            file_path: 文件路径

        返回:
            Tuple[bool, str]: (是否成功添加, 错误消息)
        """
        logger.debug(f"添加文件到选项卡式列表: {file_path}")

        category = get_actual_file_category(file_path)
        if category == "unknown":
            return False, t("messages.unsupported_file_type")

        if category in self.file_lists:
            success, error_msg = self.file_lists[category].add_file(file_path)
            if success:
                logger.info(f"文件已添加到 {category} 选项卡: {file_path}")
                self._update_tab_title(category)
            return success, error_msg
        else:
            logger.error(f"未知的文件类别: {category}")
            return False, t("messages.unknown_file_category", category=category)

    def add_files(self, file_paths: list[str]) -> tuple[list[str], list[tuple[str, str]]]:
        """
        批量添加文件到对应的选项卡

        基于实际文件格式进行分类，避免伪装后缀名问题。

        参数:
            file_paths: 文件路径列表

        返回:
            Tuple[List[str], List[Tuple[str, str]]]:
                (成功添加的文件列表, 失败的文件和错误消息列表)
        """
        logger.debug(f"批量添加文件到选项卡式列表: {len(file_paths)} 个文件")

        added_files = []
        failed_files = []

        # 统计各类别文件数量（基于实际文件格式）
        category_count = {"text": 0, "document": 0, "spreadsheet": 0, "layout": 0, "image": 0}
        categorized_files = {"text": [], "document": [], "spreadsheet": [], "layout": [], "image": []}

        # 基于实际文件格式分类文件
        for file_path in file_paths:
            category = get_actual_file_category(file_path)
            if category in categorized_files:
                category_count[category] += 1
                categorized_files[category].append(file_path)
            else:
                failed_files.append((file_path, t("messages.unsupported_file_type")))

        # 添加文件到对应选项卡
        is_first_batch = True
        affected_categories = []
        success_count = dict.fromkeys(categorized_files.keys(), 0)

        for category, files in categorized_files.items():
            if files:
                auto_select_first = is_first_batch
                success_files, failed_list = self.file_lists[category].add_files(
                    files, auto_select_first=auto_select_first, known_category=category
                )
                added_files.extend(success_files)
                failed_files.extend(failed_list)

                if success_files:
                    success_count[category] += len(success_files)
                    is_first_batch = False
                    affected_categories.append(category)

        # 更新所有受影响类别的选项卡标题
        for category in affected_categories:
            self._update_tab_title(category)

        # 使用统一算法激活最优选项卡
        if added_files:
            self._activate_optimal_tab_by_count(success_count)

        logger.info(f"批量添加完成: {len(added_files)}/{len(file_paths)} 个文件成功")

        # 触发文件添加完成回调
        if added_files and self.on_files_added:
            self.on_files_added(added_files, failed_files)
            logger.debug(f"触发文件添加完成回调: {len(added_files)} 个文件")

        return added_files, failed_files

    def _activate_optimal_tab(self, file_paths: list[str]):
        """
        使用统一算法激活最优选项卡

        基于实际文件格式检测，避免伪装后缀名问题。

        参数:
            file_paths: 文件路径列表
        """
        if not file_paths:
            return

        optimal_category = activate_optimal_tab(file_paths)

        if optimal_category in self.tabs:
            categories = list(self.tabs.keys())
            tab_index = categories.index(optimal_category)
            self.notebook.select(tab_index)
            self.current_tab = optimal_category
            logger.info(f"自动激活最优选项卡: {optimal_category} (文件数: {len(file_paths)})")
        else:
            logger.warning(f"无法激活未知的选项卡类别: {optimal_category}")

    def _activate_optimal_tab_by_count(self, category_count: dict[str, int]):
        if not category_count:
            return

        max_count = max(category_count.values())
        if max_count <= 0:
            return

        candidates = [cat for cat, count in category_count.items() if count == max_count]
        priority_order = ["text", "document", "spreadsheet", "layout", "image"]
        optimal_category = None
        for cat in priority_order:
            if cat in candidates:
                optimal_category = cat
                break
        if not optimal_category:
            optimal_category = candidates[0]

        if optimal_category in self.tabs:
            categories = list(self.tabs.keys())
            tab_index = categories.index(optimal_category)
            self.notebook.select(tab_index)
            self.current_tab = optimal_category
            logger.info(f"自动激活最优选项卡: {optimal_category} (文件数: {max_count})")

    def get_current_category(self) -> str | None:
        """
        获取当前激活的选项卡类别

        返回:
            Optional[str]: 当前激活的选项卡类别
        """
        return self.current_tab

    def get_current_file_list(self) -> Optional["FileSelector"]:
        """
        获取当前激活选项卡对应的文件列表对象

        返回:
            Optional[FileSelector]: 当前激活选项卡的文件列表对象
        """
        if self.current_tab and self.current_tab in self.file_lists:
            return self.file_lists[self.current_tab]
        return None

    def get_current_files(self) -> list[str]:
        """
        获取当前激活选项卡中的文件列表

        返回:
            List[str]: 当前激活选项卡中的文件路径列表
        """
        if self.current_tab and self.current_tab in self.file_lists:
            return self.file_lists[self.current_tab].get_files()
        return []

    def get_files_by_category(self, category: str) -> list[str]:
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

    def get_all_files(self) -> list[str]:
        """
        获取所有选项卡中的文件列表

        返回:
            List[str]: 所有文件路径列表
        """
        all_files = []
        for file_list in self.file_lists.values():
            all_files.extend(file_list.get_files())
        return all_files

    def get_file_count(self, category: str | None = None) -> int:
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

    def has_files(self, category: str | None = None) -> bool:
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
            return any(file_list.has_files() for file_list in self.file_lists.values())

    def clear_all(self):
        """清空所有选项卡的文件"""
        logger.debug("清空所有选项卡的文件")

        for file_list in self.file_lists.values():
            file_list.clear_all_files()

        for category in self.file_lists:
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
            self._update_tab_title(category)
            logger.info(f"{category} 选项卡的文件已清空")
        else:
            logger.warning(f"未知的文件类别: {category}")

    def update_file_status(self, file_path: str, status: str, output_path: str | None = None):
        """
        更新文件状态

        参数:
            file_path: 文件路径
            status: 新状态
            output_path: 输出文件路径（可选）
        """
        logger.debug(f"更新文件状态: {file_path} -> {status}")

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
            Optional[FileInfo]: 选中的文件信息
        """
        if self.current_tab and self.current_tab in self.file_lists:
            return self.file_lists[self.current_tab].get_selected_file()
        return None

    def get_selected_file_path(self) -> str | None:
        """
        获取当前激活选项卡中选中文件的路径

        返回:
            Optional[str]: 选中文件的路径
        """
        selected_file = self.get_selected_file()
        return selected_file.file_path if selected_file else None
