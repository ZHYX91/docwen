"""
选项卡式模板选择组件

该组件提供了一个选项卡式的模板选择界面，支持文档模板和表格模板的分类选择。
使用ttkbootstrap进行界面美化和样式管理，支持模板的动态加载和刷新。
"""

import contextlib
import logging
import threading
import tkinter as tk
import tkinter.font as tkfont
from collections.abc import Callable
from tkinter import ttk

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.gui.components.template_selector import TemplateSelector
from docwen.i18n import t
from docwen.template.loader import TemplateLoader
from docwen.utils.dpi_utils import scale

logger = logging.getLogger(__name__)


class TabbedTemplateSelector(tb.Frame):
    """
    选项卡式模板选择组件
    """

    def __init__(
        self,
        master,
        on_template_selected: Callable | None = None,
        on_tab_changed: Callable | None = None,
        **kwargs,
    ):
        logger.debug("初始化选项卡式模板选择组件")

        super().__init__(master, **kwargs)

        self.on_template_selected = on_template_selected
        self.on_tab_changed = on_tab_changed

        self.tabs: dict[str, tb.Frame] = {}
        self.template_lists: dict[str, TemplateSelector] = {}
        self.template_loader = TemplateLoader()

        self.current_tab: str | None = None
        self._hovered_tab_index: int | None = None
        self._tab_tooltip: tk.Toplevel | None = None
        self._tab_tooltip_label: tb.Label | None = None
        self._full_tab_names: dict[str, str] = {}
        self._short_tab_names: dict[str, str] = {}
        self._refresh_inflight: bool = False

        self._create_widgets()
        self._load_templates()

        logger.info("选项卡式模板选择组件初始化完成")

    def _create_widgets(self):
        logger.debug("创建选项卡式模板选择界面元素")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.notebook = tb.Notebook(self, bootstyle="info")
        self.notebook.grid(row=0, column=0, sticky="nsew")
        try:
            style = tb.Style.get_instance()
            tab_padding = (scale(16), scale(10))
            tab_margins = (scale(6), scale(4), scale(6), 0)
            notebook_style = self.notebook.cget("style") or ""

            tab_styles = {"TNotebook.Tab", "info.TNotebook.Tab"}
            notebook_styles = {"TNotebook", "info.TNotebook"}
            if notebook_style:
                tab_styles.add(f"{notebook_style}.Tab")
                notebook_styles.add(notebook_style)

            for s in tab_styles:
                style.configure(s, padding=tab_padding)
            for s in notebook_styles:
                style.configure(s, tabmargins=tab_margins)
        except Exception:
            pass

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self.notebook.bind("<Motion>", self._on_notebook_motion, add="+")
        self.notebook.bind("<Leave>", self._on_notebook_leave, add="+")
        self.notebook.bind("<Configure>", self._on_notebook_configure, add="+")

        self._create_tabs()

        logger.debug("选项卡式模板选择界面元素创建完成")

    def _create_tabs(self):
        """
        创建模板类别选项卡

        根据模板类型创建对应的选项卡，每个选项卡包含一个模板选择器。
        支持文档模板和表格模板两种类型。
        """
        logger.debug("创建模板类别选项卡")

        # 定义模板类别和对应的显示名称（使用翻译函数）
        categories = {
            "docx": t("components.template_selector_tabbed.document_templates"),
            "xlsx": t("components.template_selector_tabbed.spreadsheet_templates"),
        }
        self._full_tab_names = dict(categories)
        self._short_tab_names = {
            "docx": t("components.template_selector_tabbed.document_short", default=categories["docx"]),
            "xlsx": t("components.template_selector_tabbed.spreadsheet_short", default=categories["xlsx"]),
        }

        for category, display_name in categories.items():
            # 创建选项卡框架
            tab_frame = tb.Frame(self.notebook, bootstyle="info")

            # 创建模板选择器组件
            template_list = TemplateSelector(
                tab_frame,
                template_type=category,
                on_template_selected=lambda name, cat=category: self._on_template_selected(cat, name),
            )

            # 存储选项卡和模板选择器引用
            self.tabs[category] = tab_frame
            self.template_lists[category] = template_list

            # 添加选项卡到笔记本组件
            self.notebook.add(tab_frame, text=display_name)

        # 设置默认选中的选项卡
        if categories:
            self.current_tab = next(iter(categories.keys()))
            self.notebook.select(0)
            logger.debug(f"默认选中选项卡: {self.current_tab}")

    def _is_narrow(self) -> bool:
        try:
            width = int(self.notebook.winfo_width())
        except Exception:
            width = 0
        return width > 0 and width < scale(300)

    def _update_tab_titles(self):
        narrow = self._is_narrow()
        categories = list(self.tabs.keys())
        for index, category in enumerate(categories):
            full_name = self._full_tab_names.get(category, category)
            short_name = self._short_tab_names.get(category, full_name)
            self.notebook.tab(index, text=short_name if narrow else full_name)

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
        full_text = self._full_tab_names.get(category) or category
        font = self._get_tab_label_font()
        text_width = font.measure(full_text)
        available_width = max(1, int(bbox[2]) - scale(24))
        clipped = text_width > available_width

        if clipped:
            self._show_tab_tooltip(event.x_root + scale(12), event.y_root + scale(18), full_text)
        else:
            self._hide_tab_tooltip()

    def _on_notebook_leave(self, _event):
        self._hide_tab_tooltip()

    def _on_notebook_configure(self, _event):
        self._update_tab_titles()
        self._hide_tab_tooltip()

    def _load_templates(self):
        logger.debug("加载模板")
        for category, template_list in self.template_lists.items():
            templates = self.template_loader.get_available_templates(category)
            template_list.add_templates(templates, auto_select_first=True)

    def refresh_templates(self):
        """
        智能刷新模板列表
        仅在模板目录内容发生变化时才更新UI，避免不必要的重绘
        """
        if self._refresh_inflight:
            return

        self._refresh_inflight = True
        logger.debug("异步刷新模板列表（后台扫描）")

        def worker():
            try:
                templates_by_category: dict[str, list[str]] = {}
                for category in list(self.template_lists.keys()):
                    templates_by_category[category] = self.template_loader.get_available_templates(category)
                self.after(0, lambda: self._apply_template_refresh(templates_by_category))
            except Exception:
                self.after(0, self._finish_template_refresh)

        threading.Thread(target=worker, daemon=True).start()

    def _finish_template_refresh(self):
        self._refresh_inflight = False

    def _apply_template_refresh(self, templates_by_category: dict[str, list[str]]):
        try:
            for category, template_list in self.template_lists.items():
                new_templates = templates_by_category.get(category, [])
                current_templates = template_list.templates

                if set(new_templates) != set(current_templates):
                    logger.info(f"模板列表已变化 [{category}]: {len(current_templates)} -> {len(new_templates)}")
                    current_selected = template_list.get_selected()
                    template_list.add_templates(new_templates, auto_select_first=False)
                    if current_selected and current_selected in new_templates:
                        template_list.select_template(current_selected)
                    elif new_templates:
                        template_list.select_template(new_templates[0])
                else:
                    logger.debug(f"模板列表无变化 [{category}]，跳过刷新")
        finally:
            self._finish_template_refresh()

    def _on_tab_changed(self, event):
        selected_index = self.notebook.index("current")
        categories = list(self.tabs.keys())

        if 0 <= selected_index < len(categories):
            new_tab = categories[selected_index]
            old_tab = self.current_tab
            self.current_tab = new_tab

            logger.debug(f"选项卡切换: {old_tab} -> {new_tab}")

            # 智能刷新模板列表（仅在模板目录内容变化时更新）
            self.refresh_templates()

            # Ensure a template is selected
            current_list = self.template_lists[new_tab]
            if not current_list.get_selected() and current_list.templates:
                current_list.select_template(current_list.templates[0])
            else:
                # Manually trigger callback if already selected
                self._on_template_selected(new_tab, current_list.get_selected())

            if self.on_tab_changed:
                self.on_tab_changed(new_tab, old_tab)

    def _on_template_selected(self, category: str, template_name: str | None):
        logger.debug(f"模板被选中: {category}/{template_name}")
        if not template_name:
            return
        if self.on_template_selected:
            self.on_template_selected(category, template_name)

    def get_selected_template(self) -> tuple[str, str] | None:
        if self.current_tab and self.current_tab in self.template_lists:
            selected_name = self.template_lists[self.current_tab].get_selected()
            if selected_name:
                return (self.current_tab, selected_name)
        return None

    def show(self):
        """显示模板选择器"""
        self.grid(row=0, column=0, sticky="nsew")

    def hide(self):
        """隐藏模板选择器"""
        self.grid_remove()

    def reset(self):
        """重置组件状态"""
        logger.debug("重置选项卡式模板选择器")

        # 先刷新模板列表，确保数据是最新的
        self.refresh_templates()

        # 然后切换到第一个选项卡并选择第一个模板
        self.notebook.select(0)
        for template_list in self.template_lists.values():
            if template_list.templates:
                template_list.select_template(template_list.templates[0])

    def activate_and_select(self, template_type: str):
        """激活指定类型的选项卡并选择第一个模板"""
        logger.debug(f"激活并选择模板: {template_type}")
        tab_index = 0 if template_type == "docx" else 1
        self.notebook.select(tab_index)

        # Manually trigger tab change logic to select first item
        self._on_tab_changed(None)

    def _auto_select_first_template(self, template_type: str):
        """自动选择第一个模板"""
        self.activate_and_select(template_type)

    def _on_notebook_tab_changed(self, event):
        """Alias for _on_tab_changed for compatibility"""
        self._on_tab_changed(event)
