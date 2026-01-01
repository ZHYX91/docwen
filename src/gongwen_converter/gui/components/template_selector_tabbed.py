"""
选项卡式模板选择组件
"""

import logging
from typing import List, Dict, Callable, Optional, Tuple

from gongwen_converter.gui.components.template_selector import TemplateSelector
from gongwen_converter.template.loader import TemplateLoader

import ttkbootstrap as tb
from ttkbootstrap.constants import *

logger = logging.getLogger(__name__)

class TabbedTemplateSelector(tb.Frame):
    """
    选项卡式模板选择组件
    """
    
    def __init__(self, master, 
                 on_template_selected: Optional[Callable] = None,
                 on_tab_changed: Optional[Callable] = None,
                 **kwargs):
        logger.debug("初始化选项卡式模板选择组件")
        
        super().__init__(master, **kwargs)
        
        self.on_template_selected = on_template_selected
        self.on_tab_changed = on_tab_changed
        
        self.tabs: Dict[str, tb.Frame] = {}
        self.template_lists: Dict[str, TemplateSelector] = {}
        self.template_loader = TemplateLoader()
        
        self.current_tab: Optional[str] = None
        
        self._create_widgets()
        self._load_templates()
        
        logger.info("选项卡式模板选择组件初始化完成")
    
    def _create_widgets(self):
        logger.debug("创建选项卡式模板选择界面元素")
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        self.notebook = tb.Notebook(self, bootstyle="info")
        self.notebook.grid(row=0, column=0, sticky="nsew")
        
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
        self._create_tabs()
        
        logger.debug("选项卡式模板选择界面元素创建完成")
    
    def _create_tabs(self):
        logger.debug("创建模板类别选项卡")
        
        categories = {'docx': "文档模板", 'xlsx': "表格模板"}
        
        for category, display_name in categories.items():
            tab_frame = tb.Frame(self.notebook, bootstyle="info")
            
            template_list = TemplateSelector(
                tab_frame,
                template_type=category,
                on_template_selected=lambda name, cat=category: self._on_template_selected(cat, name)
            )
            
            self.tabs[category] = tab_frame
            self.template_lists[category] = template_list
            
            self.notebook.add(tab_frame, text=f"  {display_name}  ")
        
        if categories:
            self.current_tab = list(categories.keys())[0]
            self.notebook.select(0)

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
        logger.debug("检查模板列表是否需要刷新")
        
        for category, template_list in self.template_lists.items():
            # 获取最新的模板列表
            new_templates = self.template_loader.get_available_templates(category)
            current_templates = template_list.templates
            
            # 比较新旧列表（使用集合比较，忽略顺序）
            if set(new_templates) != set(current_templates):
                logger.info(f"模板列表已变化 [{category}]: {len(current_templates)} -> {len(new_templates)}")
                
                # 记录当前选中的模板
                current_selected = template_list.get_selected()
                
                # 更新模板列表
                template_list.add_templates(new_templates, auto_select_first=False)
                
                # 尝试恢复之前选中的模板，如果不存在则选择第一个
                if current_selected and current_selected in new_templates:
                    template_list._select_template(current_selected)
                elif new_templates:
                    template_list._select_template(new_templates[0])
            else:
                logger.debug(f"模板列表无变化 [{category}]，跳过刷新")
        
        logger.debug("模板列表刷新检查完成")

    def _on_tab_changed(self, event):
        selected_index = self.notebook.index("current")
        categories = list(self.tabs.keys())
        
        if 0 <= selected_index < len(categories):
            new_tab = categories[selected_index]
            old_tab = self.current_tab
            self.current_tab = new_tab
            
            logger.debug(f"选项卡切换: {old_tab} -> {new_tab}")
            
            # Ensure a template is selected
            current_list = self.template_lists[new_tab]
            if not current_list.get_selected() and current_list.templates:
                 current_list._select_template(current_list.templates[0])
            else:
                # Manually trigger callback if already selected
                self._on_template_selected(new_tab, current_list.get_selected())

            if self.on_tab_changed:
                self.on_tab_changed(new_tab, old_tab)

    def _on_template_selected(self, category: str, template_name: str):
        logger.debug(f"模板被选中: {category}/{template_name}")
        if self.on_template_selected:
            self.on_template_selected(category, template_name)

    def get_selected_template(self) -> Optional[Tuple[str, str]]:
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
                template_list._select_template(template_list.templates[0])
    
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
