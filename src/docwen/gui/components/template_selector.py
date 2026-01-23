"""
模板列表组件

提供模板选择功能，支持：
- 模板列表显示和选择
- 模板文件位置快速打开
- 空状态和选中状态管理
- 与模板加载器集成

使用ttkbootstrap进行界面美化和样式管理。
"""

import os
import logging
import subprocess
import tkinter as tk
from typing import List, Dict, Callable, Optional, Tuple
from docwen.utils.font_utils import get_default_font, get_small_font, get_micro_font
from docwen.utils.icon_utils import load_image_icon
from docwen.i18n import t

import ttkbootstrap as tb
from ttkbootstrap.constants import *

logger = logging.getLogger(__name__)

class TemplateSelector:
    """
    模板列表组件
    """
    
    def __init__(self, master, template_type: str, on_template_selected: Optional[Callable] = None, **kwargs):
        logger.debug(f"初始化模板列表组件 (类型: {template_type})")
        
        self.master = master
        self.template_type = template_type
        self.on_template_selected = on_template_selected
        
        self.templates: List[str] = []
        self.selected_template: Optional[str] = None
        
        self.default_font, self.default_size = get_default_font()
        self.small_font, self.small_size = get_small_font()
        self.micro_font, self.micro_size = get_micro_font()
        
        self._create_widgets()
        
        logger.info("模板列表组件初始化完成")
    
    def _create_widgets(self):
        logger.debug("创建模板列表界面元素")
        
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=0)
        
        self._create_template_list_area()
        
        logger.debug("模板列表界面元素创建完成")
    
    def _create_template_list_area(self):
        logger.debug("创建模板列表区域")
        
        self.canvas = tk.Canvas(
            self.master, 
            bg="SystemButtonFace",
            highlightthickness=0
        )
        
        # 自定义滚动命令（带边界检查，防止滚动超出内容区域）
        def bounded_yview(*args):
            """带边界检查的滚动命令"""
            if args[0] == 'scroll':
                # 点击滚动条按钮时的滚动操作：检查是否已到边界
                top, bottom = self.canvas.yview()
                direction = int(args[1])
                if top <= 0 and direction < 0:  # 已在顶部，不能再向上滚动
                    return
                if bottom >= 1 and direction > 0:  # 已在底部，不能再向下滚动
                    return
            self.canvas.yview(*args)
        
        self.scrollbar = tb.Scrollbar(
            self.master, 
            orient=tk.VERTICAL, 
            command=bounded_yview,
            bootstyle="round-info"
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.inner_frame = tb.Frame(self.canvas, bootstyle="default")
        
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.inner_frame, anchor="nw"
        )
        
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.inner_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        
        self._show_empty_state()
    
    def _on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
    
    def _show_empty_state(self):
        logger.debug("显示空状态")
        
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        
        from docwen.utils.dpi_utils import scale
        empty_padding = scale(20)
        empty_label = tb.Label(
            self.inner_frame,
            text=t("components.template_selector.no_templates"),
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            justify=tk.CENTER,
            anchor=tk.CENTER,
            padding=empty_padding
        )
        empty_label.pack(expand=True, fill=tk.BOTH)
    
    def _show_template_list(self):
        logger.debug("显示模板列表")
        
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        
        for template_name in self.templates:
            self._create_template_item(template_name)
    
    def _create_template_item(self, template_name: str):
        from docwen.utils.dpi_utils import scale
        
        is_selected = self.selected_template == template_name
        
        icon_size = scale(16)
        border_width = scale(2)
        item_padx = scale(2)
        item_pady = scale(2)
        button_padx = scale(2)

        if is_selected:
            item_frame = tk.Frame(self.inner_frame)
            try:
                style = tb.Style.get_instance()
                info_color = style.colors.info
                parent_bg = style.colors.bg
                item_frame.configure(
                    bg=parent_bg,
                    highlightbackground=info_color,
                    highlightcolor=info_color,
                    highlightthickness=border_width
                )
            except Exception as e:
                item_frame.configure(
                    bg='SystemButtonFace',
                    highlightthickness=border_width
                )
        else:
            item_frame = tb.Frame(self.inner_frame, bootstyle="default")
        
        item_frame.pack(fill=tk.X, padx=item_padx, pady=item_pady)
        
        item_frame.grid_columnconfigure(0, weight=1)
        item_frame.grid_columnconfigure(1, weight=0)
        
        name_label = tb.Label(
            item_frame,
            text=template_name,
            font=(self.small_font, self.small_size),
            bootstyle="info" if is_selected else "default",
            anchor=tk.W,
            cursor="hand2"
        )
        name_label.grid(row=0, column=0, sticky="ew", padx=(5, 5))
        name_label.bind("<Button-1>", lambda e, name=template_name: self._select_template(name))

        location_icon = load_image_icon("location_icon.png", master=item_frame, size=(icon_size, icon_size))
        location_btn = tb.Button(
            item_frame,
            image=location_icon,
            bootstyle="secondary-link",
            command=lambda name=template_name: self._open_template_location(name),
            cursor="hand2"
        )
        if location_icon:
            location_btn.image = location_icon
        location_btn.grid(row=0, column=1, sticky="e", padx=button_padx)

    def _select_template(self, template_name: str):
        logger.debug(f"选中模板: {template_name}")
        
        if self.selected_template == template_name:
            return

        self.selected_template = template_name
        self._show_template_list()
        
        if self.on_template_selected:
            self.on_template_selected(template_name)
        
        logger.info(f"模板已选中: {template_name}")

    def add_templates(self, template_names: List[str], auto_select_first: bool = False):
        logger.debug(f"批量添加模板: {len(template_names)} 个")
        
        self.templates.clear()
        self.templates.extend(template_names)
        
        if not self.templates:
            self._show_empty_state()
            return

        self._show_template_list()

        if auto_select_first and self.templates:
            self._select_template(self.templates[0])

    def clear_all(self):
        logger.debug("清空模板列表")
        self.templates.clear()
        self.selected_template = None
        self._show_empty_state()
        logger.info("模板列表已清空")

    def get_selected(self) -> Optional[str]:
        return self.selected_template

    def _open_template_location(self, template_name: str):
        logger.debug(f"打开模板位置: {self.template_type}/{template_name}")
        
        try:
            from docwen.template.loader import TemplateLoader
            loader = TemplateLoader()
            template_path = loader.get_template_path(self.template_type, template_name)
            
            subprocess.Popen(['explorer', '/select,', os.path.normpath(template_path)])
            logger.info(f"已打开模板位置: {template_path}")
            
        except Exception as e:
            logger.error(f"打开模板位置失败: {str(e)}")
