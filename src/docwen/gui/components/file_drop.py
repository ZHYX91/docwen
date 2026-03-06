"""
文件拖拽区域组件

该组件提供了一个用户友好的文件拖拽界面，支持单文件模式和批量模式。
使用ttkbootstrap进行界面美化和样式管理，支持文件拖拽、类型验证和动态界面切换。
"""

import logging
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog
from typing import Any, Protocol, cast

# 导入ttkbootstrap用于界面美化和样式管理
import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.formats import category_from_actual_format
from docwen.i18n import t
from docwen.utils.dpi_utils import scale
from docwen.utils.file_type_utils import (
    get_allowed_types_by_category,
    is_supported_file,
    validate_file_format,
)
from docwen.utils.font_utils import get_default_font, get_micro_font, get_small_font, get_title_font
from docwen.utils.gui_utils import bind_single_line_ellipsis
from docwen.utils.path_utils import collect_files_from_folder

# 配置日志记录器
logger = logging.getLogger(__name__)


class _DndCapable(Protocol):
    def drop_target_register(self, *dnd_types: str) -> Any: ...
    def drop_target_unregister(self) -> Any: ...
    def dnd_bind(self, sequence: str, func: Any, add: Any = ...) -> Any: ...


# 尝试导入拖拽支持库
try:
    from tkinterdnd2 import DND_FILES

    TKINTERDND2_AVAILABLE = True
    logger.debug("tkinterdnd2库已成功导入")
except ImportError:
    # 回退到标准Tkinter（功能受限）
    DND_FILES = None
    TKINTERDND2_AVAILABLE = False
    logging.warning("tkinterdnd2未安装，拖拽功能将不可用")


class FileDropArea(tb.Frame):
    """
    文件拖拽区域组件

    主要特性包括：
    - 支持单文件模式和批量模式切换
    - 集成tkinterdnd2库实现原生文件拖拽功能
    - 自动验证文件类型，只接受支持的文件格式
    - 使用ttkbootstrap统一管理样式，支持主题切换
    - 动态界面状态管理，根据文件选择状态自动更新显示
    - 提供完整的编程接口，便于与其他组件集成

    组件布局结构：
    - 顶部控制栏：模式切换器、添加按钮、清空按钮
    - 内容区域：根据状态显示空状态提示或有文件信息

    支持的交互方式：
    - 文件拖拽（需要安装tkinterdnd2）
    - 文件选择对话框
    - 编程接口设置文件路径
    """

    def __init__(
        self,
        master,
        on_file_dropped: Callable | None = None,
        on_show_batch_panel: Callable | None = None,
        on_hide_batch_panel: Callable | None = None,
        height: int = 120,
        default_mode: str = "single",
        **kwargs,
    ):
        """
        初始化文件拖拽区域组件

        参数:
            master: 父组件，必须是ttk或ttkbootstrap组件，用于布局管理
            on_file_dropped: 文件拖拽完成后的回调函数，接受文件路径作为参数
            on_show_batch_panel: 显示批量面板的回调函数，用于模式切换时的界面更新
            on_hide_batch_panel: 隐藏批量面板的回调函数，用于模式切换时的界面更新
            height: 组件高度（像素），用于控制整体尺寸
            default_mode: 默认模式，"single"（单文件）或"batch"（批量），默认为"single"
            kwargs: 传递给tb.Frame的额外参数，用于自定义样式和行为
        """
        logger.debug("初始化文件拖拽区域组件 - 使用ttkbootstrap样式")

        # 调用父类构造函数，设置基础框架
        super().__init__(master, height=height, **kwargs)

        # 阻止子组件影响父容器尺寸，强制固定高度
        # 这确保了组件在不同DPI设置下保持一致的视觉高度
        self.pack_propagate(False)

        # 记录高度设置，便于调试和布局调整
        logger.info(f"文件拖拽区域高度设置为: {height} 像素")

        # 存储回调函数，用于组件间通信
        self.on_file_dropped = on_file_dropped
        self.on_show_batch_panel = on_show_batch_panel
        self.on_hide_batch_panel = on_hide_batch_panel
        self.on_clear_clicked: Callable[[], None] | None = None

        # 存储组件状态
        self.is_dragging = False  # 拖拽状态标志，用于视觉反馈
        self.file_path = ""  # 当前选择的文件路径（单文件模式）
        self.drag_enabled = True  # 拖拽功能启用状态
        self.mode = default_mode  # 模式："single" 单文件模式 或 "batch" 批量模式，使用传入的默认值

        # 获取字体配置（优化：只在初始化时获取一次，避免重复计算）
        # 这些字体配置确保在不同DPI设置下文字显示清晰
        self.default_font, self.default_size = get_default_font()
        self.title_font, self.title_size = get_title_font()
        self.small_font, self.small_size = get_small_font()
        self.micro_font, self.micro_size = get_micro_font()

        # 配置组件样式 - 使用ttkbootstrap主题
        # "light"样式提供清晰的视觉层次和良好的可读性
        self.configure(bootstyle="light")

        # 创建界面元素，构建完整的用户界面
        self._create_widgets()

        # 设置拖拽支持，启用文件拖拽功能
        self._setup_drag_drop()

        logger.info("文件拖拽区域组件初始化完成")

    def _create_widgets(self):
        """
        创建界面元素，全部使用ttk组件

        该方法构建组件的完整视觉界面，包括：
        - 主容器和布局配置
        - 顶部控制栏（模式切换、按钮）
        - 内容区域（空状态提示和文件信息显示）
        - 模式切换器

        布局采用grid系统，确保在不同尺寸下都能正确显示
        """
        logger.debug("创建文件拖拽区域界面元素")

        # 创建主容器 - 使用ttkbootstrap卡片样式，使用grid布局
        # padding参数根据DPI自动缩放，确保在不同显示器上都有合适的间距
        self.container = tb.Frame(self, bootstyle="default", padding=scale(5))
        self.container.pack(expand=True, fill=tk.BOTH)

        # 配置grid布局权重，实现灵活的响应式布局
        # 第0行：控制栏（固定高度），第1行：内容区域（占据剩余空间）
        self.container.grid_rowconfigure(0, weight=0)  # 顶部控制栏（固定高度）
        self.container.grid_rowconfigure(1, weight=1)  # 内容区域（占据剩余空间）
        self.container.grid_columnconfigure(0, weight=1)  # 单列，占据全部宽度

        # 创建顶部控制栏（固定置顶）
        # 控制栏包含模式切换、添加按钮和清空按钮，提供主要操作入口
        self.control_bar = tb.Frame(self.container, bootstyle="default")
        self.control_bar.grid(row=0, column=0, sticky="ew", pady=(0, scale(5)))

        # 配置控制栏的列权重（5列布局）
        # 这种布局确保按钮位置固定，中间有弹性空间适应不同宽度
        self.control_bar.grid_columnconfigure(0, weight=0)  # 第0列：模式切换器（固定宽度）
        self.control_bar.grid_columnconfigure(1, weight=5)  # 第1列：左侧弹性列
        self.control_bar.grid_columnconfigure(2, weight=0)  # 第2列：添加按钮（固定宽度）
        self.control_bar.grid_columnconfigure(3, weight=1)  # 第3列：右侧弹性空间
        self.control_bar.grid_columnconfigure(4, weight=0)  # 第4列：清空按钮（固定宽度）

        # 创建添加按钮（放在控制栏中间）
        # outline-primary样式提供清晰的视觉层次，与主题保持一致
        self.add_button = tb.Button(
            self.control_bar,
            text=t("components.file_drop.add_button"),
            bootstyle="outline-primary",
            command=self._on_add_button_clicked,
            width=6,
        )

        # 创建清空按钮（放在控制栏右侧，初始禁用）
        self.clear_button = tb.Button(
            self.control_bar,
            text=t("components.file_drop.clear_button"),
            bootstyle="outline-danger",
            command=self._on_clear_clicked,
            width=6,
            state="disabled",  # 初始禁用，只有在有文件时才启用
        )
        # 清空按钮始终显示在第4列，右对齐
        self.clear_button.grid(row=0, column=4, sticky="e")

        logger.debug("顶部控制栏创建完成")

        # 创建标题区域（用于无文件时显示"拖拽文件到这里"）
        # 这个区域在空状态下显示，提供用户操作指引
        self.title_frame = tb.Frame(self.container, bootstyle="default")

        # 配置标题区域grid布局实现垂直居中
        # 通过权重分配实现内容在垂直方向上的居中显示
        self.title_frame.grid_rowconfigure(0, weight=2)  # 顶部弹性空间
        self.title_frame.grid_rowconfigure(1, weight=0)  # 标题行
        self.title_frame.grid_rowconfigure(2, weight=0)  # 提示行1
        self.title_frame.grid_rowconfigure(3, weight=0)  # 提示行2
        self.title_frame.grid_rowconfigure(4, weight=0)  # 提示行3
        self.title_frame.grid_rowconfigure(5, weight=0)  # 提示行4
        self.title_frame.grid_rowconfigure(6, weight=0)  # 提示行5
        self.title_frame.grid_rowconfigure(7, weight=1)  # 底部弹性空间
        self.title_frame.grid_columnconfigure(0, weight=1)

        # 创建标题标签（根据当前模式动态显示提示）
        hint_key = (
            "components.file_drop.empty_hint_single"
            if self.mode == "single"
            else "components.file_drop.empty_hint_batch"
        )
        self.title_label = tb.Label(
            self.title_frame,
            text=t(hint_key),
            font=(self.title_font, self.title_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
        )
        self.title_label.grid(row=1, column=0, sticky="ew", pady=(0, scale(5)))

        # 创建拖拽提示区域（初始显示）
        self.prompt_frame = tb.Frame(self.title_frame, bootstyle="default")
        self.prompt_frame.grid(row=2, column=0, rowspan=5, sticky="ew")

        # 第一行提示：支持文本文件
        self.text_prompt_label = tb.Label(
            self.prompt_frame,
            text=t("file_types.text") + ": " + get_allowed_types_by_category("text"),
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
            padding=(0, 3, 0, 0),
        )
        self.text_prompt_label.pack(expand=False, fill=tk.X)

        # 第二行提示：支持表格文件
        self.spreadsheet_prompt_label = tb.Label(
            self.prompt_frame,
            text=t("file_types.spreadsheet") + ": " + get_allowed_types_by_category("spreadsheet"),
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
            padding=(0, 3, 0, 0),
        )
        self.spreadsheet_prompt_label.pack(expand=False, fill=tk.X)

        # 第三行提示：支持文档文件
        self.document_prompt_label = tb.Label(
            self.prompt_frame,
            text=t("file_types.document") + ": " + get_allowed_types_by_category("document"),
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
            padding=(0, 3, 0, 0),
        )
        self.document_prompt_label.pack(expand=False, fill=tk.X)

        # 第四行提示：支持图片文件
        self.image_prompt_label = tb.Label(
            self.prompt_frame,
            text=t("file_types.image") + ": " + get_allowed_types_by_category("image"),
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
            padding=(0, 3, 0, 0),
        )
        self.image_prompt_label.pack(expand=False, fill=tk.X)

        # 第五行提示：支持版式文件
        self.layout_prompt_label = tb.Label(
            self.prompt_frame,
            text=t("file_types.layout") + ": " + get_allowed_types_by_category("layout"),
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            anchor=tk.CENTER,
            padding=(0, 3, 0, 0),
        )
        self.layout_prompt_label.pack(expand=False, fill=tk.X)
        logger.debug("五行拖拽提示标签创建完成")

        # 创建统一的文件信息显示区域（单文件模式和批量模式共用）
        # 这个区域在有文件时显示，根据模式展示不同的文件信息
        self.file_info_frame = tb.Frame(self.container, bootstyle="default")

        logger.debug("统一文件信息显示区域创建完成")

        # 创建模式切换器（药丸型按钮组）
        self._create_mode_switcher(self.mode)

        # 设置初始状态为空状态
        self._switch_to_empty_state()
        logger.info("文件拖拽区域界面元素创建完成")

    def _create_mode_switcher(self, default_mode: str = "single"):
        """
        创建模式切换器（药丸型按钮组）

        模式切换器采用药丸型设计，提供清晰的视觉反馈：
        - 批量模式：绿色主题，表示多文件处理
        - 单文件模式：蓝色主题，表示单文件处理

        这种设计帮助用户快速理解当前模式并轻松切换

        参数:
            default_mode: 默认模式，"single" 或 "batch"
        """
        logger.debug(f"创建模式切换器，默认模式: {default_mode}")

        # 创建容器框架（放在控制栏左侧）
        self.mode_frame = tb.Frame(self.control_bar, bootstyle="default")
        self.mode_frame.grid(row=0, column=0, sticky="w")

        # 模式变量，用于跟踪当前选择的模式，使用传入的默认值
        self.mode_var = tk.StringVar(value=default_mode)

        # 批量按钮（放在左边）
        # success-toolbutton样式使用绿色主题，表示积极操作
        self.batch_btn = tb.Radiobutton(
            self.mode_frame,
            text=t("components.file_drop.batch_mode"),
            variable=self.mode_var,
            value="batch",
            bootstyle="success-toolbutton",
            command=self._on_mode_changed,
            width=8,
        )
        self.batch_btn.pack(side=tk.LEFT)

        # 单文件按钮（放在右边）
        # info-toolbutton样式使用蓝色主题，表示信息性操作
        self.single_btn = tb.Radiobutton(
            self.mode_frame,
            text=t("components.file_drop.single_mode"),
            variable=self.mode_var,
            value="single",
            bootstyle="info-toolbutton",
            command=self._on_mode_changed,
            width=8,
        )
        self.single_btn.pack(side=tk.LEFT)

        logger.debug("模式切换器创建完成")

    def _on_mode_changed(self):
        """
        处理模式切换事件

        当用户切换模式时，该方法负责：
        1. 更新内部状态
        2. 通知相关组件模式变化
        3. 更新界面显示
        4. 记录日志用于调试

        模式切换会影响文件处理逻辑和界面显示方式
        """
        new_mode = self.mode_var.get()
        logger.info(f"模式切换: {self.mode} -> {new_mode}")

        # 更新模式状态
        old_mode = self.mode
        self.mode = new_mode

        # 通知主窗口切换面板
        if new_mode == "batch":
            # 单→批：显示批量面板
            logger.debug("单→批：显示批量面板")
            if self.on_show_batch_panel:
                self.on_show_batch_panel()
        else:
            # 批→单：隐藏批量面板
            logger.debug("批→单：隐藏批量面板")
            if self.on_hide_batch_panel:
                self.on_hide_batch_panel()

        # 调用TabbedFileManager的模式切换回调
        # 确保文件管理器组件也响应模式变化
        self._call_tabbed_file_manager_on_mode_changed(new_mode)

        # 更新提示文本
        hint_key = (
            "components.file_drop.empty_hint_single"
            if new_mode == "single"
            else "components.file_drop.empty_hint_batch"
        )
        self.title_label.configure(text=t(hint_key))

        # 更新拖拽区显示
        # 根据新模式重新构建文件信息显示
        file_manager = self._get_tabbed_file_manager()
        if file_manager:
            selected_file = file_manager.selected_file
            self.update_display(selected_file)
            logger.debug(f"拖拽区显示已更新，选中文件: {selected_file.file_path if selected_file else 'None'}")

        logger.info(f"模式已切换: {old_mode} -> {new_mode}")

    def update_display(self, selected_file_info=None):
        """
        根据当前模式和文件信息更新显示

        这是外部调用的公共接口，用于在模式切换或文件选择变化时更新拖拽区显示。
        该方法确保界面始终与底层数据状态保持一致。

        参数:
            selected_file_info: 选中的文件信息对象（FileInfo类型），包含file_path等属性
                             在单文件模式下使用，批量模式下忽略此参数
        """
        logger.debug(
            f"update_display 被调用，模式: {self.mode}, 文件: {selected_file_info.file_path if selected_file_info else 'None'}"
        )

        if self.mode == "single":
            # 单文件模式：根据选中文件更新显示
            if selected_file_info and hasattr(selected_file_info, "file_path"):
                self.file_path = selected_file_info.file_path
                self._switch_to_file_state()
                logger.info(f"单文件模式：显示选中文件 {self.file_path}")
            else:
                self.file_path = ""
                self._switch_to_empty_state()
                logger.info("单文件模式：无选中文件，显示空状态")
        else:
            # 批量模式：根据批量文件列表更新显示
            batch_list = self._get_tabbed_batch_file_list()
            if batch_list and batch_list.has_files():
                self._switch_to_file_state()
                logger.info("批量模式：显示批量信息")
            else:
                self._switch_to_empty_state()
                logger.info("批量模式：无文件，显示空状态")

    def _switch_to_empty_state(self):
        """
        切换到无文件状态

        当没有文件被选择时显示空状态界面：
        - 显示拖拽提示和文件类型说明
        - 根据是否有其他选项卡文件决定清空按钮状态
        - 隐藏文件信息区域，显示标题区域

        这种状态提供清晰的操作指引，帮助用户了解功能范围
        """
        logger.debug("切换到无文件状态")

        # 检查所有选项卡是否都没有文件（批量和单文件模式都检查）
        # 这确保清空按钮的状态准确反映整体文件状态
        batch_list = self._get_tabbed_batch_file_list()
        has_any_files = batch_list and batch_list.has_files()

        # 添加按钮始终显示在第2列
        self.add_button.grid(row=0, column=2, sticky="e")

        if has_any_files:
            # 还有其他选项卡有文件，启用清空按钮
            # 这种情况发生在切换模式时，其他模式仍有文件
            logger.debug(f"{self.mode}模式：其他选项卡仍有文件，启用清空按钮")
            self.clear_button.configure(state="normal")
            # 只调整内容区域显示，保持控制栏状态
            self.title_frame.grid(row=1, column=0, sticky="nsew")
            self.file_info_frame.grid_forget()
        else:
            # 所有选项卡都没有文件时，禁用清空按钮
            # 避免用户执行无意义的清空操作
            logger.debug(f"{self.mode}模式：所有选项卡都无文件，禁用清空按钮")
            self.clear_button.configure(state="disabled")
            # 内容区域：显示标题区域，隐藏文件信息区域
            self.title_frame.grid(row=1, column=0, sticky="nsew")
            self.file_info_frame.grid_forget()

        logger.debug("无文件状态设置完成")

    def _switch_to_file_state(self):
        """
        切换到有文件状态

        当有文件被选择时显示文件信息界面：
        - 隐藏标题区域，显示文件信息区域
        - 启用清空按钮，允许用户清除选择
        - 根据当前模式重建文件信息显示

        这种状态提供详细的文件信息和操作选项
        """
        logger.debug("切换到有文件状态")

        # 控制栏：添加按钮始终显示在第2列，清空按钮显示在第4列并启用
        self.add_button.grid(row=0, column=2, sticky="e")
        self.clear_button.grid(row=0, column=4, sticky="e")
        self.clear_button.configure(state="normal")

        # 隐藏标题区域，显示文件信息区域
        # 使用grid_forget和grid控制显示切换，避免组件重叠
        self.title_frame.grid_forget()
        self.file_info_frame.grid(row=1, column=0, sticky="nsew")

        # 根据模式重建文件信息显示
        # 单文件模式和批量模式有不同的信息展示方式
        self._rebuild_file_info_display()

        logger.debug("有文件状态设置完成")

    def _rebuild_file_info_display(self):
        """
        根据当前模式动态重建文件信息显示区域

        该方法负责：
        - 清空旧的显示内容
        - 根据当前模式创建相应的显示组件
        - 配置布局确保信息正确显示

        这种动态重建方式确保界面始终反映最新的文件状态
        """
        logger.debug(f"重建文件信息显示，模式: {self.mode}")

        # 清空旧内容
        # 遍历所有子组件并销毁，为新的显示内容做准备
        for widget in self.file_info_frame.winfo_children():
            widget.destroy()

        # 配置grid布局
        # 确保新的内容能够正确填充可用空间
        self.file_info_frame.grid_rowconfigure(0, weight=1)
        self.file_info_frame.grid_columnconfigure(0, weight=1)

        if self.mode == "single":
            self._create_single_file_display()
        else:
            self._create_batch_info_display()

        logger.debug("文件信息显示重建完成")

    def _create_single_file_display(self):
        """
        创建单文件模式的显示内容

        单文件模式显示详细信息：
        - 文件名（主要信息）
        - 文件目录路径（辅助信息）
        - 文件格式验证警告（如有问题）

        这种详细显示帮助用户确认选择的文件是否正确
        """
        from docwen.utils.file_type_utils import validate_file_format

        row = 1  # 从第1行开始，第0行保留为弹性空间

        if self.file_path:
            # 显示文件名（主要信息）
            file_name = Path(self.file_path).name
            file_label = tb.Label(
                self.file_info_frame,
                text=file_name,
                font=(self.default_font, self.default_size),
                bootstyle="primary",  # 主要样式，突出显示
                anchor=tk.CENTER,
                justify=tk.CENTER,
            )
            file_label.grid(row=row, column=0, sticky="nsew", pady=(0, scale(5)))
            bind_single_line_ellipsis(file_label, file_name, padding=scale(12))
            self.file_info_frame.grid_rowconfigure(row, weight=0)
            row += 1

            # 显示文件目录路径（辅助信息）
            file_dir = str(Path(self.file_path).parent)
            if file_dir:
                dir_label = tb.Label(
                    self.file_info_frame,
                    text=f"📁 {file_dir}",
                    font=(self.small_font, self.small_size),
                    bootstyle="secondary",  # 次要样式，不突出
                    anchor=tk.CENTER,
                    justify=tk.CENTER,
                )
                dir_label.grid(row=row, column=0, sticky="nsew", pady=(0, scale(5)))
                bind_single_line_ellipsis(dir_label, f"📁 {file_dir}", padding=scale(12))
                self.file_info_frame.grid_rowconfigure(row, weight=0)
                row += 1

            # 显示文件格式验证警告（如有问题）
            validation_result = validate_file_format(self.file_path, t_func=t)
            if validation_result["warning_message"]:
                warning_text = validation_result["warning_message"]
                warning_label = tb.Label(
                    self.file_info_frame,
                    text=warning_text,
                    font=(self.small_font, self.small_size),
                    bootstyle="warning",  # 警告样式，提醒用户注意
                    anchor=tk.CENTER,
                    justify=tk.CENTER,
                )
                warning_label.grid(row=row, column=0, sticky="nsew")
                bind_single_line_ellipsis(warning_label, warning_text, padding=scale(12))
                self.file_info_frame.grid_rowconfigure(row, weight=0)
                row += 1

        # 配置底部弹性空间，确保内容垂直居中
        self.file_info_frame.grid_rowconfigure(row, weight=1)

    def _create_batch_info_display(self):
        """
        创建批量模式的显示内容

        批量模式显示统计信息：
        - 文件总数
        - 文件类别
        - 处理状态

        这种简洁显示适合批量操作场景，避免信息过载
        """
        # 获取选项卡式批量文件列表
        batch_list = self._get_tabbed_batch_file_list()
        if not batch_list:
            logger.warning("无法获取选项卡式批量文件列表")
            return

        # 获取当前类别的文件统计信息
        current_category = batch_list.get_current_category()
        total_files = batch_list.get_file_count(current_category)  # 传入类别参数，获取当前选项卡的文件数
        category_names = {
            "text": t("file_types.text_category"),
            "document": t("file_types.document_category"),
            "spreadsheet": t("file_types.spreadsheet_category"),
            "layout": t("file_types.layout_category"),
            "image": t("file_types.image_category"),
        }
        category_name = category_names.get(current_category, current_category)

        row = 1  # 从第1行开始，第0行保留为弹性空间

        # 显示批量处理统计信息
        total_label = tb.Label(
            self.file_info_frame,
            text=t("components.file_drop.batch_summary", count=total_files, category=category_name),
            font=(self.default_font, self.default_size),
            bootstyle="primary",
            anchor=tk.CENTER,
            justify=tk.CENTER,
        )
        total_label.grid(row=row, column=0, sticky="nsew", pady=(0, scale(10)))
        self.file_info_frame.grid_rowconfigure(row, weight=0)
        row += 1

        # 配置底部弹性空间，确保内容垂直居中
        self.file_info_frame.grid_rowconfigure(row, weight=1)

    def _on_clear_clicked(self):
        """
        处理清空按钮点击事件

        清空操作负责：
        1. 重置内部文件路径状态
        2. 调用外部清空回调（清空所有文件）
        3. 更新界面状态

        这是一个破坏性操作，需要谨慎处理
        """
        logger.debug("清空按钮被点击")
        self.file_path = ""

        # 先调用外部清空回调（清空所有文件）
        # 确保数据层和界面层同步清空
        if self.on_clear_clicked:
            self.on_clear_clicked()

        # 清空完成后，更新UI状态（此时会正确检测到所有文件已清空）
        self._switch_to_empty_state()

        logger.info("文件已彻底清空")

    def _setup_drag_drop(self):
        """
        设置拖拽支持

        该方法配置tkinterdnd2库的拖拽功能：
        - 注册拖拽目标
        - 绑定拖拽事件处理器
        - 处理库不可用的情况

        拖拽功能大大提升了用户体验，但需要额外的库支持
        """
        logger.debug("设置文件拖拽支持")

        if not TKINTERDND2_AVAILABLE:
            logger.error("tkinterdnd2未安装，拖拽功能不可用")
            return

        try:
            dnd_self = cast(_DndCapable, self)
            dnd_files = cast(str, DND_FILES)
            # 注册拖拽目标，声明该组件可以接受文件拖拽
            dnd_self.drop_target_register(dnd_files)
            # 绑定拖拽事件处理器
            dnd_self.dnd_bind("<<DragEnter>>", self._on_drag_enter)
            dnd_self.dnd_bind("<<DragLeave>>", self._on_drag_leave)
            dnd_self.dnd_bind("<<Drop>>", self._on_drop)
            logger.debug("文件拖拽支持设置成功")
        except Exception as e:
            logger.error(f"设置拖拽支持失败: {e!s}")
            self.drag_enabled = False

    def _on_drag_enter(self, event):
        """
        处理拖拽进入事件

        当文件被拖拽到组件上方时触发：
        - 设置拖拽状态标志
        - 可在此处添加视觉反馈（如改变边框颜色）
        - 返回"copy"表示接受文件复制操作

        参数:
            event: 拖拽事件对象，包含拖拽相关信息
        """
        if not self.drag_enabled:
            return

        logger.debug("文件拖拽进入区域")
        self.is_dragging = True

        return "copy"  # 表示接受文件复制操作

    def _on_drag_leave(self, event):
        """
        处理拖拽离开事件

        当文件被拖拽离开组件区域时触发：
        - 重置拖拽状态标志
        - 可在此处恢复视觉样式

        参数:
            event: 拖拽事件对象，包含拖拽相关信息
        """
        if not self.drag_enabled:
            return

        logger.debug("文件拖拽离开区域")
        self.is_dragging = False

    def _get_tabbed_batch_file_list(self):
        """
        获取选项卡式批量文件列表组件

        该方法通过组件树向上查找主窗口，然后获取批量文件列表组件。
        这种设计实现了组件间的松耦合通信。

        返回:
            TabbedBatchFileList组件实例或None（如果不可用）
        """
        try:
            current = self
            # 向上遍历组件树，查找包含_main_window属性的组件
            while current and not hasattr(current, "_main_window"):
                current = current.master

            if current and hasattr(current, "_main_window"):
                main_window = getattr(current, "_main_window", None)
                if main_window is not None and hasattr(main_window, "tabbed_batch_file_list"):
                    return main_window.tabbed_batch_file_list
        except Exception as e:
            logger.error(f"获取选项卡式批量文件列表失败: {e}")

        return None

    def _get_tabbed_file_manager(self):
        """
        获取选项卡式文件管理器组件

        该方法通过组件树向上查找主窗口，然后获取文件管理器组件。
        用于在单文件模式下管理文件选择状态。

        返回:
            TabbedFileManager组件实例或None（如果不可用）
        """
        try:
            current = self
            # 向上遍历组件树，查找包含_main_window属性的组件
            while current and not hasattr(current, "_main_window"):
                current = current.master

            if current and hasattr(current, "_main_window"):
                main_window = getattr(current, "_main_window", None)
                if main_window is not None and hasattr(main_window, "tabbed_file_manager"):
                    return main_window.tabbed_file_manager
        except Exception as e:
            logger.error(f"获取选项卡式文件管理器失败: {e}")

        return None

    def _call_tabbed_file_manager_on_mode_changed(self, new_mode: str):
        """
        调用TabbedFileManager的模式切换回调

        确保文件管理器组件能够响应模式变化，保持组件间状态同步。

        参数:
            new_mode: 新的模式（"single"或"batch"）
        """
        tabbed_file_manager = self._get_tabbed_file_manager()
        if tabbed_file_manager and hasattr(tabbed_file_manager, "on_mode_changed"):
            logger.debug(f"调用TabbedFileManager.on_mode_changed: {new_mode}")
            tabbed_file_manager.on_mode_changed(new_mode)
        else:
            logger.warning("TabbedFileManager不可用或没有on_mode_changed方法")

    def _validate_single_file_mode(self, files: list[str]) -> tuple[bool, str]:
        """
        验证单文件模式下的拖拽文件

        单文件模式有严格的验证规则：
        - 只能有一个文件
        - 不能是文件夹
        - 必须是支持的文件类型

        参数:
            files: 文件路径列表

        返回:
            (是否有效, 错误消息)
        """
        if len(files) > 1:
            return False, t("messages.single_file_only")

        if len(files) == 0:
            return False, t("messages.no_valid_files")

        file_path = files[0]

        # 检查是否是文件夹
        if Path(file_path).is_dir():
            return False, t("messages.no_folder_in_single_mode")

        # 检查文件类型是否支持
        if not self._validate_file_type(file_path):
            return False, t("messages.invalid_file_type")

        return True, ""

    def _process_batch_files(self, files: list[str]) -> list[str]:
        """
        处理批量模式下的拖拽文件，支持文件夹递归

        批量模式的处理逻辑：
        - 如果是文件夹，递归收集其中的支持文件
        - 验证单个文件类型
        - 返回支持的文件列表

        参数:
            files: 原始文件路径列表

        返回:
            处理后的支持文件列表
        """
        processed_files = []

        for file_path in files:
            if Path(file_path).is_dir():
                # 批量模式下递归处理文件夹
                logger.debug(f"批量模式递归处理文件夹: {file_path}")
                folder_files = collect_files_from_folder(file_path)
                for f in folder_files:
                    if self._validate_file_type(f):
                        processed_files.append(f)
                    else:
                        logger.debug(f"跳过不支持的文件: {f}")
            else:
                # 如果是文件，检查是否支持
                if self._validate_file_type(file_path):
                    processed_files.append(file_path)
                else:
                    logger.debug(f"跳过不支持的文件: {file_path}")

        return processed_files

    def _on_drop(self, event):
        """
        处理文件放置事件

        这是拖拽操作的核心处理方法：
        - 解析拖拽数据
        - 根据模式验证文件
        - 处理有效的文件
        - 更新界面状态
        - 调用回调函数

        参数:
            event: 放置事件对象，包含文件数据
        """
        if not self.drag_enabled:
            return

        logger.debug("文件放置事件触发")
        self.is_dragging = False

        try:
            file_path = event.data.strip()
            logger.debug(f"原始拖拽数据: {file_path}")

            # 解析拖拽数据，处理不同格式的文件路径
            files = self._parse_dropped_files(file_path)

            if self.mode == "batch":
                if len(files) > 0:
                    logger.info(f"批量模式拖拽文件: {len(files)} 个文件")

                    # 处理批量文件（支持文件夹递归）
                    processed_files = self._process_batch_files(files)
                    logger.info(f"处理后得到 {len(processed_files)} 个支持文件")

                    if processed_files:
                        # 获取选项卡式批量文件列表
                        tabbed_list = self._get_tabbed_batch_file_list()
                        if tabbed_list:
                            added_files, _failed_files = tabbed_list.add_files(processed_files)

                            if added_files:
                                self._switch_to_file_state()

                                if self.on_file_dropped:
                                    self.on_file_dropped(added_files)
                        else:
                            logger.error("错误: 选项卡式批量文件列表不可用")
                    else:
                        logger.error("错误: 未找到任何支持的文件")
                else:
                    logger.error("错误: 未检测到有效文件")

            else:
                # 单文件模式验证
                is_valid, error_msg = self._validate_single_file_mode(files)
                if not is_valid:
                    logger.error(f"单文件模式验证失败: {error_msg}")
                    return

                file_path = files[0]
                logger.info(f"单文件模式拖拽文件: {file_path}")

                self.file_path = file_path
                logger.info(f"文件拖拽成功: {file_path}")

                self._switch_to_file_state()

                if self.on_file_dropped:
                    self.on_file_dropped([file_path])

        except Exception as e:
            logger.error(f"处理文件拖拽失败: {e!s}")

    def _parse_dropped_files(self, file_data: str) -> list[str]:
        """
        解析拖拽的文件数据

        能够正确处理以下格式：
        - {file with spaces.jpg} file2.jpg
        - {file1} {file2}
        - file1 file2
        - 单个文件

        参数:
            file_data: 原始拖拽数据字符串

        返回:
            清理后的文件路径列表
        """
        import re

        files = []

        # 策略1: 使用正则表达式提取所有花括号包裹的路径
        # 匹配模式: {路径内容}
        brace_pattern = r"\{([^}]+)\}"
        brace_matches = re.findall(brace_pattern, file_data)

        if brace_matches:
            # 找到了花括号包裹的路径
            files.extend(brace_matches)

            # 移除已提取的花括号部分，处理剩余内容
            remaining = re.sub(brace_pattern, "", file_data).strip()

            if remaining:
                # 剩余部分可能包含没有花括号的路径（空格分隔）
                potential_paths = remaining.split()
                for path in potential_paths:
                    path = path.strip()
                    if path and Path(path).exists():
                        files.append(path)
        else:
            # 没有花括号，使用传统的空格分隔方式
            if " " in file_data:
                # 先尝试整个字符串是否是有效路径
                if Path(file_data).exists():
                    files = [file_data]
                else:
                    # 按空格分割并验证每个路径
                    potential_files = file_data.split(" ")
                    files = [f for f in potential_files if f.strip() and Path(f.strip()).exists()]
                    if not files:
                        # 如果都无效，保留原始数据
                        files = [file_data]
            else:
                # 单个文件路径
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

    def _validate_file_type(self, file_path: str) -> bool:
        """
        验证文件类型是否支持

        使用工具函数检查文件扩展名是否在支持列表中

        参数:
            file_path: 文件路径

        返回:
            是否支持该文件类型
        """
        try:
            validation = validate_file_format(file_path, t_func=t)
            actual_format = validation["actual_format"]
            return category_from_actual_format(actual_format) != "unknown"
        except Exception as e:
            logger.debug(f"实际格式检测失败，回退到扩展名检查: {e}")
            return is_supported_file(file_path)

    def clear(self):
        """
        清空当前选择的文件和状态

        公共方法，允许外部代码清空组件状态。
        常用于重置操作或错误恢复。
        """
        self.file_path = ""
        self._switch_to_empty_state()

    def get_file_path(self) -> str:
        """
        获取当前选择的文件路径

        返回:
            当前文件路径（单文件模式）或空字符串
        """
        return self.file_path

    def set_file_path(self, file_path: str):
        """
        手动设置文件路径

        允许通过编程方式设置文件，绕过拖拽或文件选择对话框。
        会自动验证文件类型并更新界面。

        参数:
            file_path: 要设置的文件路径
        """
        if self._validate_file_type(file_path):
            self.file_path = file_path
            self._switch_to_file_state()

            if self.on_file_dropped:
                self.on_file_dropped([file_path])
        else:
            logger.error(f"尝试设置不支持的文件类型: {file_path}")

    def set_drag_enabled(self, enabled: bool):
        """
        启用或禁用拖拽功能

        允许动态控制拖拽功能，适用于需要临时禁用拖拽的场景。

        参数:
            enabled: 是否启用拖拽功能
        """
        if not TKINTERDND2_AVAILABLE:
            return

        try:
            dnd_self = cast(_DndCapable, self)
            if enabled:
                dnd_files = cast(str, DND_FILES)
                # 启用拖拽：注册目标并绑定事件
                dnd_self.drop_target_register(dnd_files)
                dnd_self.dnd_bind("<<DragEnter>>", self._on_drag_enter)
                dnd_self.dnd_bind("<<DragLeave>>", self._on_drag_leave)
                dnd_self.dnd_bind("<<Drop>>", self._on_drop)
                self.drag_enabled = True
            else:
                # 禁用拖拽：取消注册目标
                dnd_self.drop_target_unregister()
                self.drag_enabled = False
        except Exception as e:
            logger.error(f"更改拖拽状态失败: {e!s}")
            self.drag_enabled = False

    def set_callback(self, callback: Callable):
        """
        设置文件拖拽回调函数

        允许外部代码注册文件拖拽完成时的回调函数。

        参数:
            callback: 回调函数，接受文件路径作为参数
        """
        self.on_file_dropped = callback

    def set_clear_callback(self, callback: Callable[[], None]):
        """
        设置清空按钮回调函数

        允许外部代码注册清空操作时的回调函数。

        参数:
            callback: 清空回调函数
        """
        self.on_clear_clicked = callback

    def get_mode(self) -> str:
        """
        获取当前模式

        返回:
            当前模式："single" 或 "batch"
        """
        return self.mode

    def set_mode(self, mode: str):
        """
        设置模式

        允许通过编程方式切换模式，会自动更新界面和调用相关回调。

        参数:
            mode: 要设置的模式，"single" 或 "batch"
        """
        if mode in ["single", "batch"]:
            self.mode_var.set(mode)
            self._on_mode_changed()

    def _on_add_button_clicked(self):
        """
        处理添加按钮点击事件

        当用户点击添加按钮时，打开相应的文件选择对话框。
        根据当前模式选择单个文件或多个文件。
        """
        logger.debug("添加按钮被点击")
        self._open_file_dialog()

    def _open_file_dialog(self):
        """
        打开文件选择对话框

        根据当前模式显示不同的文件选择对话框：
        - 单文件模式：选择单个文件
        - 批量模式：选择多个文件

        文件类型过滤器帮助用户快速找到支持的文件。
        """
        logger.debug("打开文件选择对话框")

        if self.mode == "single":
            # 单文件模式：选择单个文件
            file_path = filedialog.askopenfilename(
                title=t("common.select"),
                filetypes=[
                    (
                        t("messages.all_supported_files"),
                        "*.md;*.txt;*.docx;*.doc;*.wps;*.rtf;*.odt;*.xlsx;*.xls;*.et;*.csv;*.ods;*.pdf;*.xps;*.ofd;*.tif;*.tiff;*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.heic;*.heif;*.webp",
                    ),
                    (t("file_types.text"), "*.md;*.txt"),
                    (t("file_types.document"), "*.docx;*.doc;*.wps;*.rtf;*.odt"),
                    (t("file_types.spreadsheet"), "*.xlsx;*.xls;*.et;*.csv;*.ods"),
                    (t("file_types.layout"), "*.pdf;*.xps;*.ofd"),
                    (t("file_types.image"), "*.tif;*.tiff;*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.heic;*.heif;*.webp"),
                    (t("messages.all_files"), "*.*"),
                ],
            )

            if file_path:
                logger.info(f"用户选择了文件: {file_path}")
                # 模拟拖拽事件，复用现有的文件处理逻辑
                self._simulate_drop([file_path])

        else:
            # 批量模式：选择多个文件（不能选择文件夹）
            file_paths = filedialog.askopenfilenames(
                title=t("messages.select_multiple_files"),
                filetypes=[
                    (
                        t("messages.all_supported_files"),
                        "*.md;*.txt;*.docx;*.doc;*.wps;*.rtf;*.odt;*.xlsx;*.xls;*.et;*.csv;*.ods;*.pdf;*.xps;*.ofd;*.tif;*.tiff;*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.heic;*.heif;*.webp",
                    ),
                    (t("file_types.text"), "*.md;*.txt"),
                    (t("file_types.document"), "*.docx;*.doc;*.wps;*.rtf;*.odt"),
                    (t("file_types.spreadsheet"), "*.xlsx;*.xls;*.et;*.csv;*.ods"),
                    (t("file_types.layout"), "*.pdf;*.xps;*.ofd"),
                    (t("file_types.image"), "*.tif;*.tiff;*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.heic;*.heif;*.webp"),
                    (t("messages.all_files"), "*.*"),
                ],
            )

            if file_paths:
                logger.info(f"用户选择了 {len(file_paths)} 个文件")
                # 模拟拖拽事件，复用现有的文件处理逻辑
                self._simulate_drop(list(file_paths))

    def _simulate_drop(self, file_paths: list[str]):
        """
        模拟文件拖拽事件，处理选择的文件

        该方法复用拖拽事件的处理逻辑，确保文件选择对话框和
        拖拽操作有相同的行为和验证规则。

        参数:
            file_paths: 文件路径列表
        """
        logger.debug(f"模拟拖拽事件，文件列表: {file_paths}")

        if self.mode == "batch":
            if len(file_paths) > 0:
                logger.info(f"批量模式处理文件: {len(file_paths)} 个文件/文件夹")

                # 处理批量文件（支持文件夹递归）
                processed_files = self._process_batch_files(file_paths)
                logger.info(f"处理后得到 {len(processed_files)} 个支持文件")

                if processed_files:
                    # 获取选项卡式批量文件列表
                    tabbed_list = self._get_tabbed_batch_file_list()
                    if tabbed_list:
                        added_files, _failed_files = tabbed_list.add_files(processed_files)

                        if added_files:
                            self._switch_to_file_state()

                        if self.on_file_dropped:
                            self.on_file_dropped(added_files)
                    else:
                        logger.error("错误: 选项卡式批量文件列表不可用")
                else:
                    logger.error("错误: 未找到任何支持的文件")
            else:
                logger.error("错误: 未检测到有效文件")

        else:
            # 单文件模式验证
            is_valid, error_msg = self._validate_single_file_mode(file_paths)
            if not is_valid:
                logger.error(f"单文件模式验证失败: {error_msg}")
                return

            file_path = file_paths[0]
            logger.info(f"单文件模式处理文件: {file_path}")

            self.file_path = file_path
            logger.info(f"文件选择成功: {file_path}")

            self._switch_to_file_state()

            if self.on_file_dropped:
                self.on_file_dropped([file_path])

    def refresh_style(self, theme_name: str | None = None):
        """
        刷新组件样式以匹配当前主题

        当应用程序主题发生变化时调用此方法，确保组件样式
        与当前主题保持一致。

        参数:
            theme_name: 主题名称（可选）
        """
        # 更新所有主要组件的样式
        self.container.configure(bootstyle="default")
        self.title_label.configure(bootstyle="secondary")
        self.text_prompt_label.configure(bootstyle="secondary")
        self.document_prompt_label.configure(bootstyle="secondary")
        self.spreadsheet_prompt_label.configure(bootstyle="secondary")
        self.layout_prompt_label.configure(bootstyle="secondary")
        self.image_prompt_label.configure(bootstyle="secondary")
        self.text_prompt_label.configure(bootstyle="secondary")
