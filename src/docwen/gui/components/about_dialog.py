"""
关于对话框组件 - 显示应用程序信息

提供关于对话框功能，显示：
- 应用程序名称和版本信息
- 作者和联系方式
- 版权声明
- 与主界面保持一致的图标和样式

继承自BaseDialog基类，自动获取父窗口图标。
"""

import logging

import ttkbootstrap as tb
from ttkbootstrap.constants import *

from docwen.i18n import t

from .base_dialog import BaseDialog

# 配置日志记录器
logger = logging.getLogger(__name__)


class AboutDialog(BaseDialog):
    """
    关于对话框类
    显示应用程序的版本信息、联系方式、免责声明和版权信息
    与主界面和设置界面保持一致的图标和样式
    """

    def __init__(self, parent: tb.Window, title: str | None = None):
        """
        初始化关于对话框

        参数:
            parent: 父窗口对象
            title: 对话框标题，默认为None（使用国际化标题）
        """
        if title is None:
            title = t("about.title")
        super().__init__(parent, title=title, modal=True)
        logger.debug("初始化关于对话框")

        # 设置对话框属性
        width = self.scale(400)
        height = self.scale(680)
        self.geometry(f"{width}x{height}")
        self.resizable(False, False)

        # 初始化字体配置 - 优化重复调用
        self._initialize_fonts()

        # 创建界面
        self._create_interface()

        # 居中对话框
        self.center_on_parent()

        logger.info("关于对话框初始化完成")

    def _initialize_fonts(self):
        """
        初始化字体配置

        在初始化阶段一次性获取所有需要的字体配置，避免在界面创建过程中重复调用。
        包括默认字体、小字体和标题字体。
        """
        logger.debug("初始化字体配置")
        from docwen.utils.font_utils import get_default_font, get_small_font, get_title_font

        self.default_font, self.default_size = get_default_font()
        self.small_font, self.small_size = get_small_font()
        self.title_font, self.title_size = get_title_font()
        logger.debug("字体配置初始化完成")

    def _create_interface(self):
        """
        创建对话框界面

        构建完整的关于对话框界面，包括：
        - 应用程序标题和副标题
        - 版本信息
        - 联系方式和作者信息
        - 版权声明
        - 关闭按钮

        所有元素使用网格布局排列，支持DPI缩放。
        """
        logger.debug("创建关于对话框界面")

        # 配置对话框网格权重 - 增加致谢部分和免责声明
        self.grid_rowconfigure(0, weight=0)  # 标题
        self.grid_rowconfigure(1, weight=0)  # 副标题
        self.grid_rowconfigure(2, weight=0)  # 版本
        self.grid_rowconfigure(3, weight=0)  # 联系方式
        self.grid_rowconfigure(4, weight=0)  # 版权
        self.grid_rowconfigure(5, weight=0)  # 免责声明
        self.grid_rowconfigure(6, weight=1)  # 致谢
        self.grid_rowconfigure(7, weight=0)  # 关闭按钮
        self.grid_columnconfigure(0, weight=1)

        # 添加简称标题
        title_label = tb.Label(
            self, text=t("common.app_name"), font=(self.title_font, self.title_size, "bold"), bootstyle="primary"
        )
        title_label.grid(row=0, column=0, pady=(self.scale(20), self.scale(10)), sticky="n")

        # 导入信息图标创建函数
        # 添加版本信息（带信息图标）
        from docwen import __version__
        from docwen.utils.gui_utils import create_info_icon

        version_frame = tb.Frame(self)
        version_frame.grid(row=2, column=0, pady=(0, self.scale(10)), sticky="n")

        version_label = tb.Label(
            version_frame,
            text=t("about.version_info", version=__version__),
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
        )
        version_label.pack(side="left")

        version_info_icon = create_info_icon(version_frame, t("about.version_tooltip"), "info")
        version_info_icon.pack(side="left", padx=(self.scale(5), 0))

        # 添加联系方式
        contact_label = tb.Label(
            self,
            text=t("about.contact", email="zhengyx91@hotmail.com"),
            font=(self.default_font, self.default_size),
            bootstyle="default",
            justify="center",
        )
        contact_label.grid(row=3, column=0, pady=(self.scale(10), self.scale(10)), sticky="n")

        # 添加版权信息
        copyright_label = tb.Label(
            self,
            text=t("about.copyright", year="2025", author="ZhengYX"),
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            justify="center",
        )
        copyright_label.grid(row=4, column=0, pady=(self.scale(10), self.scale(5)), sticky="n")

        # 添加免责声明
        disclaimer_label = tb.Label(
            self,
            text=t("common.disclaimer"),
            font=(self.small_font, self.small_size),
            bootstyle="warning",
            justify="center",
        )
        disclaimer_label.grid(row=5, column=0, pady=(self.scale(5), self.scale(10)), sticky="n")

        # 添加致谢部分（两列布局）
        acknowledgments_frame = tb.Labelframe(
            self, text=t("about.acknowledgments"), bootstyle="info", padding=self.scale(15)
        )
        acknowledgments_frame.grid(row=6, column=0, padx=self.scale(20), pady=self.scale(10), sticky="nsew")

        # 说明文字
        intro_label = tb.Label(
            acknowledgments_frame,
            text=t("about.acknowledgments_intro"),
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            anchor="w",
            wraplength=self.scale(330),
        )
        intro_label.pack(fill="x", pady=(0, self.scale(10)))

        # 创建七列内容框架（弹性空白 + 工具 + 图标 + 弹性空白 + 工具 + 图标 + 弹性空白）
        content_frame = tb.Frame(acknowledgments_frame)
        content_frame.pack(fill="both", expand=True)

        # 配置七列权重
        content_frame.grid_columnconfigure(0, weight=1)  # 左侧弹性空白列
        content_frame.grid_columnconfigure(1, weight=0)  # 左侧工具名称列
        content_frame.grid_columnconfigure(2, weight=0)  # 左侧信息图标列
        content_frame.grid_columnconfigure(3, weight=2)  # 中间弹性空白列
        content_frame.grid_columnconfigure(4, weight=0)  # 右侧工具名称列
        content_frame.grid_columnconfigure(5, weight=0)  # 右侧信息图标列
        content_frame.grid_columnconfigure(6, weight=2)  # 右侧弹性空白列

        # 导入信息图标创建函数
        from docwen.utils.gui_utils import create_info_icon

        # 开源工具列表（左列）
        left_tools = [
            ("Python-docx", t("about.tools.python_docx")),
            ("openpyxl", t("about.tools.openpyxl")),
            ("PyMuPDF (fitz)", t("about.tools.pymupdf")),
            ("pymupdf4llm", t("about.tools.pymupdf4llm")),
            ("pdf2docx", t("about.tools.pdf2docx")),
            ("easyofd", t("about.tools.easyofd")),
            ("RapidOCR", t("about.tools.rapidocr")),
            ("PaddleOCR", t("about.tools.paddleocr")),
            ("ONNX Runtime", t("about.tools.onnxruntime")),
            ("ttkbootstrap", t("about.tools.ttkbootstrap")),
            ("tkinterdnd2", t("about.tools.tkinterdnd2")),
            ("Pillow (PIL)", t("about.tools.pillow")),
        ]

        # 开源工具列表（右列）
        right_tools = [
            ("pillow-heif", t("about.tools.pillow_heif")),
            ("img2pdf", t("about.tools.img2pdf")),
            ("pywin32", t("about.tools.pywin32")),
            ("lxml", t("about.tools.lxml")),
            ("latex2mathml", t("about.tools.latex2mathml")),
            ("PyYAML", t("about.tools.pyyaml")),
            ("tomlkit", t("about.tools.tomlkit")),
            ("pandas", t("about.tools.pandas")),
            ("numpy", t("about.tools.numpy")),
            ("olefile", t("about.tools.olefile")),
            ("watchdog", t("about.tools.watchdog")),
            ("emoji", t("about.tools.emoji")),
        ]

        # 创建左列工具列表（使用第1列放工具名称，第2列放信息图标）
        for row, (name, tooltip) in enumerate(left_tools):
            # 工具名称放在第1列
            tool_label = tb.Label(
                content_frame, text=name, font=(self.small_font, self.small_size), bootstyle="default"
            )
            tool_label.grid(row=row, column=1, sticky="w", pady=self.scale(3))

            # 信息图标放在第2列
            info_icon = create_info_icon(content_frame, tooltip, "info")
            info_icon.grid(row=row, column=2, sticky="w", padx=(self.scale(5), 0), pady=self.scale(3))

        # 创建右列工具列表（使用第4列放工具名称，第5列放信息图标）
        for row, (name, tooltip) in enumerate(right_tools):
            # 工具名称放在第4列
            tool_label = tb.Label(
                content_frame, text=name, font=(self.small_font, self.small_size), bootstyle="default"
            )
            tool_label.grid(row=row, column=4, sticky="w", pady=self.scale(3))

            # 信息图标放在第5列
            info_icon = create_info_icon(content_frame, tooltip, "info")
            info_icon.grid(row=row, column=5, sticky="w", padx=(self.scale(5), 0), pady=self.scale(3))

        # 添加关闭按钮
        close_button = tb.Button(
            self, text=t("common.close"), command=self.destroy, bootstyle="secondary", width=self.scale(10)
        )
        close_button.grid(row=7, column=0, pady=(self.scale(10), self.scale(20)), sticky="s")

        logger.debug("关于对话框界面创建完成")

    def show(self):
        """
        显示对话框

        使对话框可见并等待用户关闭。
        这是模态对话框，会阻塞调用者直到对话框关闭。
        """
        logger.debug("显示关于对话框")
        self.deiconify()
        self.wait_window()  # 等待对话框关闭
