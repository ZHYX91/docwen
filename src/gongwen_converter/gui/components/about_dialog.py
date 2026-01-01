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
import os
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from .base_dialog import BaseDialog

# 配置日志记录器
logger = logging.getLogger(__name__)


class AboutDialog(BaseDialog):
    """
    关于对话框类
    显示应用程序的版本信息、联系方式、免责声明和版权信息
    与主界面和设置界面保持一致的图标和样式
    """
    
    def __init__(self, parent: tb.Window, title: str = "关于"):
        """
        初始化关于对话框
        
        参数:
            parent: 父窗口对象
            title: 对话框标题，默认为"关于"
        """
        super().__init__(parent, title=title, modal=True)
        logger.debug("初始化关于对话框")
        
        # 设置对话框属性 - 增加高度以容纳致谢信息（两列布局）
        width = self.scale(400)
        height = self.scale(600)
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
        from gongwen_converter.utils.font_utils import get_default_font, get_small_font, get_title_font
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
        
        # 配置对话框网格权重 - 增加致谢部分
        self.grid_rowconfigure(0, weight=0)  # 标题
        self.grid_rowconfigure(1, weight=0)  # 副标题
        self.grid_rowconfigure(2, weight=0)  # 版本
        self.grid_rowconfigure(3, weight=0)  # 联系方式
        self.grid_rowconfigure(4, weight=0)  # 版权
        self.grid_rowconfigure(5, weight=1)  # 致谢
        self.grid_rowconfigure(6, weight=0)  # 关闭按钮
        self.grid_columnconfigure(0, weight=1)
        
        # 添加简称标题
        title_label = tb.Label(
            self,
            text="公文转换器",
            font=(self.title_font, self.title_size, "bold"),
            bootstyle="primary"
        )
        title_label.grid(row=0, column=0, pady=(self.scale(20), self.scale(10)), sticky="n")
        
        # 导入信息图标创建函数
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 添加版本信息（带信息图标）
        from gongwen_converter import __version__
        version_frame = tb.Frame(self)
        version_frame.grid(row=2, column=0, pady=(0, self.scale(10)), sticky="n")
        
        version_label = tb.Label(
            version_frame,
            text=f"版本 {__version__}",
            font=(self.default_font, self.default_size),
            bootstyle="secondary"
        )
        version_label.pack(side="left")
        
        version_info_icon = create_info_icon(
            version_frame,
            "本软件禁止联网，需要用户自行检查更新",
            "info"
        )
        version_info_icon.pack(side="left", padx=(self.scale(5), 0))
        
        # 添加联系方式
        contact_label = tb.Label(
            self,
            text="联系邮箱: zhengyx91@hotmail.com",
            font=(self.default_font, self.default_size),
            bootstyle="default",
            justify="center"
        )
        contact_label.grid(row=3, column=0, pady=(self.scale(10), self.scale(10)), sticky="n")
        
        # 添加版权信息
        copyright_label = tb.Label(
            self,
            text="© 2025 ZhengYX",
            font=(self.default_font, self.default_size),
            bootstyle="secondary",
            justify="center"
        )
        copyright_label.grid(row=4, column=0, pady=(self.scale(10), self.scale(5)), sticky="n")
        
        # 添加致谢部分（两列布局）
        acknowledgments_frame = tb.Labelframe(
            self,
            text="致谢",
            bootstyle="info",
            padding=self.scale(15)
        )
        acknowledgments_frame.grid(row=5, column=0, padx=self.scale(20), pady=self.scale(10), sticky="nsew")
        
        # 说明文字
        intro_label = tb.Label(
            acknowledgments_frame,
            text="本项目使用了以下优秀的开源项目和工具，特此致谢：",
            font=(self.small_font, self.small_size),
            bootstyle="secondary",
            anchor="w"
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
        from gongwen_converter.utils.gui_utils import create_info_icon
        
        # 开源工具列表（左列）
        left_tools = [
            ("Python-docx", "用于实现 docx 与 markdown 的双向转换，处理文档样式和结构"),
            ("openpyxl", "用于实现 xlsx 与 markdown 的双向转换，处理表格数据"),
            ("PyMuPDF (fitz)", "用于 PDF/XPS 文档的内容提取和格式转换"),
            ("pymupdf4llm", "用于智能提取 PDF 文档并转换为 Markdown 格式"),
            ("pdf2docx", "用于将 PDF 文档转换为可编辑的 Word 格式"),
            ("easyofd", "用于处理国产 OFD 格式文档的读取和转换"),
            ("PaddleOCR", "用于图片和 PDF 中的文字识别（OCR），支持离线使用"),
            ("PaddlePaddle", "作为 PaddleOCR 的深度学习推理引擎"),
            ("ttkbootstrap", "用于构建现代化的图形用户界面"),
            ("tkinterdnd2", "用于实现文件拖放功能，提升用户体验"),
            ("Pillow (PIL)", "用于图片格式转换和图标资源处理"),
            ("pillow-heif", "为 Pillow 提供 HEIF 图片格式支持"),
        ]
        
        # 开源工具列表（右列）
        right_tools = [
            ("img2pdf", "用于批量将图片文件转换为 PDF 文档"),
            ("pywin32", "用于通过 COM 接口调用 WPS 和 Office 进行文档格式转换"),
            ("lxml", "用于处理 Office 文档的 XML 结构"),
            ("latex2mathml", "用于将 LaTeX 公式转换为 MathML 格式"),
            ("PyYAML", "用于处理 Markdown frontmatter 元数据"),
            ("tomlkit", "用于读取和保存配置文件"),
            ("pandas", "用于表格数据的处理和分析"),
            ("numpy", "用于数值计算和数据处理的基础支持"),
            ("olefile", "用于识别和处理旧版 Office 文档格式"),
            ("watchdog", "用于监控文件系统变化，实现进程间通信和单实例运行"),
        ]
        
        # 创建左列工具列表（使用第1列放工具名称，第2列放信息图标）
        for row, (name, tooltip) in enumerate(left_tools):
            # 工具名称放在第1列
            tool_label = tb.Label(
                content_frame,
                text=name,
                font=(self.small_font, self.small_size),
                bootstyle="default"
            )
            tool_label.grid(row=row, column=1, sticky="w", pady=self.scale(3))
            
            # 信息图标放在第2列
            info_icon = create_info_icon(content_frame, tooltip, "info")
            info_icon.grid(row=row, column=2, sticky="w", padx=(self.scale(5), 0), pady=self.scale(3))
        
        # 创建右列工具列表（使用第4列放工具名称，第5列放信息图标）
        for row, (name, tooltip) in enumerate(right_tools):
            # 工具名称放在第4列
            tool_label = tb.Label(
                content_frame,
                text=name,
                font=(self.small_font, self.small_size),
                bootstyle="default"
            )
            tool_label.grid(row=row, column=4, sticky="w", pady=self.scale(3))
            
            # 信息图标放在第5列
            info_icon = create_info_icon(content_frame, tooltip, "info")
            info_icon.grid(row=row, column=5, sticky="w", padx=(self.scale(5), 0), pady=self.scale(3))
        
        # 添加关闭按钮
        close_button = tb.Button(
            self,
            text="关闭",
            command=self.destroy,
            bootstyle="secondary",
            width=self.scale(10)
        )
        close_button.grid(row=6, column=0, pady=(self.scale(10), self.scale(20)), sticky="s")
        
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


# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 创建测试窗口
    root = tb.Window(title="关于对话框测试", themename="morph")
    root.geometry("300x200")
    
    # 测试按钮回调函数
    def show_about():
        about_dialog = AboutDialog(root)
        about_dialog.show()
    
    # 添加测试按钮
    test_button = tb.Button(
        root,
        text="显示关于对话框",
        command=show_about,
        bootstyle="primary"
    )
    test_button.pack(expand=True)
    
    logger.info("关于对话框测试启动")
    root.mainloop()
