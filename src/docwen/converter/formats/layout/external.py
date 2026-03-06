"""
版式格式转换 - 外部软件实现

提供PDF到DOCX的转换功能，支持多种转换工具的三级容错机制：
1. Microsoft Word COM接口（质量最佳）
2. LibreOffice命令行（免费开源）
3. pdf2docx纯Python库（备选方案）

依赖：
- common.detection: 软件可用性检测（如有需要）
- common.fallback: 通用容错框架（如有需要）

使用方式：
    from docwen.converter.formats.layout import pdf_to_docx

    docx_path = pdf_to_docx("input.pdf", "output.docx")
"""

import contextlib
import logging
import shutil
import subprocess
import tempfile
import threading
from collections.abc import Callable
from pathlib import Path

logger = logging.getLogger(__name__)


# ==================== LibreOffice支持 ====================


def _check_libreoffice_available() -> bool:
    """
    检查系统中是否安装了LibreOffice

    返回:
        bool: True表示可用，False表示不可用
    """
    try:
        result = subprocess.run(["soffice", "--version"], capture_output=True, timeout=5, text=True)
        if result.returncode == 0:
            logger.debug(f"检测到LibreOffice: {result.stdout.strip()}")
            return True
        return False
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False
    except Exception as e:
        logger.debug(f"检测LibreOffice时出错: {e}")
        return False


def _convert_pdf_with_libreoffice(
    pdf_path: str, output_dir: str, cancel_event: threading.Event | None = None
) -> str | None:
    """
    使用LibreOffice命令行将PDF转换为DOCX

    参数:
        pdf_path: PDF文件路径
        output_dir: 输出目录
        cancel_event: 取消事件（可选）

    返回:
        str: 转换后的DOCX文件路径，失败时返回None
    """
    try:
        if cancel_event and cancel_event.is_set():
            logger.info("LibreOffice转换被取消")
            return None

        cmd = ["soffice", "--headless", "--convert-to", "docx", "--outdir", output_dir, pdf_path]

        logger.info(f"使用LibreOffice转换PDF: {Path(pdf_path).name}")
        logger.debug(f"LibreOffice命令: {' '.join(cmd)}")

        # PDF转换可能需要更长的时间
        result = subprocess.run(
            cmd,
            capture_output=True,
            timeout=180,  # 3分钟超时
            text=True,
        )

        if result.returncode == 0:
            base_name = Path(pdf_path).stem
            output_file = str(Path(output_dir) / f"{base_name}.docx")

            if Path(output_file).exists():
                logger.info(f"✓ LibreOffice转换成功: {Path(output_file).name}")
                return output_file
            else:
                logger.error("LibreOffice执行成功但未找到输出文件")
                if result.stdout:
                    logger.debug(f"stdout: {result.stdout}")
                if result.stderr:
                    logger.debug(f"stderr: {result.stderr}")
                return None
        else:
            logger.error(f"LibreOffice转换失败（返回码: {result.returncode}）")
            if result.stderr:
                logger.error(f"错误信息: {result.stderr}")
            return None

    except FileNotFoundError:
        logger.debug("LibreOffice未安装")
        return None
    except subprocess.TimeoutExpired:
        logger.error("LibreOffice转换超时（超过3分钟）")
        return None
    except Exception as e:
        logger.error(f"LibreOffice转换出错: {e}")
        return None


# ==================== Word COM接口支持 ====================


def _convert_pdf_with_word(pdf_path: str, output_path: str, cancel_event: threading.Event | None = None) -> str | None:
    """
    使用Microsoft Word COM接口将PDF转换为DOCX

    参数:
        pdf_path: PDF文件路径
        output_path: 输出DOCX文件路径
        cancel_event: 取消事件（可选）

    返回:
        str: 转换后的DOCX文件路径，失败时返回None
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        logger.debug("pywin32未安装，无法使用Word COM接口")
        return None

    if cancel_event and cancel_event.is_set():
        logger.info("Word转换被取消")
        return None

    word_app = None
    doc = None

    try:
        pythoncom.CoInitialize()

        logger.info(f"使用Microsoft Word转换PDF: {Path(pdf_path).name}")

        # 创建Word应用程序实例
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False

        # 打开PDF文件
        abs_pdf_path = str(Path(pdf_path).resolve())
        abs_output_path = str(Path(output_path).resolve())

        logger.debug(f"打开PDF: {abs_pdf_path}")
        doc = word_app.Documents.Open(abs_pdf_path, ReadOnly=True, ConfirmConversions=False, AddToRecentFiles=False)

        if cancel_event and cancel_event.is_set():
            logger.info("Word转换被取消")
            return None

        # 保存为DOCX格式
        # FileFormat: 12 = wdFormatXMLDocument (DOCX)
        logger.debug(f"保存为DOCX: {abs_output_path}")
        doc.SaveAs(abs_output_path, FileFormat=12)

        logger.info(f"✓ Word转换成功: {Path(output_path).name}")
        return output_path

    except Exception as e:
        logger.debug(f"Word COM接口转换失败: {e}")
        return None
    finally:
        # 清理资源
        if doc:
            with contextlib.suppress(Exception):
                doc.Close(SaveChanges=False)
        if word_app:
            with contextlib.suppress(Exception):
                word_app.Quit()
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()


# ==================== pdf2docx库支持 ====================


def _convert_pdf_with_pdf2docx(
    pdf_path: str, output_path: str, cancel_event: threading.Event | None = None
) -> str | None:
    """
    使用pdf2docx纯Python库将PDF转换为DOCX

    这是最后的备选方案，不需要外部软件。
    优点: 纯Python实现，跨平台
    缺点: 转换质量可能不如Word/LibreOffice

    参数:
        pdf_path: PDF文件路径
        output_path: 输出DOCX文件路径
        cancel_event: 取消事件（可选）

    返回:
        str: 转换后的DOCX文件路径，失败时返回None
    """
    try:
        from pdf2docx import Converter
    except ImportError:
        logger.debug("pdf2docx库未安装")
        return None

    if cancel_event and cancel_event.is_set():
        logger.info("pdf2docx转换被取消")
        return None

    try:
        logger.info(f"使用pdf2docx库转换PDF: {Path(pdf_path).name}")

        cv = Converter(pdf_path)
        cv.convert(output_path)
        cv.close()

        if Path(output_path).exists():
            logger.info(f"✓ pdf2docx转换成功: {Path(output_path).name}")
            return output_path
        else:
            logger.error("pdf2docx执行完成但未找到输出文件")
            return None

    except Exception as e:
        logger.error(f"pdf2docx转换失败: {e}")
        return None


# ==================== 统一的PDF转DOCX接口 ====================


def pdf_to_docx(
    pdf_path: str,
    output_path: str,
    cancel_event: threading.Event | None = None,
    software_callback: Callable[[str], None] | None = None,
    headless: bool = False,
) -> str | None:
    """
    将PDF转换为DOCX（三级容错 + 临时目录保护）

    使用临时目录处理PDF副本，保护原始PDF文件不被修改或锁定。

    转换优先级:
    1. Microsoft Word COM接口（质量最佳，但需要安装Word）
    2. LibreOffice命令行（免费开源，跨平台）
    3. pdf2docx纯Python库（备选方案，不需要外部软件）

    参数:
        pdf_path: PDF文件路径
        output_path: 输出DOCX文件路径
        cancel_event: 取消事件（可选）
        software_callback: 软件使用回调函数（可选），用于记录使用了哪个工具
        headless: True表示禁止任何UI交互（如弹窗），适用于CLI/服务端环境

    返回:
        str: 成功时返回output_path
        None: 失败时返回None

    异常:
        InterruptedError: 操作被用户取消

    使用示例:
        >>> docx_path = pdf_to_docx("test.pdf", "test.docx")
        >>> if docx_path:
        ...     print(f"转换成功: {docx_path}")
    """
    if not Path(pdf_path).exists():
        logger.error(f"PDF文件不存在: {pdf_path}")
        return None

    if cancel_event and cancel_event.is_set():
        logger.info("转换在开始前被取消")
        raise InterruptedError("操作已被取消")

    logger.info(f"开始PDF转DOCX: {Path(pdf_path).name}")
    logger.debug(f"输入: {pdf_path}")
    logger.debug(f"输出: {output_path}")

    # 使用临时目录处理转换
    with tempfile.TemporaryDirectory() as temp_dir:
        # 步骤1: 复制PDF到临时目录（保护原始文件）
        temp_pdf = str(Path(temp_dir) / "input.pdf")
        shutil.copy2(pdf_path, temp_pdf)
        logger.debug(f"已创建PDF副本: {Path(temp_pdf).name}")

        # 步骤2: 在临时目录生成临时DOCX
        temp_docx = str(Path(temp_dir) / "temp_output.docx")

        # 步骤3: 三级容错转换（操作临时副本）
        # 第1级: 尝试Microsoft Word
        logger.info("[1/3] 尝试Microsoft Word")
        result = _convert_pdf_with_word(temp_pdf, temp_docx, cancel_event)
        if result and Path(result).exists():
            logger.info("✓ 转换成功 [Microsoft Word]")
            if software_callback:
                software_callback("Microsoft Word")
            from docwen.utils.workspace_manager import finalize_output

            saved_path, _ = finalize_output(result, output_path, original_input_file=pdf_path)
            return saved_path

        if cancel_event and cancel_event.is_set():
            logger.info("转换被用户取消")
            raise InterruptedError("操作已被取消")

        # 第2级: 尝试LibreOffice
        logger.info("[2/3] 尝试LibreOffice")
        result = _convert_pdf_with_libreoffice(temp_pdf, temp_dir, cancel_event)
        if result and Path(result).exists():
            logger.info("✓ 转换成功 [LibreOffice]")
            if software_callback:
                software_callback("LibreOffice")
            from docwen.utils.workspace_manager import finalize_output

            saved_path, _ = finalize_output(result, output_path, original_input_file=pdf_path)
            return saved_path

        if cancel_event and cancel_event.is_set():
            logger.info("转换被用户取消")
            raise InterruptedError("操作已被取消")

        # 第3级: 尝试pdf2docx库
        logger.info("[3/3] 尝试pdf2docx库")
        result = _convert_pdf_with_pdf2docx(temp_pdf, temp_docx, cancel_event)
        if result and Path(result).exists():
            logger.info("✓ 转换成功 [pdf2docx库]")
            if software_callback:
                software_callback("pdf2docx")
            from docwen.utils.workspace_manager import finalize_output

            saved_path, _ = finalize_output(result, output_path, original_input_file=pdf_path)
            return saved_path

        # 所有方法都失败
        logger.error("✗ PDF转DOCX失败: 所有转换方法均不可用")
        if not headless:
            _show_conversion_failed_dialog()
        return None
    # 临时目录自动清理


def _show_conversion_failed_dialog(parent=None):
    """
    显示友好的转换失败提示对话框

    参数:
        parent: 父窗口对象（可选，用于定位对话框）
    """
    try:
        from tkinter import messagebox

        if parent is None:
            # 如果没有提供父窗口，创建临时根窗口
            import tkinter as tk

            root = tk.Tk()
            root.withdraw()
            parent_window = root
        else:
            parent_window = parent
            root = None

        messagebox.showerror(
            "PDF转换失败",
            "无法将PDF转换为DOCX格式，可能的原因:\n\n"
            "❌ 未安装支持的转换工具\n"
            "   • Microsoft Word\n"
            "   • LibreOffice\n"
            "   • pdf2docx库\n\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "建议解决方案:\n\n"
            "1. 安装Microsoft Word（推荐，转换质量最佳）\n"
            "2. 安装LibreOffice（免费开源）\n"
            "3. 安装pdf2docx库:\n"
            "   pip install pdf2docx\n\n"
            "4. 或使用只提取文本(a)、只提取图片(b)、\n"
            "   只OCR(c)、图片+OCR(bc)等\n"
            "   不需要外部工具的组合",
            parent=parent_window,
        )

        if root:
            root.destroy()
    except Exception as e:
        logger.error(f"显示错误对话框失败: {e}")


# ==================== 公开API ====================

__all__ = [
    "pdf_to_docx",
]
