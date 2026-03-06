"""
版式文件预处理模块

将 OFD/XPS/CAJ 统一转换为 PDF。

支持的转换：
- OFD → PDF: 使用 easyofd 库
- XPS → PDF: 使用 pymupdf
- CAJ → PDF: 待实现

依赖：
- easyofd: OFD 转换
- pymupdf (fitz): XPS 转换

使用方式:
    from docwen.converter.formats.layout.preprocess import ofd_to_pdf
"""

import logging
import threading
from pathlib import Path

logger = logging.getLogger(__name__)


def caj_to_pdf(caj_path: str, cancel_event: threading.Event | None = None, output_dir: str | None = None) -> str:
    """
    将CAJ文件转换为PDF格式（待实现）

    CAJ (CNKI Article Journal) 是中国知网的文档格式，
    此函数将其转换为更通用的PDF格式。

    参数:
        caj_path: CAJ文件路径
        cancel_event: 取消事件（可选）
        output_dir: 输出目录（可选）。如果为None，输出到原文件同目录

    返回:
        str: 转换后的PDF文件路径

    异常:
        NotImplementedError: 功能尚未实现

    说明:
        可能的实现方案：
        1. 使用caj2pdf第三方库（https://github.com/caj2pdf/caj2pdf）
        2. 使用CAJViewer的命令行接口（如果可用）
        3. 通过脚本调用CAJViewer进行批量转换
    """
    logger.error("CAJ转PDF功能尚未实现")
    raise NotImplementedError(
        "CAJ转PDF功能尚未实现。\n"
        "可能的解决方案：\n"
        "1. 使用CAJViewer手动转换\n"
        "2. 安装caj2pdf库: pip install caj2pdf\n"
        "3. 等待未来版本支持"
    )


def xps_to_pdf(xps_path: str, cancel_event: threading.Event | None = None, output_dir: str | None = None) -> str:
    """
    将XPS文件转换为PDF格式

    XPS (XML Paper Specification) 是微软的固定布局文档格式，
    此函数使用pymupdf (fitz)库将其转换为更通用的PDF格式。

    参数:
        xps_path: XPS文件路径
        cancel_event: 取消事件（可选）
        output_dir: 输出目录（可选）。如果为None，输出到原文件同目录

    返回:
        str: 转换后的PDF文件路径

    异常:
        RuntimeError: 转换失败时抛出
    """
    try:
        import fitz  # pymupdf

        logger.info(f"开始转换XPS文件: {Path(xps_path).name}")

        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("XPS转PDF操作被取消")
            raise InterruptedError("操作已取消")

        # 使用path_utils生成输出路径（统一使用标准化命名）
        from docwen.utils.path_utils import generate_output_path

        pdf_path = generate_output_path(
            xps_path, output_dir=output_dir, section="", add_timestamp=True, description="fromXps", file_type="pdf"
        )

        logger.debug(f"输出PDF路径: {pdf_path}")

        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("XPS转PDF操作被取消")
            raise InterruptedError("操作已取消")

        # 打开XPS文件
        doc = fitz.open(xps_path)
        logger.debug(f"成功打开XPS文件，共{doc.page_count}页")

        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            doc.close()
            logger.info("XPS转PDF操作被取消")
            raise InterruptedError("操作已取消")

        # 转换为PDF字节流
        pdfbytes = doc.convert_to_pdf()
        doc.close()

        logger.debug(f"XPS已转换为PDF字节流，大小: {len(pdfbytes)} 字节")

        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("XPS转PDF操作被取消")
            raise InterruptedError("操作已取消")

        # 从字节流创建PDF文档并保存
        pdf_doc = fitz.open("pdf", pdfbytes)
        pdf_doc.save(pdf_path)
        pdf_doc.close()

        logger.info(f"成功转换为PDF: {Path(pdf_path).name}")
        return pdf_path

    except InterruptedError:
        raise
    except ImportError as e:
        logger.error(f"缺少必要的库: {e}")
        raise RuntimeError("XPS转PDF需要pymupdf库。\n请安装: pip install pymupdf") from e
    except Exception as e:
        logger.error(f"XPS转PDF失败: {e}", exc_info=True)
        raise RuntimeError(f"XPS转PDF失败: {e}") from e


def ofd_to_pdf(ofd_path: str, cancel_event: threading.Event | None = None, output_dir: str | None = None) -> str:
    """
    将OFD文件转换为PDF格式

    OFD (Open Fixed-layout Document) 是中国国家电子文件标准格式，
    此函数使用easyofd库将其转换为更通用的PDF格式。

    参数:
        ofd_path: OFD文件路径
        cancel_event: 取消事件（可选）
        output_dir: 输出目录（可选）。如果为None，输出到原文件同目录

    返回:
        str: 转换后的PDF文件路径

    异常:
        RuntimeError: 转换失败时抛出
    """
    try:
        from easyofd import OFD
        from docwen.compat import apply_easyofd_patches

        apply_easyofd_patches()

        logger.info(f"开始转换OFD文件: {Path(ofd_path).name}")

        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("OFD转PDF操作被取消")
            raise InterruptedError("操作已取消")

        # 使用path_utils生成输出路径（统一使用标准化命名）
        from docwen.utils.path_utils import generate_output_path

        pdf_path = generate_output_path(
            ofd_path, output_dir=output_dir, section="", add_timestamp=True, description="fromOfd", file_type="pdf"
        )

        logger.debug(f"输出PDF路径: {pdf_path}")

        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("OFD转PDF操作被取消")
            raise InterruptedError("操作已取消")

        # 读取OFD文件
        ofd = OFD()

        try:
            ofd.read(str(ofd_path), fmt="path")
            logger.debug("成功读取OFD文件")
        except AttributeError as e:
            logger.warning(f"读取OFD时遇到AttributeError（可能是easyofd库的已知问题）: {e}")
            # 继续尝试转换，因为某些AttributeError不影响最终结果

        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("OFD转PDF操作被取消")
            raise InterruptedError("操作已取消")

        # 转换为PDF字节流
        pdf_bytes = None
        try:
            pdf_bytes = ofd.to_pdf()
            logger.debug(f"OFD已转换为PDF字节流，大小: {len(pdf_bytes)} 字节")
        except AttributeError as e:
            # easyofd库存在已知问题
            logger.warning(
                f"OFD转PDF过程中遇到AttributeError（easyofd库的已知问题）: {e}\n"
                f"这通常发生在处理某些OFD注释对象时，但可能不影响最终转换结果。"
            )
            if not pdf_bytes:
                logger.error("由于easyofd内部错误，无法生成PDF字节流")
                raise RuntimeError(
                    "OFD转PDF失败：easyofd库内部错误（AttributeError）。\n"
                    "建议：\n"
                    "1. 升级easyofd库到最新版本：pip install --upgrade easyofd\n"
                    "2. 或尝试使用其他OFD查看器手动转换"
                ) from e

        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("OFD转PDF操作被取消")
            raise InterruptedError("操作已取消")

        # 写入PDF文件
        with Path(pdf_path).open("wb") as f:
            f.write(pdf_bytes)

        logger.info(f"成功转换为PDF: {Path(pdf_path).name}")
        return pdf_path

    except InterruptedError:
        raise
    except ImportError as e:
        logger.error(f"缺少必要的库: {e}")
        raise RuntimeError("OFD转PDF需要easyofd库。\n请安装: pip install easyofd") from e
    except Exception as e:
        logger.error(f"OFD转PDF失败: {e}", exc_info=True)
        raise RuntimeError(f"OFD转PDF失败: {e}") from e
