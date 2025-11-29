"""
版式文件格式支持模块

支持CAJ、OFD等版式文件格式转换为PDF
版式文件的特点：
- 固定布局，不会因设备或软件变化而改变
- 通常不可编辑或编辑困难
- 适合作为最终版本保存和传播
"""

import os
import logging
from typing import Optional
import threading

# 配置日志
logger = logging.getLogger(__name__)


def caj_to_pdf(caj_path: str, cancel_event: Optional[threading.Event] = None, output_dir: Optional[str] = None) -> str:
    """
    将CAJ文件转换为PDF格式（待实现）
    
    CAJ (CNKI Article Journal) 是中国知网的文档格式，
    此函数将其转换为更通用的PDF格式。
    
    参数:
        caj_path: CAJ文件路径
        cancel_event: 取消事件（可选）
        output_dir: 输出目录（可选）。如果为None，输出到原文件同目录；如果指定，输出到该目录（通常是临时目录）
        
    返回:
        str: 转换后的PDF文件路径
        
    异常:
        NotImplementedError: 功能尚未实现
        
    说明:
        可能的实现方案：
        1. 使用caj2pdf第三方库（https://github.com/caj2pdf/caj2pdf）
        2. 使用CAJViewer的命令行接口（如果可用）
        3. 通过脚本调用CAJViewer进行批量转换
        
    注意:
        CAJ格式是专有格式，转换可能需要第三方工具或库
    """
    logger.error("CAJ转PDF功能尚未实现")
    raise NotImplementedError(
        "CAJ转PDF功能尚未实现。\n"
        "可能的解决方案：\n"
        "1. 使用CAJViewer手动转换\n"
        "2. 安装caj2pdf库: pip install caj2pdf\n"
        "3. 等待未来版本支持"
    )


def xps_to_pdf(xps_path: str, cancel_event: Optional[threading.Event] = None, output_dir: Optional[str] = None) -> str:
    """
    将XPS文件转换为PDF格式
    
    XPS (XML Paper Specification) 是微软的固定布局文档格式，
    此函数使用pymupdf (fitz)库将其转换为更通用的PDF格式。
    
    参数:
        xps_path: XPS文件路径
        cancel_event: 取消事件（可选）
        output_dir: 输出目录（可选）。如果为None，输出到原文件同目录；如果指定，输出到该目录（通常是临时目录）
        
    返回:
        str: 转换后的PDF文件路径
        
    异常:
        RuntimeError: 转换失败时抛出
        
    说明:
        使用pymupdf (fitz)库进行转换，pymupdf原生支持XPS格式。
        转换流程：
        1. 使用fitz.open()打开XPS文件
        2. 使用convert_to_pdf()方法转换为PDF字节流
        3. 保存为PDF文件
        
        注意：统一使用标准化命名规则（带时间戳），无论是否在临时目录中
    """
    try:
        import fitz  # pymupdf
        
        logger.info(f"开始转换XPS文件: {os.path.basename(xps_path)}")
        
        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("XPS转PDF操作被取消")
            raise InterruptedError("操作已取消")
        
        # 使用path_utils生成输出路径（统一使用标准化命名）
        from gongwen_converter.utils.path_utils import generate_output_path
        
        pdf_path = generate_output_path(
            xps_path,
            output_dir=output_dir,
            section="",
            add_timestamp=True,
            description="fromXps",
            file_type="pdf"
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
        
        logger.info(f"成功转换为PDF: {os.path.basename(pdf_path)}")
        return pdf_path
        
    except InterruptedError:
        raise
    except ImportError as e:
        logger.error(f"缺少必要的库: {e}")
        raise RuntimeError(
            "XPS转PDF需要pymupdf库。\n"
            "请安装: pip install pymupdf"
        )
    except Exception as e:
        logger.error(f"XPS转PDF失败: {e}", exc_info=True)
        raise RuntimeError(f"XPS转PDF失败: {e}")


def ofd_to_pdf(ofd_path: str, cancel_event: Optional[threading.Event] = None, output_dir: Optional[str] = None) -> str:
    """
    将OFD文件转换为PDF格式
    
    OFD (Open Fixed-layout Document) 是中国国家电子文件标准格式，
    此函数使用easyofd库将其转换为更通用的PDF格式。
    
    参数:
        ofd_path: OFD文件路径
        cancel_event: 取消事件（可选）
        output_dir: 输出目录（可选）。如果为None，输出到原文件同目录；如果指定，输出到该目录（通常是临时目录）
        
    返回:
        str: 转换后的PDF文件路径
        
    异常:
        RuntimeError: 转换失败时抛出
        
    说明:
        使用easyofd库进行转换。
        转换流程：
        1. 使用easyofd.OFD类读取OFD文件
        2. 使用to_pdf()方法转换为PDF字节流
        3. 保存为PDF文件
        
        注意：统一使用标准化命名规则（带时间戳），无论是否在临时目录中
    """
    try:
        from easyofd import OFD
        
        # Monkey Patch: 修复 easyofd 库中 draw_annotation 方法的已知 Bug
        # 原版存在拼写错误 "ImgageObject" 且缺乏空值检查，可能导致 AttributeError
        try:
            from easyofd.draw.draw_pdf import DrawPDF
            
            def draw_annotation_patched(self, canvas, annota_info, images, page_size):
                """
                写入注释 (Monkey Patched 增强版)
                修复了 easyofd 原版中 ImgageObject 拼写错误导致的潜在空指针异常
                """
                try:
                    img_list = []
                    if not annota_info:
                        return

                    for key, annotation in annota_info.items():
                        try:
                            if not annotation: 
                                continue
                                
                            anno_type_obj = annotation.get("AnnoType")
                            if not anno_type_obj or anno_type_obj.get("type") != "Stamp":
                                continue

                            # 兼容 ImgageObject (源码拼写错误) 和 ImageObject (正确拼写)
                            img_obj = annotation.get("ImgageObject") or annotation.get("ImageObject")
                            if not img_obj:
                                continue

                            # 安全获取数据
                            boundary_str = img_obj.get("Boundary") or ""
                            pos_str = boundary_str.split(" ") if boundary_str else []
                            pos = [float(i) for i in pos_str] if pos_str else []

                            appearance = annotation.get("Appearance") or {}
                            wrap_boundary_str = appearance.get("Boundary") or ""
                            wrap_pos_str = wrap_boundary_str.split(" ") if wrap_boundary_str else []
                            wrap_pos = [float(i) for i in wrap_pos_str] if wrap_pos_str else []

                            ctm_str = img_obj.get("CTM") or ""
                            ctm_split = ctm_str.split(" ") if ctm_str else []
                            CTM = [float(i) for i in ctm_split] if ctm_split else []

                            img_list.append({
                                "wrap_pos": wrap_pos,
                                "pos": pos,
                                "CTM": CTM,
                                "ResourceID": img_obj.get("ResourceID", ""),
                            })
                        except Exception:
                            # 忽略单个注释的解析错误，避免影响整体转换
                            continue

                    if hasattr(self, 'draw_img'):
                        self.draw_img(canvas, img_list, images, page_size)
                except Exception as e:
                    # 记录 patch 方法本身的异常，但不抛出，尽量让流程继续
                    logger.warning(f"处理OFD注释时发生异常 (Patched): {e}")

            # 应用 Patch
            DrawPDF.draw_annotation = draw_annotation_patched
            logger.debug("已应用 easyofd.draw_annotation Monkey Patch")
            
        except ImportError:
            logger.warning("无法导入 easyofd.draw.draw_pdf，Monkey Patch 未应用")
        except Exception as e:
            logger.warning(f"应用 Monkey Patch 失败: {e}")

        logger.info(f"开始转换OFD文件: {os.path.basename(ofd_path)}")
        
        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("OFD转PDF操作被取消")
            raise InterruptedError("操作已取消")
        
        # 使用path_utils生成输出路径（统一使用标准化命名）
        from gongwen_converter.utils.path_utils import generate_output_path
        
        pdf_path = generate_output_path(
            ofd_path,
            output_dir=output_dir,
            section="",
            add_timestamp=True,
            description="fromOfd",
            file_type="pdf"
        )
        
        logger.debug(f"输出PDF路径: {pdf_path}")
        
        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("OFD转PDF操作被取消")
            raise InterruptedError("操作已取消")
        
        # 读取OFD文件
        ofd = OFD()
        
        try:
            ofd.read(str(ofd_path), fmt='path')
            logger.debug(f"成功读取OFD文件")
        except AttributeError as e:
            logger.warning(f"读取OFD时遇到AttributeError（可能是easyofd库的已知问题）: {e}")
            # 继续尝试转换，因为某些AttributeError不影响最终结果
        
        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("OFD转PDF操作被取消")
            raise InterruptedError("操作已取消")
        
        # 转换为PDF字节流
        try:
            pdf_bytes = ofd.to_pdf()
            logger.debug(f"OFD已转换为PDF字节流，大小: {len(pdf_bytes)} 字节")
        except AttributeError as e:
            # easyofd库存在已知问题：draw_annotation函数中当annotation.get("ImgageObject")返回None时
            # 会尝试对None调用.get()方法，导致AttributeError
            logger.warning(
                f"OFD转PDF过程中遇到AttributeError（easyofd库的已知问题）: {e}\n"
                f"这通常发生在处理某些OFD注释对象时，但可能不影响最终转换结果。"
            )
            # 尝试继续，因为错误可能只影响部分注释，PDF主体内容可能仍然可用
            # 如果确实失败，下面的写入步骤会检测到
            if not hasattr(ofd, '_pdf_bytes') or not pdf_bytes:
                logger.error("由于easyofd内部错误，无法生成PDF字节流")
                raise RuntimeError(
                    "OFD转PDF失败：easyofd库内部错误（AttributeError）。\n"
                    "建议：\n"
                    "1. 升级easyofd库到最新版本：pip install --upgrade easyofd\n"
                    "2. 或尝试使用其他OFD查看器手动转换"
                )
        
        # 检查取消事件
        if cancel_event and cancel_event.is_set():
            logger.info("OFD转PDF操作被取消")
            raise InterruptedError("操作已取消")
        
        # 写入PDF文件
        with open(pdf_path, 'wb') as f:
            f.write(pdf_bytes)
        
        logger.info(f"成功转换为PDF: {os.path.basename(pdf_path)}")
        return pdf_path
        
    except InterruptedError:
        raise
    except ImportError as e:
        logger.error(f"缺少必要的库: {e}")
        raise RuntimeError(
            "OFD转PDF需要easyofd库。\n"
            "请安装: pip install easyofd"
        )
    except Exception as e:
        logger.error(f"OFD转PDF失败: {e}", exc_info=True)
        raise RuntimeError(f"OFD转PDF失败: {e}")


# 模块测试
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s:%(lineno)d - %(message)s",
        handlers=[
            logging.StreamHandler()
        ]
    )
    
    logger.info("版式文件支持模块测试开始")
    
    # 测试CAJ转PDF（预期失败）
    try:
        logger.info("测试CAJ转PDF...")
        result = caj_to_pdf("test.caj")
    except NotImplementedError as e:
        logger.info(f"预期的NotImplementedError: {e}")
    
    # 测试OFD转PDF（预期失败）
    try:
        logger.info("测试OFD转PDF...")
        result = ofd_to_pdf("test.ofd")
    except NotImplementedError as e:
        logger.info(f"预期的NotImplementedError: {e}")
    
    logger.info("版式文件支持模块测试结束")
