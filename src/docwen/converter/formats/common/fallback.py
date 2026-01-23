"""
三级容错转换框架

提供 WPS → Office → LibreOffice 的容错机制，
支持自定义转换器列表和回调。

依赖：
- win32com.client: Windows COM 接口
- pythoncom: COM 初始化
- subprocess: LibreOffice 命令行

使用方式:
    from docwen.converter.formats.common.fallback import (
        convert_with_fallback,
        build_com_converters,
    )
"""

import os
import logging
import shutil
import subprocess
import threading
from typing import Optional, Callable, Tuple, List

import win32com.client
import pythoncom

from docwen.config import SOFTWARE_ID_MAPPING

logger = logging.getLogger(__name__)


# ==================== LibreOffice 转换 ====================

def convert_with_libreoffice(
    input_path: str,
    output_format: str,
    output_dir: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None
) -> Optional[str]:
    """
    使用 LibreOffice 命令行进行格式转换
    
    参数:
        input_path: 输入文件路径
        output_format: 目标格式（如 'ods', 'xls', 'docx' 等）
        output_dir: 输出目录（可选）
        cancel_event: 取消事件（可选）
        
    返回:
        str: 转换后的文件路径，失败时返回 None
    """
    try:
        if cancel_event and cancel_event.is_set():
            logger.info("LibreOffice 转换被取消")
            return None
        
        if output_dir is None:
            output_dir = os.path.dirname(input_path)
        
        cmd = [
            "soffice",
            "--headless",
            "--convert-to", output_format,
            "--outdir", output_dir,
            input_path
        ]
        
        logger.info(f"使用 LibreOffice 转换: {os.path.basename(input_path)} → {output_format}")
        logger.debug(f"LibreOffice 命令: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, capture_output=True, timeout=60, text=True)
        
        if result.returncode == 0:
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            output_file = os.path.join(output_dir, f"{base_name}.{output_format}")
            
            if os.path.exists(output_file):
                logger.info(f"✓ LibreOffice 转换成功: {os.path.basename(output_file)}")
                return output_file
            else:
                logger.error(f"LibreOffice 执行成功但未找到输出文件")
                return None
        else:
            logger.error(f"LibreOffice 转换失败: {result.stderr}")
            return None
            
    except FileNotFoundError:
        logger.debug("LibreOffice 未安装")
        return None
    except subprocess.TimeoutExpired:
        logger.error("LibreOffice 转换超时")
        return None
    except Exception as e:
        logger.error(f"LibreOffice 转换出错: {e}")
        return None


# ==================== COM 转换器构建 ====================

def build_com_converters(
    software_list: List[str],
    converter_func: Callable
) -> List[Tuple[str, str, Callable]]:
    """
    根据软件优先级列表构建COM转换器列表
    
    参数:
        software_list: 软件标识符列表
        converter_func: 转换器函数
        
    返回:
        List[Tuple[str, str, Callable]]: 转换器列表 [(应用名, ProgID, 转换函数)]
    """
    converters = []
    for software_id in software_list:
        if software_id in SOFTWARE_ID_MAPPING:
            prog_id = SOFTWARE_ID_MAPPING[software_id]
            display_name = {
                'wps_writer': 'WPS',
                'wps_spreadsheets': 'WPS Spreadsheets',
                'msoffice_word': 'Microsoft Word',
                'msoffice_excel': 'Microsoft Excel'
            }.get(software_id, software_id)
            converters.append((display_name, prog_id, converter_func))
    return converters


# ==================== 三级容错转换框架 ====================

def convert_with_fallback(
    input_path: str,
    output_path: str,
    com_converters: List[Tuple[str, str, Callable]],
    libreoffice_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    software_callback: Optional[Callable[[str], None]] = None
) -> Tuple[Optional[str], Optional[str]]:
    """
    通用的三级容错转换框架
    
    转换优先级：
    1. COM 接口转换器（WPS/Office 等）
    2. LibreOffice 命令行
    3. 全部失败时返回 None
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径
        com_converters: COM 转换器列表 [(应用名, ProgID, 转换函数)]
        libreoffice_format: LibreOffice 格式代码（可选）
        cancel_event: 取消事件（可选）
        software_callback: 软件使用回调函数（可选）
        
    返回:
        Tuple[Optional[str], Optional[str]]: (输出文件路径, 使用的软件名称) 或 (None, None) 失败时
    """
    if cancel_event and cancel_event.is_set():
        logger.info("转换在开始前被取消")
        return None, None
    
    abs_input_path = os.path.abspath(input_path)
    abs_output_path = os.path.abspath(output_path)
    
    total_attempts = len(com_converters) + (1 if libreoffice_format else 0)
    current_attempt = 0
    
    # 尝试 COM 接口转换
    for app_name, prog_id, converter_func in com_converters:
        current_attempt += 1
        
        if cancel_event and cancel_event.is_set():
            logger.info("转换被取消")
            return None, None
        
        logger.info(f"[{current_attempt}/{total_attempts}] 尝试 {app_name}")
        
        try:
            result = converter_func(abs_input_path, abs_output_path, prog_id)
            if result:
                logger.info(f"✓ 转换成功 [{app_name}]")
                # 调用软件使用回调
                if software_callback:
                    software_callback(app_name)
                return output_path, app_name
        except Exception as e:
            logger.debug(f"{app_name} 转换失败: {e}")
            continue
    
    # 尝试 LibreOffice
    if libreoffice_format:
        current_attempt += 1
        
        if cancel_event and cancel_event.is_set():
            logger.info("转换被取消")
            return None, None
        
        logger.info(f"[{current_attempt}/{total_attempts}] 尝试 LibreOffice")
        
        output_dir = os.path.dirname(output_path)
        libre_result = convert_with_libreoffice(
            input_path,
            libreoffice_format,
            output_dir=output_dir,
            cancel_event=cancel_event
        )
        
        if libre_result:
            if libre_result != output_path and os.path.exists(libre_result):
                try:
                    shutil.move(libre_result, output_path)
                    logger.info(f"✓ 转换成功 [LibreOffice]")
                    # 调用软件使用回调
                    if software_callback:
                        software_callback("LibreOffice")
                    return output_path, "LibreOffice"
                except Exception as e:
                    logger.warning(f"重命名文件失败: {e}")
                    if software_callback:
                        software_callback("LibreOffice")
                    return libre_result, "LibreOffice"
            logger.info(f"✓ 转换成功 [LibreOffice]")
            if software_callback:
                software_callback("LibreOffice")
            return libre_result, "LibreOffice"
    
    # 全部失败
    logger.error("✗ 所有转换方法均失败")
    return None, None


# ==================== COM 转换核心函数 ====================

def try_com_conversion(
    input_path: str,
    output_path: str,
    prog_id: str,
    file_format: int,
    app_type: str = 'excel'
) -> Optional[str]:
    """
    使用 COM 接口进行格式转换
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径
        prog_id: COM 程序 ID（如 'Ket.Application'）
        file_format: Office 文件格式代码
        app_type: 应用类型（'excel' 或 'word'）
        
    返回:
        str: 成功时返回输出文件路径
        None: 失败时返回 None
    """
    app = None
    doc_or_wb = None
    
    try:
        pythoncom.CoInitialize()
        
        app = win32com.client.Dispatch(prog_id)
        app.Visible = False
        app.DisplayAlerts = False
        
        if app_type == 'excel':
            doc_or_wb = app.Workbooks.Open(input_path)
        else:
            doc_or_wb = app.Documents.Open(
                input_path,
                ReadOnly=True,
                ConfirmConversions=False,
                AddToRecentFiles=False
            )
        
        # 判断是否为PDF转换
        is_pdf_export = (
            (file_format == 57 and app_type == 'excel') or  # Excel PDF
            (file_format == 17 and app_type == 'word')       # Word PDF
        )
        
        if is_pdf_export:
            # PDF导出：统一使用ExportAsFixedFormat（标准方法，WPS/Office通用）
            if app_type == 'excel':
                # Excel/WPS Spreadsheets
                doc_or_wb.ExportAsFixedFormat(
                    Type=0,                    # 0=PDF, 1=XPS
                    Filename=output_path,
                    Quality=0,                 # 0=标准质量, 1=最小文件大小
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
                logger.debug(f"使用ExportAsFixedFormat导出Excel PDF: {os.path.basename(output_path)}")
            else:  # app_type == 'word'
                # Word/WPS Writer
                doc_or_wb.ExportAsFixedFormat(
                    OutputFileName=output_path,
                    ExportFormat=17,           # 17=PDF, 18=XPS
                    OpenAfterExport=False,
                    OptimizeFor=0,            # 0=标准质量, 1=最小文件大小
                    CreateBookmarks=1,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False
                )
                logger.debug(f"使用ExportAsFixedFormat导出Word PDF: {os.path.basename(output_path)}")
        else:
            # 其他格式转换：使用SaveAs
            doc_or_wb.SaveAs(output_path, FileFormat=file_format)
        
        return output_path
        
    except Exception as e:
        logger.debug(f"COM 转换出错: {e}")
        return None
    finally:
        if doc_or_wb:
            try:
                doc_or_wb.Close(SaveChanges=False)
            except:
                pass
        if app:
            try:
                app.Quit()
            except:
                pass
        try:
            pythoncom.CoUninitialize()
        except:
            pass
