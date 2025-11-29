"""
Office 格式转换核心模块

支持 Office 文档和表格的格式转换，基于统一临时工作空间架构。

核心特性：
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. 所有转换都在临时目录中进行
2. 支持三级容错机制（WPS/Office/LibreOffice）
3. 原文件完全隔离，永不被修改
4. 自动处理扩展名不匹配的文件
5. 支持三级降级保存策略
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

支持的转换：
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
文档格式：
  DOC ⇄ DOCX, WPS → DOCX, RTF ⇄ DOCX, ODT ⇄ DOCX
  
表格格式：
  XLS ⇄ XLSX, ET → XLSX, ODS ⇄ XLSX
  
导出格式：
  DOCX/DOC/WPS → PDF
  XLSX/XLS/ET/CSV → PDF
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

import os
import logging
import win32com.client
import pythoncom
import subprocess
import shutil
import threading
from typing import Optional, Callable, Tuple, List
from gongwen_converter.config.config_manager import config_manager
from gongwen_converter.config.constants import SOFTWARE_ID_MAPPING

logger = logging.getLogger(__name__)


# ==================== 自定义异常 ====================

class OfficeSoftwareNotFoundError(Exception):
    """当系统中找不到所需的Office或WPS软件时抛出"""
    pass


# ==================== Office软件可用性检测 ====================

def _check_wps_word_available() -> bool:
    """
    检查WPS Writer是否可用
    
    返回:
        bool: True 表示可用，False 表示不可用
    """
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Kwps.Application")
        app.Quit()
        pythoncom.CoUninitialize()
        logger.debug("检测到 WPS Writer 可用")
        return True
    except Exception as e:
        logger.debug(f"WPS Writer 不可用: {e}")
        return False


def _check_wps_spreadsheets_available() -> bool:
    """
    检查WPS Spreadsheets是否可用
    
    返回:
        bool: True 表示可用，False 表示不可用
    """
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Ket.Application")
        app.Quit()
        pythoncom.CoUninitialize()
        logger.debug("检测到 WPS Spreadsheets 可用")
        return True
    except Exception as e:
        logger.debug(f"WPS Spreadsheets 不可用: {e}")
        return False


def _check_office_word_available() -> bool:
    """
    检查Microsoft Word是否可用
    
    返回:
        bool: True 表示可用，False 表示不可用
    """
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Word.Application")
        app.Quit()
        pythoncom.CoUninitialize()
        logger.debug("检测到 Microsoft Word 可用")
        return True
    except Exception as e:
        logger.debug(f"Microsoft Word 不可用: {e}")
        return False


def _check_office_excel_available() -> bool:
    """
    检查Microsoft Excel是否可用
    
    返回:
        bool: True 表示可用，False 表示不可用
    """
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("Excel.Application")
        app.Quit()
        pythoncom.CoUninitialize()
        logger.debug("检测到 Microsoft Excel 可用")
        return True
    except Exception as e:
        logger.debug(f"Microsoft Excel 不可用: {e}")
        return False


def _check_libreoffice_available() -> bool:
    """
    检查系统中是否安装了 LibreOffice
    
    返回:
        bool: True 表示可用，False 表示不可用
    """
    try:
        result = subprocess.run(
            ["soffice", "--version"],
            capture_output=True,
            timeout=5,
            text=True
        )
        if result.returncode == 0:
            logger.debug(f"检测到 LibreOffice: {result.stdout.strip()}")
            return True
        return False
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False
    except Exception as e:
        logger.debug(f"检测 LibreOffice 时出错: {e}")
        return False


# 可用性检测结果缓存
_office_availability_cache = {}


def check_office_availability(target_format: str) -> Tuple[bool, str]:
    """
    检查是否有可用的Office软件来完成目标格式转换
    
    此函数会检测系统中是否安装了必要的办公软件（WPS/Office/LibreOffice），
    并根据目标格式返回相应的检测结果。检测结果会被缓存以提高性能。
    
    参数:
        target_format: 目标格式（如'doc', 'xls', 'wps', 'et', 'odt', 'ods'）
    
    返回:
        Tuple[bool, str]: (是否可用, 错误消息)
        - 如果有可用软件，返回 (True, "")
        - 如果无可用软件，返回 (False, 详细错误消息)
    
    示例:
        >>> available, error_msg = check_office_availability('doc')
        >>> if not available:
        ...     print(error_msg)
    """
    # 不需要Office软件的格式，直接返回True
    if target_format not in ['doc', 'xls', 'wps', 'et', 'rtf', 'odt', 'ods']:
        return True, ""
    
    # 检查缓存
    if target_format in _office_availability_cache:
        logger.debug(f"使用缓存的检测结果: {target_format}")
        return _office_availability_cache[target_format]
    
    # 确定需要检测的软件类型
    if target_format in ['doc', 'wps', 'rtf']:
        # 文档格式：检测Word类软件
        if _check_wps_word_available():
            result = (True, "")
        elif _check_office_word_available():
            result = (True, "")
        elif _check_libreoffice_available():
            result = (True, "")
        else:
            error_msg = (
                f"无法转换为 {target_format.upper()} 格式\n\n"
                "未检测到可用的文档处理软件。\n"
                "需要安装以下任一软件：\n"
                "• WPS Office\n"
                "• Microsoft Office\n"
                "• LibreOffice\n\n"
                "请安装后重试。"
            )
            result = (False, error_msg)
    
    elif target_format in ['xls', 'et']:
        # 表格格式：检测Excel类软件
        if _check_wps_spreadsheets_available():
            result = (True, "")
        elif _check_office_excel_available():
            result = (True, "")
        elif _check_libreoffice_available():
            result = (True, "")
        else:
            error_msg = (
                f"无法转换为 {target_format.upper()} 格式\n\n"
                "未检测到可用的表格处理软件。\n"
                "需要安装以下任一软件：\n"
                "• WPS Office\n"
                "• Microsoft Office\n"
                "• LibreOffice\n\n"
                "请安装后重试。"
            )
            result = (False, error_msg)
    
    elif target_format == 'odt':
        # ODT格式：只支持Microsoft Word和LibreOffice（WPS不支持ODT）
        if _check_office_word_available():
            result = (True, "")
        elif _check_libreoffice_available():
            result = (True, "")
        else:
            error_msg = (
                f"无法转换为 {target_format.upper()} 格式\n\n"
                "未检测到兼容软件。\n"
                "ODT格式需要：\n"
                "• Microsoft Word\n"
                "• LibreOffice\n\n"
                "注意：WPS Office 不支持此格式\n"
                "请安装上述软件后重试。"
            )
            result = (False, error_msg)
    
    elif target_format == 'ods':
        # ODS格式：只支持Microsoft Excel和LibreOffice（WPS不支持ODS）
        if _check_office_excel_available():
            result = (True, "")
        elif _check_libreoffice_available():
            result = (True, "")
        else:
            error_msg = (
                f"无法转换为 {target_format.upper()} 格式\n\n"
                "未检测到兼容软件。\n"
                "ODS格式需要：\n"
                "• Microsoft Excel\n"
                "• LibreOffice\n\n"
                "注意：WPS Office 不支持此格式\n"
                "请安装上述软件后重试。"
            )
            result = (False, error_msg)
    
    else:
        result = (True, "")
    
    # 缓存结果
    _office_availability_cache[target_format] = result
    logger.debug(f"缓存检测结果: {target_format} -> {result[0]}")
    
    return result


# ==================== LibreOffice 支持 ====================


def _convert_with_libreoffice(
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


def _show_software_not_found_dialog():
    """显示友好的软件未找到提示对话框"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()
        
        messagebox.showerror(
            "格式转换失败",
            "无法完成格式转换，可能的原因：\n\n"
            "❌ 未安装办公软件\n"
            "   • WPS Office\n"
            "   • Microsoft Office\n"
            "   • LibreOffice\n\n"
            "❌ 或已安装但版本不支持此转换\n\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "建议解决方案：\n\n"
            "1. 安装任一办公软件\n"
            "2. 确保软件版本支持该格式转换\n"
            "3. 重启软件后重试"
        )
        
        root.destroy()
    except Exception as e:
        logger.error(f"显示错误对话框失败: {e}")


# ==================== 通用三级容错转换框架 ====================

def _build_com_converters(software_list: List[str], converter_func: Callable) -> List[Tuple[str, str, Callable]]:
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


def _convert_with_fallback(
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
    3. 全部失败时显示提示对话框
    
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
        libre_result = _convert_with_libreoffice(
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
    
    # 全部失败 - 不在后台线程中显示对话框，让主线程处理错误显示
    logger.error("✗ 所有转换方法均失败")
    # 移除：_show_software_not_found_dialog() - 避免在后台线程中创建GUI对话框
    # 错误将通过返回的None和日志信息在主线程中正确处理
    return None, None


def _try_com_conversion(
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


# ==================== COM 转换器工厂函数 ====================

def _make_xlsx_to_ods_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """XLSX → ODS COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 60, 'excel')


def _make_ods_to_xlsx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """ODS → XLSX COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 51, 'excel')


def _make_xlsx_to_xls_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """XLSX → XLS COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 56, 'excel')


def _make_xls_to_xlsx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """XLS → XLSX COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 51, 'excel')


def _make_doc_to_docx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """DOC→DOCX 和 WPS→DOCX COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 12, 'word')


def _make_odt_to_docx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """ODT → DOCX COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 12, 'word')


def _make_rtf_to_docx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """RTF → DOCX COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 12, 'word')


def _make_docx_to_odt_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """DOCX → ODT COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 23, 'word')


def _make_docx_to_doc_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """DOCX → DOC COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 0, 'word')


def _make_docx_to_rtf_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """DOCX → RTF COM 转换器"""
    return _try_com_conversion(input_path, output_path, prog_id, 6, 'word')


# ==================== 文档格式转换 ====================

def office_to_docx(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    software_callback: Optional[Callable[[str], None]] = None,
    **kwargs
) -> Optional[str]:
    """
    DOC→DOCX 和 WPS→DOCX 单向转换
    
    支持将旧版文档格式转换为DOCX格式
    
    参数:
        input_path: 输入文件路径（扩展名可能不匹配）
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（用于扩展名修正，可选）
        cancel_event: 取消事件（可选）
        software_callback: 软件使用回调函数（可选）
        
    返回:
        成功时返回output_path，失败时返回None
        
    注意:
        如需扩展名修正，由调用者使用 prepare_input_file()
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → DOCX")
    
    # 使用配置管理器获取文档处理软件优先级
    software_priority = config_manager.get_word_processors_priority()
    logger.debug(f"文档处理软件优先级: {software_priority}")
    
    converters = _build_com_converters(software_priority, _make_doc_to_docx_converter)
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='docx',
        cancel_event=cancel_event,
        software_callback=software_callback
    )
    
    if result:
        logger.info(f"DOCX转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    else:
        logger.error("DOCX转换失败")
    
    return result


def rtf_to_docx(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    **kwargs
) -> Optional[str]:
    """
    RTF → DOCX 转换
    
    三级容错：WPS → Word → LibreOffice
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（可选）
        cancel_event: 取消事件（可选）
        
    返回:
        成功时返回output_path，失败时返回None
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → DOCX")
    
    converters = [
        ("WPS", "Kwps.Application", _make_rtf_to_docx_converter),
        ("Microsoft Word", "Word.Application", _make_rtf_to_docx_converter)
    ]
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='docx',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ DOCX 转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    
    return result


def odt_to_docx(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    **kwargs
) -> Optional[str]:
    """
    ODT → DOCX 转换
    
    二级容错：Word → LibreOffice（不使用 WPS，WPS不支持ODT）
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（可选）
        cancel_event: 取消事件（可选）
        
    返回:
        成功时返回output_path，失败时返回None
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → DOCX")
    
    converters = [
        ("Microsoft Word", "Word.Application", _make_odt_to_docx_converter)
    ]
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='docx',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ DOCX 转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    
    return result


def docx_to_odt(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    **kwargs
) -> Optional[str]:
    """
    DOCX → ODT 转换
    
    二级容错：Word → LibreOffice（不使用 WPS）
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（可选）
        cancel_event: 取消事件（可选）
        
    返回:
        成功时返回output_path，失败时返回None
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → ODT")
    
    converters = [
        ("Microsoft Word", "Word.Application", _make_docx_to_odt_converter)
    ]
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='odt',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ ODT 转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    
    return result


def convert_docx_to_doc(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    **kwargs
) -> Optional[str]:
    """
    DOCX → DOC 转换
    
    三级容错：WPS → Word → LibreOffice
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（可选）
        cancel_event: 取消事件（可选）
        progress_callback: 进度回调（可选）
        
    返回:
        成功时返回output_path，失败时返回None
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → DOC")
    
    if progress_callback:
        progress_callback("正在转换为 DOC 格式...")
    
    converters = [
        ("WPS", "Kwps.Application", _make_docx_to_doc_converter),
        ("Microsoft Word", "Word.Application", _make_docx_to_doc_converter)
    ]
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='doc',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ DOC 转换完成: {os.path.basename(result)}")
    
    return result


def convert_docx_to_rtf(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    **kwargs
) -> Optional[str]:
    """
    DOCX → RTF 转换
    
    三级容错：WPS → Word → LibreOffice
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（可选）
        cancel_event: 取消事件（可选）
        progress_callback: 进度回调（可选）
        
    返回:
        成功时返回output_path，失败时返回None
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → RTF")
    
    if progress_callback:
        progress_callback("正在转换为 RTF 格式...")
    
    converters = [
        ("WPS", "Kwps.Application", _make_docx_to_rtf_converter),
        ("Microsoft Word", "Word.Application", _make_docx_to_rtf_converter)
    ]
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='rtf',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ RTF 转换完成: {os.path.basename(result)}")
    
    return result


# ==================== 表格格式转换 ====================

def office_to_xlsx(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    **kwargs
) -> Optional[str]:
    """
    XLS→XLSX 和 ET→XLSX 单向转换
    
    支持将旧版表格格式转换为XLSX格式
    
    参数:
        input_path: 输入文件路径（扩展名可能不匹配）
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（用于扩展名修正，可选）
        cancel_event: 取消事件（可选）
        
    返回:
        成功时返回output_path，失败时返回None
        
    注意:
        如需扩展名修正，由调用者使用 prepare_input_file()
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → XLSX")
    
    converters = [
        ("WPS Spreadsheets", "Ket.Application", _make_xls_to_xlsx_converter),
        ("Microsoft Excel", "Excel.Application", _make_xls_to_xlsx_converter)
    ]
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='xlsx',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"XLSX转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    
    return result


def xlsx_to_ods(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    **kwargs
) -> Optional[str]:
    """
    XLSX → ODS 转换
    
    二级容错：Excel → LibreOffice（不使用 WPS，WPS不支持ODS）
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（可选）
        cancel_event: 取消事件（可选）
        
    返回:
        成功时返回output_path，失败时返回None
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → ODS")
    
    converters = [
        ("Microsoft Excel", "Excel.Application", _make_xlsx_to_ods_converter)
    ]
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='ods',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ ODS 转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    
    return result


def ods_to_xlsx(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    **kwargs
) -> Optional[str]:
    """
    ODS → XLSX 转换
    
    二级容错：Excel → LibreOffice（不使用 WPS）
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（可选）
        cancel_event: 取消事件（可选）
        
    返回:
        成功时返回output_path，失败时返回None
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → XLSX")
    
    converters = [
        ("Microsoft Excel", "Excel.Application", _make_ods_to_xlsx_converter)
    ]
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='xlsx',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ XLSX 转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    
    return result


def convert_xlsx_to_xls(
    input_path: str,
    output_path: str,
    actual_format: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    **kwargs
) -> Optional[str]:
    """
    XLSX → XLS 转换
    
    三级容错：WPS → Excel → LibreOffice
    
    参数:
        input_path: 输入文件路径
        output_path: 输出文件路径（完整路径，包含符合规则的文件名）
        actual_format: 文件真实格式（可选）
        cancel_event: 取消事件（可选）
        progress_callback: 进度回调（可选）
        
    返回:
        成功时返回output_path，失败时返回None
    """
    logger.info(f"开始转换: {os.path.basename(input_path)} → XLS")
    
    if progress_callback:
        progress_callback("正在转换为 XLS 格式...")
    
    converters = [
        ("WPS Spreadsheets", "Ket.Application", _make_xlsx_to_xls_converter),
        ("Microsoft Excel", "Excel.Application", _make_xlsx_to_xls_converter)
    ]
    
    result, used_software = _convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='xls',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ XLS 转换完成: {os.path.basename(result)}")
    
    return result


# ==================== PDF 导出转换器 ====================

def _make_word_to_pdf_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """Word/WPS → PDF COM 转换器（FileFormat=17）"""
    return _try_com_conversion(input_path, output_path, prog_id, 17, 'word')


def _make_excel_to_pdf_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """Excel/WPS Spreadsheets → PDF COM 转换器（FileFormat=57）"""
    return _try_com_conversion(input_path, output_path, prog_id, 57, 'excel')


def docx_to_pdf(
    input_path: str,
    output_path: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None
) -> Optional[str]:
    """
    文档转 PDF（支持 DOCX/DOC/WPS/RTF/ODT 输入）
    
    转换流程：
    1. 如果输入不是 DOCX，先转换为 DOCX
    2. 将 DOCX 转换为 PDF
    3. 使用 office_to_pdf 配置的软件优先级
    
    参数:
        input_path: 输入文件路径（支持 DOCX/DOC/WPS/RTF/ODT）
        output_path: 输出PDF文件路径（可选，如不提供则自动生成）
        cancel_event: 取消事件（可选）
        
    返回:
        成功时返回输出PDF文件路径，失败时返回 None
        
    异常:
        OfficeSoftwareNotFoundError: 未找到可用的 Office 软件
    """
    import tempfile
    from gongwen_converter.utils.file_type_utils import detect_actual_file_format
    from gongwen_converter.utils.path_utils import generate_output_path
    
    # 检测实际文件格式
    actual_format = detect_actual_file_format(input_path)
    logger.info(f"文档转 PDF: {os.path.basename(input_path)} (格式: {actual_format})")
    
    # 如果未提供输出路径，自动生成
    if output_path is None:
        output_path = generate_output_path(
            input_path,
            section="",
            add_timestamp=True,
            description=f"from{actual_format.capitalize()}",
            file_type="pdf"
        )
        logger.debug(f"自动生成输出路径: {output_path}")
    
    try:
        # 使用临时目录处理中间文件
        with tempfile.TemporaryDirectory() as temp_dir:
            # 步骤1：如果不是 DOCX，先转换为 DOCX
            if actual_format != 'docx':
                logger.info(f"需要先将 {actual_format.upper()} 转换为 DOCX")
                
                # 选择合适的转换函数
                if actual_format in ['doc', 'wps']:
                    converter = office_to_docx
                elif actual_format == 'rtf':
                    converter = rtf_to_docx
                elif actual_format == 'odt':
                    converter = odt_to_docx
                else:
                    raise ValueError(f"不支持的文档格式: {actual_format}")
                
                # 生成临时 DOCX 文件路径
                temp_docx = os.path.join(temp_dir, "intermediate.docx")
                
                # 转换为 DOCX
                docx_result = converter(
                    input_path,
                    temp_docx,
                    actual_format=actual_format,
                    cancel_event=cancel_event
                )
                
                if not docx_result or not os.path.exists(docx_result):
                    logger.error(f"{actual_format.upper()} 转 DOCX 失败")
                    raise OfficeSoftwareNotFoundError(f"无法将 {actual_format.upper()} 转换为 DOCX")
                
                # 使用转换后的 DOCX 作为输入
                pdf_input = docx_result
                logger.info(f"✓ {actual_format.upper()} → DOCX 转换成功")
            else:
                # 已是 DOCX，直接使用
                pdf_input = input_path
                logger.debug("输入已是 DOCX 格式，直接转 PDF")
            
            if cancel_event and cancel_event.is_set():
                logger.info("转换被取消")
                return None
            
            # 步骤2：DOCX → PDF
            logger.info("开始转换 DOCX → PDF")
            
            # 获取文档转PDF优先级配置
            software_priority = config_manager.get_document_to_pdf_priority()
            logger.debug(f"文档转PDF软件优先级: {software_priority}")
            
            # 构建转换器列表
            converters = _build_com_converters(software_priority, _make_word_to_pdf_converter)
            
            # 执行转换（LibreOffice 也支持 PDF 导出）
            result, used_software = _convert_with_fallback(
                input_path=pdf_input,
                output_path=output_path,
                com_converters=converters,
                libreoffice_format='pdf',
                cancel_event=cancel_event
            )
            
            if not result:
                logger.error("DOCX → PDF 转换失败")
                raise OfficeSoftwareNotFoundError("无法转换为 PDF")
            
            logger.info(f"✓ PDF 转换成功: {os.path.basename(result)} (使用软件: {used_software})")
            return result
            
    except OfficeSoftwareNotFoundError:
        raise
    except Exception as e:
        logger.error(f"文档转 PDF 失败: {e}", exc_info=True)
        raise OfficeSoftwareNotFoundError(f"转换失败: {e}")


def xlsx_to_pdf(
    input_path: str,
    output_path: Optional[str] = None,
    cancel_event: Optional[threading.Event] = None
) -> Optional[str]:
    """
    表格转 PDF（支持 XLSX/XLS/ET/ODS/CSV 输入）
    
    转换流程：
    1. 如果输入不是 XLSX，先转换为 XLSX
    2. 将 XLSX 转换为 PDF
    3. 使用 office_to_pdf 配置的软件优先级
    
    参数:
        input_path: 输入文件路径（支持 XLSX/XLS/ET/ODS/CSV）
        output_path: 输出PDF文件路径（可选，如不提供则自动生成）
        cancel_event: 取消事件（可选）
        
    返回:
        成功时返回输出PDF文件路径，失败时返回 None
        
    异常:
        OfficeSoftwareNotFoundError: 未找到可用的 Office 软件
    """
    import tempfile
    from gongwen_converter.utils.file_type_utils import detect_actual_file_format
    from gongwen_converter.utils.path_utils import generate_output_path
    
    # 检测实际文件格式
    actual_format = detect_actual_file_format(input_path)
    logger.info(f"表格转 PDF: {os.path.basename(input_path)} (格式: {actual_format})")
    
    # 如果未提供输出路径，自动生成
    if output_path is None:
        output_path = generate_output_path(
            input_path,
            section="",
            add_timestamp=True,
            description=f"from{actual_format.capitalize()}",
            file_type="pdf"
        )
        logger.debug(f"自动生成输出路径: {output_path}")
    
    try:
        # 使用临时目录处理中间文件
        with tempfile.TemporaryDirectory() as temp_dir:
            # 步骤1：如果不是 XLSX，先转换为 XLSX
            if actual_format not in ['xlsx', 'csv']:
                logger.info(f"需要先将 {actual_format.upper()} 转换为 XLSX")
                
                # 选择合适的转换函数
                if actual_format in ['xls', 'et']:
                    converter = office_to_xlsx
                elif actual_format == 'ods':
                    converter = ods_to_xlsx
                else:
                    raise ValueError(f"不支持的表格格式: {actual_format}")
                
                # 生成临时 XLSX 文件路径
                temp_xlsx = os.path.join(temp_dir, "intermediate.xlsx")
                
                # 转换为 XLSX
                xlsx_result = converter(
                    input_path,
                    temp_xlsx,
                    actual_format=actual_format,
                    cancel_event=cancel_event
                )
                
                if not xlsx_result or not os.path.exists(xlsx_result):
                    logger.error(f"{actual_format.upper()} 转 XLSX 失败")
                    raise OfficeSoftwareNotFoundError(f"无法将 {actual_format.upper()} 转换为 XLSX")
                
                # 使用转换后的 XLSX 作为输入
                pdf_input = xlsx_result
                logger.info(f"✓ {actual_format.upper()} → XLSX 转换成功")
            elif actual_format == 'csv':
                # CSV 需要先转换为 XLSX
                logger.info("需要先将 CSV 转换为 XLSX")
                from gongwen_converter.converter.table_converters import convert_csv_to_xlsx
                
                temp_xlsx = os.path.join(temp_dir, "intermediate.xlsx")
                xlsx_result = convert_csv_to_xlsx(input_path, temp_xlsx)
                
                if not xlsx_result or not os.path.exists(xlsx_result):
                    logger.error("CSV 转 XLSX 失败")
                    raise OfficeSoftwareNotFoundError("无法将 CSV 转换为 XLSX")
                
                pdf_input = xlsx_result
                logger.info("✓ CSV → XLSX 转换成功")
            else:
                # 已是 XLSX，直接使用
                pdf_input = input_path
                logger.debug("输入已是 XLSX 格式，直接转 PDF")
            
            if cancel_event and cancel_event.is_set():
                logger.info("转换被取消")
                return None
            
            # 步骤2：XLSX → PDF
            logger.info("开始转换 XLSX → PDF")
            
            # 获取表格转PDF优先级配置
            software_priority = config_manager.get_spreadsheet_to_pdf_priority()
            logger.debug(f"表格转PDF软件优先级: {software_priority}")
            
            # 构建转换器列表
            converters = _build_com_converters(software_priority, _make_excel_to_pdf_converter)
            
            # 执行转换（LibreOffice 也支持 PDF 导出）
            result, used_software = _convert_with_fallback(
                input_path=pdf_input,
                output_path=output_path,
                com_converters=converters,
                libreoffice_format='pdf',
                cancel_event=cancel_event
            )
            
            if not result:
                logger.error("XLSX → PDF 转换失败")
                raise OfficeSoftwareNotFoundError("无法转换为 PDF")
            
            logger.info(f"✓ PDF 转换成功: {os.path.basename(result)} (使用软件: {used_software})")
            return result
            
    except OfficeSoftwareNotFoundError:
        raise
    except Exception as e:
        logger.error(f"表格转 PDF 失败: {e}", exc_info=True)
        raise OfficeSoftwareNotFoundError(f"转换失败: {e}")
