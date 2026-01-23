"""
办公软件可用性检测模块

提供 WPS/Office/LibreOffice 的可用性检测功能，
支持结果缓存以提高性能。

依赖：
- win32com.client: Windows COM 接口
- pythoncom: COM 初始化
- subprocess: LibreOffice 检测
"""

import logging
import subprocess
from typing import Tuple

import win32com.client
import pythoncom

logger = logging.getLogger(__name__)


# ==================== 自定义异常 ====================

class OfficeSoftwareNotFoundError(Exception):
    """当系统中找不到所需的Office或WPS软件时抛出"""
    pass


# ==================== 软件可用性检测函数 ====================

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


# ==================== 可用性检测缓存 ====================

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
