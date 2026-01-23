"""
表格格式转换 - 外部软件实现

使用 WPS/Excel/LibreOffice 进行表格格式转换。

支持的转换：
- XLS/ET → XLSX（预处理）
- ODS → XLSX
- XLSX → ODS
- XLSX → XLS

依赖：
- common.fallback: 三级容错机制
"""

import os
import logging
import threading
from typing import Optional, Callable

from ..common.fallback import (
    convert_with_fallback,
    try_com_conversion,
)
from docwen.i18n import t

logger = logging.getLogger(__name__)


# ==================== COM 转换器工厂函数 ====================

def _make_xlsx_to_ods_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """XLSX → ODS COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 60, 'excel')


def _make_ods_to_xlsx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """ODS → XLSX COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 51, 'excel')


def _make_xlsx_to_xls_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """XLSX → XLS COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 56, 'excel')


def _make_xls_to_xlsx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """XLS → XLSX COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 51, 'excel')


# ==================== 表格格式转换函数 ====================

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
    
    result, used_software = convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='xlsx',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ XLSX转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    
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
    
    result, used_software = convert_with_fallback(
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
    
    result, used_software = convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='xlsx',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ XLSX 转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    
    return result


def xlsx_to_xls(
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
        progress_callback(t('conversion.progress.converting_to_format', format='XLS'))
    
    converters = [
        ("WPS Spreadsheets", "Ket.Application", _make_xlsx_to_xls_converter),
        ("Microsoft Excel", "Excel.Application", _make_xlsx_to_xls_converter)
    ]
    
    result, used_software = convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='xls',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ XLS 转换完成: {os.path.basename(result)}")
    
    return result
