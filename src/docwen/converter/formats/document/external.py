"""
文档格式转换 - 外部软件实现

使用 WPS/Office/LibreOffice 进行文档格式转换。

支持的转换：
- DOC/WPS → DOCX（预处理）
- RTF → DOCX
- ODT → DOCX
- DOCX → ODT
- DOCX → DOC
- DOCX → RTF

依赖：
- common.fallback: 三级容错机制
- config: 软件优先级配置
"""

import os
import logging
import threading
from typing import Optional, Callable

from docwen.config import config_manager
from docwen.i18n import t
from ..common.fallback import (
    convert_with_fallback,
    build_com_converters,
    try_com_conversion,
)

logger = logging.getLogger(__name__)


# ==================== COM 转换器工厂函数 ====================

def _make_doc_to_docx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """DOC→DOCX 和 WPS→DOCX COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 12, 'word')


def _make_odt_to_docx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """ODT → DOCX COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 12, 'word')


def _make_rtf_to_docx_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """RTF → DOCX COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 12, 'word')


def _make_docx_to_odt_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """DOCX → ODT COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 23, 'word')


def _make_docx_to_doc_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """DOCX → DOC COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 0, 'word')


def _make_docx_to_rtf_converter(input_path: str, output_path: str, prog_id: str) -> Optional[str]:
    """DOCX → RTF COM 转换器"""
    return try_com_conversion(input_path, output_path, prog_id, 6, 'word')


# ==================== 文档格式转换函数 ====================

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
    
    converters = build_com_converters(software_priority, _make_doc_to_docx_converter)
    
    result, used_software = convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='docx',
        cancel_event=cancel_event,
        software_callback=software_callback
    )
    
    if result:
        logger.info(f"✓ DOCX转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    else:
        logger.error("✗ DOCX转换失败")
    
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
    
    result, used_software = convert_with_fallback(
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
    
    result, used_software = convert_with_fallback(
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
    
    result, used_software = convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='odt',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ ODT 转换完成: {os.path.basename(result)} (使用软件: {used_software})")
    
    return result


def docx_to_doc(
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
        progress_callback(t('conversion.progress.converting_to_format', format='DOC'))
    
    converters = [
        ("WPS", "Kwps.Application", _make_docx_to_doc_converter),
        ("Microsoft Word", "Word.Application", _make_docx_to_doc_converter)
    ]
    
    result, used_software = convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='doc',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ DOC 转换完成: {os.path.basename(result)}")
    
    return result


def docx_to_rtf(
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
        progress_callback(t('conversion.progress.converting_to_format', format='RTF'))
    
    converters = [
        ("WPS", "Kwps.Application", _make_docx_to_rtf_converter),
        ("Microsoft Word", "Word.Application", _make_docx_to_rtf_converter)
    ]
    
    result, used_software = convert_with_fallback(
        input_path=input_path,
        output_path=output_path,
        com_converters=converters,
        libreoffice_format='rtf',
        cancel_event=cancel_event
    )
    
    if result:
        logger.info(f"✓ RTF 转换完成: {os.path.basename(result)}")
    
    return result
