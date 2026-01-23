"""
文件选择器对话框辅助模块

提供文件处理状态相关的对话框显示功能：
- 显示文件跳过原因对话框
- 显示文件处理失败详情对话框
"""

import os
import logging
import subprocess
from tkinter import messagebox
from typing import TYPE_CHECKING

from docwen.i18n import t

if TYPE_CHECKING:
    from .file_info import FileInfo

logger = logging.getLogger(__name__)


def show_skip_dialog(master, file_info: 'FileInfo') -> None:
    """
    显示文件跳过原因对话框
    
    当文件被跳过转换时，显示详细的跳过原因和文件位置信息，
    并询问用户是否打开文件所在位置。
    
    参数:
        master: 父窗口对象
        file_info: 被跳过的文件信息对象
    """
    skip_reason = file_info.skip_reason or t('components.file_dialogs.skip_reason_unknown')
    file_dir = os.path.dirname(file_info.file_path)
    
    message = t(
        'components.file_dialogs.skip_message',
        reason=skip_reason,
        filename=file_info.file_name,
        location=file_dir
    )
    
    logger.info(f"显示跳过对话框: {file_info.file_name}")
    
    # 显示信息对话框
    messagebox.showinfo(
        t('components.file_dialogs.skip_title'),
        message,
        parent=master
    )
    
    # 询问是否打开文件位置
    if messagebox.askyesno(
        t('components.status_bar.open_location_title'),
        t('components.status_bar.open_location_confirm'),
        parent=master
    ):
        open_file_location(file_info.file_path)


def show_error_dialog(master, file_info: 'FileInfo') -> None:
    """
    显示文件处理失败详情对话框
    
    当文件转换失败时，显示详细的错误信息和建议操作。
    
    参数:
        master: 父窗口对象
        file_info: 处理失败的文件信息对象
    """
    error_message = file_info.error_message or t('components.file_dialogs.error_unknown')
    
    message = t(
        'components.file_dialogs.error_message',
        filename=file_info.file_name,
        error=error_message
    )
    
    logger.info(f"显示错误对话框: {file_info.file_name}")
    
    # 显示错误对话框
    messagebox.showerror(
        t('components.file_dialogs.error_title'),
        message,
        parent=master
    )


def open_file_location(file_path: str) -> None:
    """
    打开文件所在文件夹并选中文件
    
    使用系统资源管理器打开文件所在目录，并自动选中该文件。
    
    参数:
        file_path: 要定位的文件路径
    """
    logger.debug(f"打开文件所在文件夹: {file_path}")
    
    try:
        if os.path.exists(file_path):
            # Windows: 使用 explorer /select 命令打开文件夹并选中文件
            subprocess.Popen(['explorer', '/select,', os.path.normpath(file_path)])
            logger.info(f"已打开文件夹并选中文件: {file_path}")
        else:
            logger.error(f"文件不存在: {file_path}")
    except Exception as e:
        logger.error(f"打开文件夹失败: {e}")
