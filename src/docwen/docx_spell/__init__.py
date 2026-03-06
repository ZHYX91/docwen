"""
docx_spell 包入口模块
提供错别字检测和批注添加功能
"""

# 导入核心功能
from .core import process_docx
from .spell_checker import TextError, TextValidator

# 定义公开接口
__all__ = ["TextError", "TextValidator", "process_docx"]

# 包初始化日志
import logging

logger = logging.getLogger(__name__)
logger.info("docx_spell包初始化完成 - 提供错别字检测和批注功能")
