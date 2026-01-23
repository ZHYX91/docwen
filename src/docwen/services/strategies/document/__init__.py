"""
文档转换策略子包

提供文档文件转换相关的策略类：
- DocxToMdStrategy: 文档转Markdown
- DocumentToPdfStrategy: 文档转PDF
- DocumentToOfdStrategy: 文档转OFD（占位）
- DocxValidationStrategy: 文档校对
- 动态格式转换策略（通过工厂批量注册）

辅助函数：
- _preprocess_document_file: 文档文件预处理
"""

# 导入辅助函数
from .utils import _preprocess_document_file

# 导入策略类（自动注册）
from .to_markdown import DocxToMdStrategy
from .to_pdf import DocumentToPdfStrategy, DocumentToOfdStrategy
from .validation import DocxValidationStrategy

# 导入格式转换模块（触发动态策略批量注册）
from . import format_conversion

__all__ = [
    # 辅助函数
    '_preprocess_document_file',
    # 策略类
    'DocxToMdStrategy',
    'DocumentToPdfStrategy',
    'DocumentToOfdStrategy',
    'DocxValidationStrategy',
]
