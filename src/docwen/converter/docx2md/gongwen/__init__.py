"""
公文模式转换器模块

实现公文专用的转换逻辑：
- 三轮公文元素识别
- YAML元数据提取（14个字段）
- 附件内容单独输出
- 支持图片提取和OCR

模块结构：
- converter.py: 主转换入口
- content_generator.py: 内容生成器
- extractor.py: 公文元素提取
- scorer.py: 元素评分器

使用示例：
    from docwen.converter.docx2md.gongwen import convert_docx_to_md_gongwen
    
    result = convert_docx_to_md_gongwen('公文.docx')
    if result['success']:
        print(result['main_content'])
        if result['attachment_content']:
            print(result['attachment_content'])
"""

from .converter import convert_docx_to_md_gongwen
from .content_generator import generate_main_content, generate_attachment_content
from .extractor import (
    extract_doc_number_and_signers,
    extract_signers_from_text,
    extract_doc_number_and_name
)
from .scorer import create_element_scorer, ElementScorer

__all__ = [
    # 主转换函数
    'convert_docx_to_md_gongwen',
    # 内容生成
    'generate_main_content',
    'generate_attachment_content',
    # 元素提取
    'extract_doc_number_and_signers',
    'extract_signers_from_text',
    'extract_doc_number_and_name',
    # 评分器
    'create_element_scorer',
    'ElementScorer',
]
