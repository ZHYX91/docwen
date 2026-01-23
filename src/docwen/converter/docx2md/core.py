"""
DOCX文档转Markdown格式核心处理模块

本模块负责将Word文档（DOCX格式）转换为Markdown格式，并生成详细的格式分析报告。
主要功能包括：
- 解析DOCX文档结构和样式信息
- 识别公文格式元素（标题、发文机关、日期等）
- 提取YAML格式的文档元数据
- 处理文本框、表格等特殊内容
- 将附件内容单独输出为独立的Markdown文件
- 生成格式分析报告和调试文档

转换过程采用智能分析算法，能够准确识别公文的标准格式元素。
"""

import os
import re
import logging
import datetime
from docx import Document
from .shared.content_injector import process_document_with_special_content
from docwen.utils.docx_utils import extract_format_from_paragraph_style, get_effective_run_format
from .gongwen.scorer import create_element_scorer
from docwen.utils.heading_utils import split_content_by_delimiters, add_markdown_heading, detect_heading_level
from docwen.utils.path_utils import ensure_dir_exists, generate_output_path
from docwen.utils.validation_utils import contains_chinese
from docwen.utils.date_utils import convert_date_format
from docwen.utils.text_utils import (
    remove_colon,
    extract_after_colon,
    remove_brackets,
    process_copy_to,
    process_attachment_item,
    is_pure_number
)

# 配置日志
logger = logging.getLogger(__name__)

import threading
from typing import Callable, Optional

def convert_docx_to_md(
    docx_path: str,
    extract_image: bool = True,
    extract_ocr: bool = False,
    optimize_for_type: str = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    cancel_event: Optional[threading.Event] = None,
    output_folder: Optional[str] = None,
    original_file_path: Optional[str] = None,
    options: Optional[dict] = None
):
    """
    DOCX转MD路由函数
    
    根据optimize_for_type参数选择转换模式：
    - None: 简化模式（默认） - 基础样式转换
    - "gongwen": 公文优化模式 - 元素识别+YAML元数据
    
    参数:
        docx_path: DOCX文件路径（可能是临时副本）
        extract_image: 是否保留图片（由GUI传入，默认True）
        extract_ocr: 是否进行OCR识别（由GUI传入，默认False）
        optimize_for_type: 优化类型（默认None为简化模式，"gongwen"为公文模式）
        progress_callback: 进度回调函数 (可选)
        cancel_event: 取消事件 (可选)
        output_folder: 输出文件夹路径，用于保存图片 (可选)
        original_file_path: 原始文件路径（用于图片命名，可选）
        options: 转换选项字典，包含序号配置等
    
    返回:
        dict: 转换结果字典，包含以下键：
            - success: bool - 转换是否成功
            - main_content: str - 主要markdown内容（含YAML头部）
            - attachment_content: str or None - 附件markdown内容（公文模式）
            - metadata: dict - 提取的YAML元数据
            - error: str or None - 错误信息（如果失败）
    """
    logger.info(f"DOCX转MD路由: 优化类型={optimize_for_type}, 文件={docx_path}")
    
    # 准备公共参数
    common_params = {
        'docx_path': docx_path,
        'extract_image': extract_image,
        'extract_ocr': extract_ocr,
        'progress_callback': progress_callback,
        'cancel_event': cancel_event,
        'output_folder': output_folder,
        'original_file_path': original_file_path,
        'options': options
    }
    
    # 路由到对应转换器
    if optimize_for_type == "gongwen":
        logger.info("使用公文优化模式")
        from .gongwen.converter import convert_docx_to_md_gongwen
        return convert_docx_to_md_gongwen(**common_params)
    else:
        # 默认使用简化模式
        logger.info("使用简化模式")
        from .simple.converter import convert_docx_to_md_simple
        return convert_docx_to_md_simple(**common_params)


    
# 模块测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(level=logging.DEBUG)
    logger.info("DOCX转MD核心模块测试")
    
    # 测试转换功能
    result = convert_docx_to_md(
        docx_path="test.docx",
        config={}
    )
    
    if result['success']:
        print("转换成功!")
        print(f"主要内容长度: {len(result['main_content'])}")
        if result['attachment_content']:
            print(f"附件内容长度: {len(result['attachment_content'])}")
    else:
        print(f"转换失败: {result['error']}")
