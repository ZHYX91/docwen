"""
版式文件转Markdown策略

将PDF/OFD/XPS/CAJ等版式文件转换为Markdown格式。

使用 pymupdf4llm 进行转换，支持：
- 文本提取
- 图片提取
- OCR识别

依赖：
- .utils: 预处理函数
- converter.pdf2md: PDF转Markdown核心转换
"""

import os
import logging
import tempfile
from typing import Dict, Any, Callable, Optional

from docwen.services.result import ConversionResult
from docwen.services.strategies.base_strategy import BaseStrategy
from docwen.services.strategies import register_conversion, CATEGORY_LAYOUT
from docwen.i18n import t

from .utils import preprocess_layout_file

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_LAYOUT, 'md')
class LayoutToMarkdownStrategy(BaseStrategy):
    """
    使用pymupdf4llm将PDF转换为Markdown的策略
    
    新设计：所有输出（MD和图片）都放在一个文件夹内
    
    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 生成标准输出路径（含时间戳）
    3. 核心转换：PDF → Markdown（使用pymupdf4llm）
    4. 根据提取选项处理图片和OCR
    
    支持3种提取组合：
    1. ❌图片 ❌OCR：纯文本MD（放在文件夹内）
    2. ✅图片 ❌OCR：MD + 图片（同文件夹）
    3. ✅图片 ✅OCR：MD + 图片 + OCR（同文件夹）
    
    注意：内部总是提取文本，GUI不再显示"提取文字"选项
    
    输出结构：
    ```
    document_20251107_201500_fromPdf/
    ├── document_20251107_201500_fromPdf.md
    ├── image_1.png
    ├── image_2.png
    └── image_3.png
    ```
    
    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - CAJ：先转为PDF再处理（待实现）
    - OFD：先转为PDF再处理（待实现）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF到Markdown的转换（使用pymupdf4llm）
        
        参数:
            file_path: 输入的PDF文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
                - actual_format: 实际文件格式
                - extract_images: 是否提取图片（布尔值，默认False）
                - extract_ocr: 是否OCR识别（布尔值，默认False）
            progress_callback: 进度更新回调函数
            
        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")
        
        # 获取提取选项（只有2个）
        extract_images = options.get("extract_images", False)
        extract_ocr = options.get("extract_ocr", False)
        
        logger.info(
            f"PDF转Markdown - pymupdf4llm模式，"
            f"提取图片={extract_images}, OCR={extract_ocr}"
        )
        
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.preparing'))
            
            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 预处理：确保文件是PDF格式
                if actual_format and actual_format != 'pdf':
                    if progress_callback:
                        progress_callback(t('conversion.progress.converting_format_to_pdf', format=actual_format.upper()))
                    pdf_path, _ = preprocess_layout_file(
                        file_path, temp_dir, cancel_event, actual_format
                    )
                else:
                    pdf_path = file_path
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 确定输出目录
                from docwen.utils.workspace_manager import get_output_directory
                output_dir = get_output_directory(file_path)
                
                # 生成标准输出路径（不含扩展名，作为文件夹名）
                from docwen.utils.path_utils import generate_output_path
                
                folder_path_with_ext = generate_output_path(
                    file_path,
                    output_dir,
                    section="",
                    add_timestamp=True,
                    description='fromPdf',
                    file_type='md'
                )
                # 去掉.md扩展名，作为文件夹名
                basename_for_output = os.path.splitext(os.path.basename(folder_path_with_ext))[0]
                
                logger.info(f"输出文件夹基础名: {basename_for_output}")
                
                # 使用pymupdf4llm提取内容
                if progress_callback:
                    progress_callback(t('conversion.progress.extracting_pdf_content'))
                
                from docwen.converter.pdf2md import extract_pdf_with_pymupdf4llm
                
                result_data = extract_pdf_with_pymupdf4llm(
                    pdf_path,
                    extract_images,
                    extract_ocr,
                    output_dir,
                    basename_for_output,  # 传递标准化的文件夹名
                    cancel_event,
                    progress_callback  # 传递进度回调
                )
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 获取结果路径
                md_path = result_data['md_path']
                folder_path = result_data['folder_path']
                
                logger.info(f"Markdown文件已生成: {md_path}")
                logger.info(f"输出文件夹: {folder_path}")
                logger.info(
                    f"统计信息 - 图片: {result_data['image_count']}, "
                    f"OCR: {result_data['ocr_count']}"
                )
                
                return ConversionResult(
                    success=True,
                    output_path=md_path,  # 返回MD文件路径
                    message=t('conversion.messages.conversion_to_format_success', format='Markdown')
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
        except ImportError as e:
            error_msg = str(e)
            logger.error(f"缺少必要的库: {error_msg}")
            return ConversionResult(success=False, message=error_msg)
        except Exception as e:
            logger.error(f"执行 LayoutToMarkdownStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=f"转换失败: {str(e)}", error=e)


# 为了向后兼容，保留旧类名作为别名
LayoutToMarkdownPymupdf4llmStrategy = LayoutToMarkdownStrategy
