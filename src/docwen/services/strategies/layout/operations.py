"""
PDF操作策略

提供PDF文件的合并、拆分等操作功能。

支持的操作：
- 合并多个PDF/版式文件为一个PDF
- 拆分PDF文件为多个文件

依赖：
- .utils: 预处理函数
- PyMuPDF (fitz): PDF处理
"""

import os
import logging
import tempfile
import datetime
from typing import Dict, Any, Callable, Optional

from docwen.services.result import ConversionResult
from docwen.services.strategies.base_strategy import BaseStrategy
from docwen.services.strategies import register_action
from docwen.i18n import t

from .utils import preprocess_layout_file

logger = logging.getLogger(__name__)


@register_action("merge_pdfs")
class MergePdfsStrategy(BaseStrategy):
    """
    合并多个PDF/版式文件为一个PDF的策略
    
    触发条件：批量模式下，版式文件数量 > 1
    
    转换流程：
    1. 预处理：将所有非PDF格式文件转换为PDF（XPS/OFD/CAJ）
    2. 按文件列表顺序合并PDF
    3. 输出到第一个文件所在目录
    
    命名规则：
    - 输出文件名：合并文档_{时间戳}.pdf
    
    支持的输入格式：
    - PDF：直接合并
    - XPS：先转为PDF再合并
    - OFD：先转为PDF再合并（待实现）
    - CAJ：先转为PDF再合并（待实现）
    - 混合：支持以上格式混合合并
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF合并操作
        
        参数:
            file_path: 第一个文件的路径（用于确定输出目录）
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
                - file_list: 要合并的文件列表（必需）
                - actual_formats: 各文件的实际格式列表（可选）
            progress_callback: 进度更新回调函数
            
        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        file_list = options.get("file_list", [])
        actual_formats = options.get("actual_formats", [])
        
        if not file_list:
            return ConversionResult(
                success=False,
                message=t('conversion.messages.no_files_to_merge')
            )
        
        if len(file_list) < 2:
            return ConversionResult(
                success=False,
                message=t('conversion.messages.need_at_least_two_files')
            )
        
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.preparing_merge', count=len(file_list)))
            
            logger.info(f"开始合并{len(file_list)}个PDF/版式文件")
            
            # 使用临时目录处理中间文件
            with tempfile.TemporaryDirectory() as temp_dir:
                pdf_paths = []
                
                # 步骤1：预处理所有文件，确保都是PDF格式
                for i, input_file in enumerate(file_list):
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                    
                    file_name = os.path.basename(input_file)
                    actual_format = actual_formats[i] if i < len(actual_formats) else None
                    
                    if progress_callback:
                        progress_callback(t('conversion.progress.processing_file_progress', current=i+1, total=len(file_list), filename=file_name))
                    
                    # 自动检测格式（如果未提供）
                    if not actual_format:
                        from docwen.utils.file_type_utils import detect_actual_file_format
                        actual_format = detect_actual_file_format(input_file)
                    
                    # 如果是PDF，直接使用；否则先转换
                    if actual_format == 'pdf':
                        pdf_paths.append(input_file)
                        logger.debug(f"文件{i+1}已是PDF: {file_name}")
                    else:
                        logger.info(f"文件{i+1}需要转换: {file_name} ({actual_format} -> PDF)")
                        pdf_path, _ = preprocess_layout_file(
                            input_file, temp_dir, cancel_event, actual_format
                        )
                        pdf_paths.append(pdf_path)
                
                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                
                # 步骤2：合并所有PDF
                if progress_callback:
                    progress_callback(t('conversion.progress.merging_pdf_files'))
                
                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    return ConversionResult(
                        success=False,
                        message=t('conversion.messages.missing_pymupdf')
                    )
                
                # 创建输出PDF
                merged_pdf = fitz.open()
                
                try:
                    for i, pdf_path in enumerate(pdf_paths):
                        if cancel_event and cancel_event.is_set():
                            merged_pdf.close()
                            return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                        
                        if progress_callback:
                            progress_callback(t('conversion.progress.merging_file_progress', current=i+1, total=len(pdf_paths)))
                        
                        # 打开并插入PDF
                        with fitz.open(pdf_path) as pdf:
                            merged_pdf.insert_pdf(pdf)
                        
                        logger.debug(f"已合并文件{i+1}: {os.path.basename(pdf_path)}")
                    
                    # 生成输出路径（使用选中文件所在目录）
                    from docwen.utils.workspace_manager import get_output_directory
                    selected_file = options.get("selected_file", file_path)
                    output_dir = get_output_directory(selected_file)
                    logger.info(f"输出目录基于: {os.path.basename(selected_file)}")
                    
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_path = os.path.join(output_dir, f"{t('conversion.filenames.merged_pdf')}_{timestamp}.pdf")
                    
                    # 保存合并后的PDF
                    merged_pdf.save(output_path)
                    logger.info(f"合并PDF成功，共{len(merged_pdf)}页: {output_path}")
                    
                    return ConversionResult(
                        success=True,
                        output_path=output_path,
                        message=t('conversion.messages.merge_pdf_success', file_count=len(file_list), page_count=len(merged_pdf))
                    )
                
                finally:
                    merged_pdf.close()
        
        except InterruptedError:
            return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
        except Exception as e:
            logger.error(f"执行 MergePdfsStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=t('conversion.messages.conversion_failed_with_error', error=str(e)), error=e)


@register_action("split_pdf")
class SplitPdfStrategy(BaseStrategy):
    """
    拆分PDF/版式文件为两个PDF的策略
    
    触发条件：单个版式文件 + 页码输入合法
    
    转换流程：
    1. 预处理：如果非PDF格式，先转换为PDF（XPS/OFD/CAJ）
    2. 根据用户输入的页码拆分为两个PDF文件：
       - 第1个文件：用户输入的页码
       - 第2个文件：剩余页码（如果有）
    3. 输出到原文件所在目录
    
    命名规则：
    - 第1个文件：原文件名_拆分1_{时间戳}_split.pdf
    - 第2个文件：原文件名_拆分2_{时间戳}_split.pdf（如果有剩余页）
    
    支持的输入格式：
    - PDF：直接拆分
    - XPS：先转为PDF再拆分
    - OFD：先转为PDF再拆分（待实现）
    - CAJ：先转为PDF再拆分（待实现）
    """
    
    def execute(
        self,
        file_path: str,
        options: Optional[Dict[str, Any]] = None,
        progress_callback: Optional[Callable[[str], None]] = None
    ) -> ConversionResult:
        """
        执行PDF拆分操作
        
        参数:
            file_path: 输入的PDF/版式文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
                - pages: 要提取的页码列表（必需，已排序去重）
                - total_pages: PDF总页数（可选，用于验证）
                - actual_format: 实际文件格式（可选）
            progress_callback: 进度更新回调函数
            
        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        pages = options.get("pages", [])
        total_pages = options.get("total_pages", 0)
        actual_format = options.get("actual_format")
        
        if not pages:
            return ConversionResult(
                success=False,
                message=t('conversion.messages.no_pages_to_split')
            )
        
        try:
            if progress_callback:
                progress_callback(t('conversion.progress.preparing_split_pdf'))
            
            logger.info(f"开始拆分PDF，页码: {pages}")
            
            # 使用临时目录处理中间文件
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理，确保文件是PDF格式
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
                
                # 步骤2：打开PDF并验证页码
                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    return ConversionResult(
                        success=False,
                        message=t('conversion.messages.missing_pymupdf')
                    )
                
                with fitz.open(pdf_path) as pdf:
                    actual_total_pages = len(pdf)
                    logger.info(f"PDF总页数: {actual_total_pages}")
                    
                    # 验证并过滤页码
                    valid_pages = [p for p in pages if 1 <= p <= actual_total_pages]
                    if not valid_pages:
                        return ConversionResult(
                            success=False,
                            message=t('conversion.messages.all_pages_invalid', total=actual_total_pages)
                        )
                    
                    # 计算剩余页码
                    all_pages = set(range(1, actual_total_pages + 1))
                    remaining_pages = sorted(list(all_pages - set(valid_pages)))
                    
                    logger.info(f"第1个文件页码: {valid_pages}")
                    logger.info(f"第2个文件页码: {remaining_pages if remaining_pages else '无'}")
                    
                    # 验证：确保有剩余页码（不能拆分全部页码）
                    if not remaining_pages:
                        return ConversionResult(
                            success=False,
                            message=t('conversion.messages.split_failed_all_pages')
                        )
                    
                    # 检查取消
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                    
                    # 步骤3：生成输出路径
                    from docwen.utils.path_utils import generate_output_path
                    from docwen.utils.workspace_manager import get_output_directory
                    output_dir = get_output_directory(file_path)
                    
                    # 第1个文件路径
                    output_path1 = generate_output_path(
                        file_path,
                        output_dir=output_dir,
                        section=t('conversion.filenames.split_part1'),
                        add_timestamp=True,
                        description="split",
                        file_type="pdf"
                    )
                    
                    # 步骤4：创建第1个PDF（用户选择的页码）
                    if progress_callback:
                        progress_callback(t('conversion.progress.creating_pdf_part', part=1))
                    
                    pdf1 = fitz.open()
                    try:
                        for page_num in valid_pages:
                            pdf1.insert_pdf(pdf, from_page=page_num-1, to_page=page_num-1)
                        
                        pdf1.save(output_path1)
                        logger.info(f"第1个PDF已保存: {output_path1}，共{len(pdf1)}页")
                    finally:
                        pdf1.close()
                    
                    # 检查取消
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
                    
                    # 步骤5：创建第2个PDF（剩余页码，如果有）
                    output_path2 = None
                    if remaining_pages:
                        if progress_callback:
                            progress_callback(t('conversion.progress.creating_pdf_part', part=2))
                        
                        # 第2个文件路径
                        output_path2 = generate_output_path(
                            file_path,
                            output_dir=output_dir,
                            section=t('conversion.filenames.split_part2'),
                            add_timestamp=True,
                            description="split",
                            file_type="pdf"
                        )
                        
                        pdf2 = fitz.open()
                        try:
                            for page_num in remaining_pages:
                                pdf2.insert_pdf(pdf, from_page=page_num-1, to_page=page_num-1)
                            
                            pdf2.save(output_path2)
                            logger.info(f"第2个PDF已保存: {output_path2}，共{len(pdf2)}页")
                        finally:
                            pdf2.close()
                    else:
                        logger.info("无剩余页码，不生成第2个PDF")
                
                # 构建结果消息
                if output_path2:
                    message = t('conversion.messages.split_pdf_success_two', pages1=len(valid_pages), pages2=len(remaining_pages))
                else:
                    message = t('conversion.messages.split_pdf_success_one', pages=len(valid_pages))
                
                return ConversionResult(
                    success=True,
                    output_path=output_path1,  # 返回第1个文件路径
                    message=message
                )
        
        except InterruptedError:
            return ConversionResult(success=False, message=t('conversion.messages.operation_cancelled'))
        except Exception as e:
            logger.error(f"执行 SplitPdfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(success=False, message=t('conversion.messages.conversion_failed_with_error', error=str(e)), error=e)
