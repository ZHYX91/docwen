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

import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.translation import t
from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_DEPENDENCY_MISSING,
    ERROR_CODE_INVALID_INPUT,
    ERROR_CODE_NOT_IMPLEMENTED,
    ERROR_CODE_OPERATION_CANCELLED,
)
from docwen.services.result import ConversionResult
from docwen.services.strategies import register_action
from docwen.services.strategies.base_strategy import BaseStrategy

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
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
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
            msg = t("conversion.messages.no_files_to_merge")
            return ConversionResult(success=False, message=msg, error_code=ERROR_CODE_INVALID_INPUT, details=msg)

        if len(file_list) < 2:
            msg = t("conversion.messages.need_at_least_two_files")
            return ConversionResult(success=False, message=msg, error_code=ERROR_CODE_INVALID_INPUT, details=msg)

        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing_merge", count=len(file_list)))

            logger.info(f"开始合并{len(file_list)}个PDF/版式文件")

            # 使用临时目录处理中间文件
            with tempfile.TemporaryDirectory() as temp_dir:
                pdf_paths = []

                # 步骤1：预处理所有文件，确保都是PDF格式
                for i, input_file in enumerate(file_list):
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(
                            success=False,
                            message=t("conversion.messages.operation_cancelled"),
                            error_code=ERROR_CODE_OPERATION_CANCELLED,
                        )

                    file_name = Path(input_file).name
                    actual_format = actual_formats[i] if i < len(actual_formats) else None

                    if progress_callback:
                        progress_callback(
                            t(
                                "conversion.progress.processing_file_progress",
                                current=i + 1,
                                total=len(file_list),
                                filename=file_name,
                            )
                        )

                    # 自动检测格式（如果未提供）
                    if not actual_format:
                        from docwen.utils.file_type_utils import detect_actual_file_format

                        actual_format = detect_actual_file_format(input_file)

                    # 如果是PDF，直接使用；否则先转换
                    if actual_format != "pdf":
                        logger.info(f"文件{i + 1}需要转换: {file_name} ({actual_format} -> PDF)")

                    if actual_format == "caj":
                        return ConversionResult(
                            success=False,
                            message=t(
                                "conversion.messages.format_to_pdf_not_implemented",
                                default="{format} 转 PDF 未实现",
                                format="CAJ",
                            ),
                            error_code=ERROR_CODE_NOT_IMPLEMENTED,
                        )

                    per_file_temp_dir = str(Path(temp_dir) / f"file_{i + 1}")
                    Path(per_file_temp_dir).mkdir(parents=True, exist_ok=True)

                    preprocess_result = preprocess_layout_file(
                        input_file, per_file_temp_dir, cancel_event, actual_format
                    )
                    pdf_paths.append(preprocess_result.processed_file)
                    logger.debug(f"文件{i + 1}预处理完成: {Path(preprocess_result.processed_file).name}")

                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 步骤2：合并所有PDF
                if progress_callback:
                    progress_callback(t("conversion.progress.merging_pdf_files"))

                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    msg = t("conversion.messages.missing_pymupdf")
                    return ConversionResult(
                        success=False, message=msg, error_code=ERROR_CODE_DEPENDENCY_MISSING, details=msg
                    )

                # 创建输出PDF
                merged_pdf = fitz.open()

                try:
                    for i, pdf_path in enumerate(pdf_paths):
                        if cancel_event and cancel_event.is_set():
                            merged_pdf.close()
                            return ConversionResult(
                                success=False,
                                message=t("conversion.messages.operation_cancelled"),
                                error_code=ERROR_CODE_OPERATION_CANCELLED,
                            )

                        if progress_callback:
                            progress_callback(
                                t("conversion.progress.merging_file_progress", current=i + 1, total=len(pdf_paths))
                            )

                        # 打开并插入PDF
                        with fitz.open(pdf_path) as pdf:
                            merged_pdf.insert_pdf(pdf)

                        logger.debug(f"已合并文件{i + 1}: {Path(pdf_path).name}")

                    # 生成输出路径（使用选中文件所在目录）
                    from docwen.utils.workspace_manager import get_output_directory

                    selected_file = options.get("selected_file", file_path)
                    output_dir = get_output_directory(selected_file)
                    logger.info(f"输出目录基于: {Path(selected_file).name}")

                    from docwen.utils.path_utils import generate_named_output_path

                    output_path = generate_named_output_path(
                        output_dir=output_dir,
                        base_name=t("conversion.filenames.merged_pdf"),
                        file_type="pdf",
                        add_timestamp=True,
                    )

                    from docwen.utils.workspace_manager import finalize_output

                    temp_output_path = str(Path(temp_dir) / Path(output_path).name)
                    merged_pdf.save(temp_output_path)
                    saved_path, _ = finalize_output(
                        temp_output_path,
                        output_path,
                        original_input_file=selected_file,
                    )
                    if not saved_path:
                        return ConversionResult(
                            success=False,
                            message=t(
                                "conversion.messages.conversion_failed_with_error",
                                default="转换失败: {error}",
                                error="保存输出文件失败",
                            ),
                            error_code=ERROR_CODE_CONVERSION_FAILED,
                        )
                    output_path = saved_path
                    logger.info(f"合并PDF成功，共{len(merged_pdf)}页: {output_path}")

                    return ConversionResult(
                        success=True,
                        output_path=output_path,
                        message=t(
                            "conversion.messages.merge_pdf_success",
                            file_count=len(file_list),
                            page_count=len(merged_pdf),
                        ),
                    )

                finally:
                    merged_pdf.close()

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except Exception as e:
            logger.error(f"执行 MergePdfsStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t(
                    "conversion.messages.conversion_failed_with_error",
                    default="转换失败: {error}",
                    error=str(e),
                ),
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )


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
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
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
        split_mode = options.get("split_mode", "custom")
        pages = options.get("pages", [])
        actual_format = options.get("actual_format")

        if split_mode == "custom" and not pages:
            msg = t("conversion.messages.no_pages_to_split")
            return ConversionResult(success=False, message=msg, error_code=ERROR_CODE_INVALID_INPUT, details=msg)

        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing_split_pdf"))

            logger.info(f"开始拆分PDF，模式: {split_mode}, 页码: {pages}")

            # 使用临时目录处理中间文件
            with tempfile.TemporaryDirectory() as temp_dir:
                # 步骤1：预处理，确保文件是PDF格式
                if actual_format and actual_format != "pdf" and progress_callback:
                    progress_callback(t("conversion.progress.converting_format_to_pdf", format=actual_format.upper()))

                if actual_format == "caj":
                    return ConversionResult(
                        success=False,
                        message=t(
                            "conversion.messages.format_to_pdf_not_implemented",
                            default="{format} 转 PDF 未实现",
                            format="CAJ",
                        ),
                        error_code=ERROR_CODE_NOT_IMPLEMENTED,
                    )

                preprocess_result = preprocess_layout_file(file_path, temp_dir, cancel_event, actual_format)
                pdf_path = preprocess_result.processed_file

                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 步骤2：打开PDF并验证页码
                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    msg = t("conversion.messages.missing_pymupdf")
                    return ConversionResult(
                        success=False, message=msg, error_code=ERROR_CODE_DEPENDENCY_MISSING, details=msg
                    )

                with fitz.open(pdf_path) as pdf:
                    actual_total_pages = len(pdf)
                    logger.info(f"PDF总页数: {actual_total_pages}")

                    if actual_total_pages <= 1:
                        msg = t(
                            "conversion.messages.split_pdf_single_page_not_needed",
                            default="当前文件只有1页，无需拆分",
                        )
                        return ConversionResult(
                            success=False, message=msg, error_code=ERROR_CODE_INVALID_INPUT, details=msg
                        )

                    # 检查取消
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(
                            success=False,
                            message=t("conversion.messages.operation_cancelled"),
                            error_code=ERROR_CODE_OPERATION_CANCELLED,
                        )

                    from docwen.utils.workspace_manager import get_output_directory
                    from docwen.utils.workspace_manager import finalize_output
                    from docwen.utils.path_utils import generate_named_output_path, generate_timestamp

                    output_dir = get_output_directory(file_path)
                    base_name = Path(file_path).stem
                    shared_timestamp = generate_timestamp()

                    def _save_output(doc: fitz.Document, output_path: str) -> str | None:
                        temp_output_path = str(Path(temp_dir) / Path(output_path).name)
                        doc.save(temp_output_path)
                        saved_path, _ = finalize_output(
                            temp_output_path,
                            output_path,
                            original_input_file=file_path,
                        )
                        return saved_path

                    if split_mode == "every_page":
                        output_paths: list[str] = []
                        for page_num in range(1, actual_total_pages + 1):
                            if cancel_event and cancel_event.is_set():
                                return ConversionResult(
                                    success=False,
                                    message=t("conversion.messages.operation_cancelled"),
                                    error_code=ERROR_CODE_OPERATION_CANCELLED,
                                )

                            section = t("conversion.filenames.page_n", n=page_num)
                            output_path = generate_named_output_path(
                                output_dir=output_dir,
                                base_name=base_name,
                                file_type="pdf",
                                section=section,
                                add_timestamp=True,
                                description="split",
                                timestamp_override=shared_timestamp,
                            )

                            pdf_single = fitz.open()
                            try:
                                pdf_single.insert_pdf(pdf, from_page=page_num - 1, to_page=page_num - 1)
                                saved_path = _save_output(pdf_single, output_path)
                                if not saved_path:
                                    return ConversionResult(
                                        success=False,
                                        message=t(
                                            "conversion.messages.conversion_failed_with_error",
                                            default="转换失败: {error}",
                                            error="保存拆分文件失败",
                                        ),
                                        error_code=ERROR_CODE_CONVERSION_FAILED,
                                    )
                                output_paths.append(saved_path)
                            finally:
                                pdf_single.close()

                        message = t(
                            "conversion.messages.split_pdf_success_many",
                            default="已拆分: 共生成 {count} 个文件",
                            count=len(output_paths),
                        )
                        return ConversionResult(success=True, output_path=output_paths[0], message=message)

                    if split_mode == "odd_even":
                        odd_pages = [p for p in range(1, actual_total_pages + 1) if p % 2 == 1]
                        even_pages = [p for p in range(1, actual_total_pages + 1) if p % 2 == 0]

                        output_path_odd = generate_named_output_path(
                            output_dir=output_dir,
                            base_name=base_name,
                            file_type="pdf",
                            section=t("conversion.filenames.odd_pages"),
                            add_timestamp=True,
                            description="split",
                            timestamp_override=shared_timestamp,
                        )
                        pdf_odd = fitz.open()
                        try:
                            for page_num in odd_pages:
                                pdf_odd.insert_pdf(pdf, from_page=page_num - 1, to_page=page_num - 1)
                            saved_odd = _save_output(pdf_odd, output_path_odd)
                            if not saved_odd:
                                return ConversionResult(
                                    success=False,
                                    message=t(
                                        "conversion.messages.conversion_failed_with_error",
                                        default="转换失败: {error}",
                                        error="保存拆分文件失败",
                                    ),
                                    error_code=ERROR_CODE_CONVERSION_FAILED,
                                )
                            output_path_odd = saved_odd
                        finally:
                            pdf_odd.close()

                        output_path_even = generate_named_output_path(
                            output_dir=output_dir,
                            base_name=base_name,
                            file_type="pdf",
                            section=t("conversion.filenames.even_pages"),
                            add_timestamp=True,
                            description="split",
                            timestamp_override=shared_timestamp,
                        )
                        pdf_even = fitz.open()
                        try:
                            for page_num in even_pages:
                                pdf_even.insert_pdf(pdf, from_page=page_num - 1, to_page=page_num - 1)
                            saved_even = _save_output(pdf_even, output_path_even)
                            if not saved_even:
                                return ConversionResult(
                                    success=False,
                                    message=t(
                                        "conversion.messages.conversion_failed_with_error",
                                        default="转换失败: {error}",
                                        error="保存拆分文件失败",
                                    ),
                                    error_code=ERROR_CODE_CONVERSION_FAILED,
                                )
                            output_path_even = saved_even
                        finally:
                            pdf_even.close()

                        message = t(
                            "conversion.messages.split_pdf_success_two",
                            pages1=len(odd_pages),
                            pages2=len(even_pages),
                        )
                        return ConversionResult(success=True, output_path=output_path_odd, message=message)

                    # 验证并过滤页码
                    valid_pages = [p for p in pages if 1 <= p <= actual_total_pages]
                    if not valid_pages:
                        return ConversionResult(
                            success=False,
                            message=t("conversion.messages.all_pages_invalid", total=actual_total_pages),
                            error_code=ERROR_CODE_INVALID_INPUT,
                        )

                    all_pages = set(range(1, actual_total_pages + 1))
                    remaining_pages = sorted(all_pages - set(valid_pages))

                    logger.info(f"第1个文件页码: {valid_pages}")
                    logger.info(f"第2个文件页码: {remaining_pages if remaining_pages else '无'}")

                    if not remaining_pages:
                        msg = t("conversion.messages.split_failed_all_pages")
                        return ConversionResult(
                            success=False, message=msg, error_code=ERROR_CODE_INVALID_INPUT, details=msg
                        )

                    output_path1 = generate_named_output_path(
                        output_dir=output_dir,
                        base_name=base_name,
                        file_type="pdf",
                        section=t("conversion.filenames.split_part1"),
                        add_timestamp=True,
                        description="split",
                        timestamp_override=shared_timestamp,
                    )

                    # 步骤4：创建第1个PDF（用户选择的页码）
                    if progress_callback:
                        progress_callback(t("conversion.progress.creating_pdf_part", part=1))

                    pdf1 = fitz.open()
                    try:
                        for page_num in valid_pages:
                            pdf1.insert_pdf(pdf, from_page=page_num - 1, to_page=page_num - 1)
                        saved_path1 = _save_output(pdf1, output_path1)
                        if not saved_path1:
                            return ConversionResult(
                                success=False,
                                message=t(
                                    "conversion.messages.conversion_failed_with_error",
                                    default="转换失败: {error}",
                                    error="保存拆分文件失败",
                                ),
                                error_code=ERROR_CODE_CONVERSION_FAILED,
                            )
                        output_path1 = saved_path1
                        logger.info(f"第1个PDF已保存: {output_path1}，共{len(pdf1)}页")
                    finally:
                        pdf1.close()

                    # 检查取消
                    if cancel_event and cancel_event.is_set():
                        return ConversionResult(
                            success=False,
                            message=t("conversion.messages.operation_cancelled"),
                            error_code=ERROR_CODE_OPERATION_CANCELLED,
                        )

                    # 步骤5：创建第2个PDF（剩余页码，如果有）
                    output_path2 = None
                    if remaining_pages:
                        if progress_callback:
                            progress_callback(t("conversion.progress.creating_pdf_part", part=2))

                        output_path2 = generate_named_output_path(
                            output_dir=output_dir,
                            base_name=base_name,
                            file_type="pdf",
                            section=t("conversion.filenames.split_part2"),
                            add_timestamp=True,
                            description="split",
                            timestamp_override=shared_timestamp,
                        )

                        pdf2 = fitz.open()
                        try:
                            for page_num in remaining_pages:
                                pdf2.insert_pdf(pdf, from_page=page_num - 1, to_page=page_num - 1)
                            saved_path2 = _save_output(pdf2, output_path2)
                            if not saved_path2:
                                return ConversionResult(
                                    success=False,
                                    message=t(
                                        "conversion.messages.conversion_failed_with_error",
                                        default="转换失败: {error}",
                                        error="保存拆分文件失败",
                                    ),
                                    error_code=ERROR_CODE_CONVERSION_FAILED,
                                )
                            output_path2 = saved_path2
                            logger.info(f"第2个PDF已保存: {output_path2}，共{len(pdf2)}页")
                        finally:
                            pdf2.close()
                    else:
                        logger.info("无剩余页码，不生成第2个PDF")

                # 构建结果消息
                if output_path2:
                    message = t(
                        "conversion.messages.split_pdf_success_two",
                        pages1=len(valid_pages),
                        pages2=len(remaining_pages),
                    )
                else:
                    message = t("conversion.messages.split_pdf_success_one", pages=len(valid_pages))

                return ConversionResult(
                    success=True,
                    output_path=output_path1,  # 返回第1个文件路径
                    message=message,
                )

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except Exception as e:
            logger.error(f"执行 SplitPdfStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t(
                    "conversion.messages.conversion_failed_with_error",
                    default="转换失败: {error}",
                    error=str(e),
                ),
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )
