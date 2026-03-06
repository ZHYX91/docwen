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
    ERROR_CODE_OPERATION_CANCELLED,
)
from docwen.services.result import ConversionResult
from docwen.services.strategies import CATEGORY_LAYOUT, register_conversion
from docwen.services.strategies.base_strategy import BaseStrategy
from docwen.utils.validation_utils import validate_ocr_requires_images

from .utils import preprocess_layout_file, should_keep_intermediates

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_LAYOUT, "md")
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
    document_20251107_201500_from{ActualFormat}/
    ├── document_20251107_201500_from{ActualFormat}.md
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
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行PDF到Markdown的转换（使用pymupdf4llm）

        参数:
            file_path: 输入的PDF文件路径
            options: 转换选项字典，包含：
                - cancel_event: 取消事件
                - actual_format: 实际文件格式
                - extract_image: 是否提取图片（布尔值，默认False）
                - extract_ocr: 是否OCR识别（布尔值，默认False）
            progress_callback: 进度更新回调函数

        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        cancel_event = options.get("cancel_event")
        actual_format = options.get("actual_format")

        if not actual_format:
            from docwen.utils.file_type_utils import detect_actual_file_format

            actual_format = detect_actual_file_format(file_path)

        extract_image = bool(options.get("extract_image", False))
        extract_ocr = options.get("extract_ocr", False)

        from docwen.config.config_manager import config_manager

        options.setdefault(
            "to_md_image_extraction_mode",
            config_manager.get_layout_to_md_image_extraction_mode(),
        )
        options.setdefault(
            "to_md_ocr_placement_mode",
            config_manager.get_layout_to_md_ocr_placement_mode(),
        )

        ok, reason = validate_ocr_requires_images(extract_image, extract_ocr)
        if not ok:
            return ConversionResult(
                success=False,
                message=reason,
                error_code=ERROR_CODE_INVALID_INPUT,
                details=reason,
            )

        logger.info(f"PDF转Markdown - pymupdf4llm模式，提取图片={extract_image}, OCR={extract_ocr}")

        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing"))

            optimize_for_type = options.get("optimize_for_type")
            if optimize_for_type == "invoice_cn" and actual_format in {"ofd", "pdf"}:
                try:
                    from docwen.utils.path_utils import generate_output_path
                    from docwen.utils.workspace_manager import get_output_directory

                    output_dir = get_output_directory(file_path)
                    folder_path_with_ext = generate_output_path(
                        file_path,
                        output_dir,
                        section="",
                        add_timestamp=True,
                        description=f"from{actual_format.capitalize()}",
                        file_type="md",
                    )
                    basename_for_output = Path(folder_path_with_ext).stem

                    if progress_callback:
                        progress_callback(t("conversion.progress.extracting_pdf_content"))

                    from docwen.converter.layout2md.invoice_parser import convert_invoice_layout_to_md

                    result_data = convert_invoice_layout_to_md(
                        file_path=file_path,
                        actual_format=actual_format,
                        output_dir=output_dir,
                        basename_for_output=basename_for_output,
                        original_file_stem=Path(file_path).stem,
                        cancel_event=cancel_event,
                        progress_callback=progress_callback,
                    )

                    md_path = result_data["md_path"]
                    folder_path = result_data["folder_path"]
                    page_count = int(result_data.get("page_count", 1) or 1)
                    message = (
                        t(
                            "conversion.messages.invoice_md_multi_page_success",
                            default="转换成功，共生成 {count} 个发票 Markdown 文件",
                            count=page_count,
                        )
                        if page_count > 1
                        else t("conversion.messages.conversion_to_format_success", format="Markdown")
                    )

                    try:
                        from docwen.config.config_manager import config_manager

                        if config_manager.get_save_manifest():
                            from docwen.utils.workspace_manager import build_manifest, write_manifest_json

                            manifest = build_manifest(
                                file_path=file_path,
                                actual_format=actual_format,
                                preprocess_chain=[],
                                saved_intermediate_items=[],
                                options=options,
                                success=True,
                                message=message,
                                output_path=md_path,
                                mask_input=config_manager.get_mask_manifest_input_path(),
                            )
                            write_manifest_json(folder_path, manifest)
                    except Exception:
                        pass

                    return ConversionResult(
                        success=True,
                        output_path=md_path,
                        message=message,
                    )
                except Exception as e:
                    logger.warning(f"发票优化解析失败，回退到默认转换: {e}", exc_info=True)

            # 使用临时目录（ignore_cleanup_errors 避免 Windows 上 PyMuPDF 文件句柄未释放导致清理失败）
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp_dir:
                if actual_format and actual_format != "pdf" and progress_callback:
                    progress_callback(t("conversion.progress.converting_format_to_pdf", format=actual_format.upper()))

                preprocess_result = preprocess_layout_file(file_path, temp_dir, cancel_event, actual_format)
                pdf_path = preprocess_result.processed_file
                actual_format = preprocess_result.actual_format
                options["actual_format"] = actual_format

                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

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
                    description=f"from{actual_format.capitalize()}",
                    file_type="md",
                )
                # 去掉.md扩展名，作为文件夹名
                basename_for_output = Path(folder_path_with_ext).stem

                logger.info(f"输出文件夹基础名: {basename_for_output}")

                if optimize_for_type == "invoice_cn":
                    try:
                        from docwen.converter.layout2md.invoice_parser import convert_invoice_layout_to_md

                        result_data = convert_invoice_layout_to_md(
                            file_path=pdf_path,
                            actual_format="pdf",
                            output_dir=output_dir,
                            basename_for_output=basename_for_output,
                            original_file_stem=Path(file_path).stem,
                            cancel_event=cancel_event,
                            progress_callback=progress_callback,
                        )

                        md_path = result_data["md_path"]
                        folder_path = result_data["folder_path"]
                        page_count = int(result_data.get("page_count", 1) or 1)
                        message = (
                            t(
                                "conversion.messages.invoice_md_multi_page_success",
                                default="转换成功，共生成 {count} 个发票 Markdown 文件",
                                count=page_count,
                            )
                            if page_count > 1
                            else t("conversion.messages.conversion_to_format_success", format="Markdown")
                        )

                        try:
                            from docwen.config.config_manager import config_manager

                            if config_manager.get_save_manifest():
                                from docwen.utils.workspace_manager import build_manifest, write_manifest_json

                                manifest = build_manifest(
                                    file_path=file_path,
                                    actual_format=actual_format,
                                    preprocess_chain=preprocess_result.preprocess_chain,
                                    saved_intermediate_items=[],
                                    options=options,
                                    success=True,
                                    message=message,
                                    output_path=md_path,
                                    mask_input=config_manager.get_mask_manifest_input_path(),
                                )
                                write_manifest_json(folder_path, manifest)
                        except Exception:
                            pass

                        return ConversionResult(
                            success=True,
                            output_path=md_path,
                            message=message,
                        )
                    except Exception as e:
                        logger.warning(f"发票优化解析失败（PDF模式），回退到默认转换: {e}", exc_info=True)

                # 使用pymupdf4llm提取内容
                if progress_callback:
                    progress_callback(t("conversion.progress.extracting_pdf_content"))

                from docwen.converter.pdf2md import extract_pdf_with_pymupdf4llm

                result_data = extract_pdf_with_pymupdf4llm(
                    pdf_path,
                    extract_image,
                    extract_ocr,
                    output_dir,
                    basename_for_output,  # 传递标准化的文件夹名
                    cancel_event,
                    progress_callback,  # 传递进度回调
                    ocr_placement_mode=options.get("to_md_ocr_placement_mode", "image_md"),
                    extraction_mode=options.get("to_md_image_extraction_mode", "file"),
                )

                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 获取结果路径
                md_path = result_data["md_path"]
                folder_path = result_data["folder_path"]

                logger.info(f"Markdown文件已生成: {md_path}")
                logger.info(f"输出文件夹: {folder_path}")
                logger.info(f"统计信息 - 图片: {result_data['image_count']}, OCR: {result_data['ocr_count']}")

                saved_items = []
                if should_keep_intermediates():
                    from docwen.utils.workspace_manager import save_intermediate_items

                    saved_items = save_intermediate_items(preprocess_result.intermediates, output_dir, move=True)
                    for _, saved_path in saved_items:
                        logger.info(f"保留中间文件: {Path(saved_path).name}")

                try:
                    from docwen.config.config_manager import config_manager

                    if config_manager.get_save_manifest():
                        from docwen.utils.workspace_manager import build_manifest, write_manifest_json

                        manifest = build_manifest(
                            file_path=file_path,
                            actual_format=actual_format,
                            preprocess_chain=preprocess_result.preprocess_chain,
                            saved_intermediate_items=saved_items,
                            options=options,
                            success=True,
                            message=t("conversion.messages.conversion_to_format_success", format="Markdown"),
                            output_path=md_path,
                            mask_input=config_manager.get_mask_manifest_input_path(),
                        )
                        write_manifest_json(folder_path, manifest)
                except Exception:
                    pass

                return ConversionResult(
                    success=True,
                    output_path=md_path,  # 返回MD文件路径
                    message=t("conversion.messages.conversion_to_format_success", format="Markdown"),
                )

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except ImportError as e:
            error_msg = str(e)
            logger.error(f"缺少必要的库: {error_msg}")
            try:
                from docwen.config.config_manager import config_manager

                if config_manager.get_save_manifest():
                    from datetime import datetime

                    from docwen.utils.workspace_manager import build_manifest, write_manifest_json

                    manifest = build_manifest(
                        file_path=file_path,
                        actual_format=locals().get("actual_format"),
                        preprocess_chain=getattr(locals().get("preprocess_result"), "preprocess_chain", None),
                        saved_intermediate_items=[],
                        options=(options or {}),
                        success=False,
                        message=error_msg,
                        output_path=None,
                        mask_input=config_manager.get_mask_manifest_input_path(),
                    )
                    write_manifest_json(
                        locals().get("output_dir") or str(Path(file_path).parent),
                        manifest,
                        filename=f"manifest_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    )
            except Exception:
                pass
            return ConversionResult(
                success=False,
                message=error_msg,
                error_code=ERROR_CODE_DEPENDENCY_MISSING,
                details=error_msg or None,
                error=e,
            )
        except Exception as e:
            logger.error(f"执行 LayoutToMarkdownStrategy 时出错: {e}", exc_info=True)
            try:
                from docwen.config.config_manager import config_manager
                from docwen.utils.workspace_manager import build_manifest, write_manifest_json

                failure_output_dir = locals().get("output_dir") or str(Path(file_path).parent)

                saved_failure_items = []
                if should_keep_intermediates() and "preprocess_result" in locals():
                    try:
                        from docwen.utils.workspace_manager import save_intermediate_items

                        saved_failure_items = save_intermediate_items(
                            preprocess_result.intermediates, failure_output_dir, move=True
                        )
                    except Exception:
                        saved_failure_items = []

                if config_manager.get_save_manifest():
                    from datetime import datetime

                    manifest = build_manifest(
                        file_path=file_path,
                        actual_format=locals().get("actual_format"),
                        preprocess_chain=getattr(locals().get("preprocess_result"), "preprocess_chain", None),
                        saved_intermediate_items=saved_failure_items,
                        options=(options or {}),
                        success=False,
                        message=str(e),
                        output_path=None,
                        mask_input=config_manager.get_mask_manifest_input_path(),
                    )
                    write_manifest_json(
                        failure_output_dir,
                        manifest,
                        filename=f"manifest_failed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    )
            except Exception:
                pass
            return ConversionResult(
                success=False,
                message=f"转换失败: {e!s}",
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )
