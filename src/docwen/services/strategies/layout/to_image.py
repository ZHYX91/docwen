"""
版式文件转图片格式策略

将PDF/OFD/XPS/CAJ等版式文件的每一页转换为图片。

支持的目标格式：
- PNG：无损压缩，支持透明
- JPG：有损压缩，文件较小
- TIF：支持多页，专业格式

依赖：
- .utils: 预处理函数
- PyMuPDF (fitz): PDF渲染
- PIL: 图片处理（TIF多页）
"""

import contextlib
import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from docwen.translation import t
from docwen.services.error_codes import (
    ERROR_CODE_CONVERSION_FAILED,
    ERROR_CODE_DEPENDENCY_MISSING,
    ERROR_CODE_OPERATION_CANCELLED,
)
from docwen.services.result import ConversionResult
from docwen.services.strategies import CATEGORY_LAYOUT, register_conversion
from docwen.services.strategies.base_strategy import BaseStrategy

from .utils import preprocess_layout_file, should_keep_intermediates

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_LAYOUT, "png")
class LayoutToPngStrategy(BaseStrategy):
    """
    将版式文件的每一页转换为PNG图片的策略

    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 创建输出子文件夹（在临时目录）
    3. 使用PyMuPDF统一渲染所有页面为PNG图片（按指定DPI）
    4. 将子文件夹移动到目标目录
    5. 根据配置决定是否保留中间PDF文件

    输出结构：
    原文件名_时间戳_from原格式/
    ├── page_01.png
    ├── page_02.png
    └── page_03.png

    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - OFD：先转为PDF再处理
    - CAJ：先转为PDF再处理
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行版式文件到PNG图片的转换

        参数:
            file_path: 输入的版式文件路径
            options: 转换选项字典，包含：
                - dpi: 渲染DPI（150/300/600）
                - actual_format: 实际文件格式
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数

        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        dpi = options.get("dpi", 150)
        actual_format = options.get("actual_format")
        cancel_event = options.get("cancel_event")

        if not actual_format:
            from docwen.utils.file_type_utils import detect_actual_file_format

            actual_format = detect_actual_file_format(file_path)

        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing"))

            logger.info(f"开始转换 {actual_format.upper()} 为PNG图片，DPI: {dpi}")

            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                if actual_format != "pdf" and progress_callback:
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

                # 步骤2：生成输出文件夹路径
                from docwen.utils.workspace_manager import get_output_directory

                output_dir = get_output_directory(file_path)

                from docwen.utils.path_utils import generate_output_path

                # 生成标准化文件夹名（使用原格式）
                description = f"from{actual_format.capitalize()}"
                folder_path_with_ext = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="png",
                )

                # 去掉.png扩展名，作为文件夹名
                folder_name = Path(folder_path_with_ext).stem
                temp_output_folder = Path(temp_dir) / folder_name
                temp_output_folder.mkdir(parents=True, exist_ok=True)

                logger.info(f"输出文件夹: {folder_name}")

                # 步骤3：打开PDF并渲染所有页面
                if progress_callback:
                    progress_callback(t("conversion.progress.rendering_pdf_to_png"))

                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    msg = t("conversion.messages.missing_pymupdf")
                    return ConversionResult(
                        success=False, message=msg, error_code=ERROR_CODE_DEPENDENCY_MISSING, details=msg
                    )

                with fitz.open(pdf_path) as doc:
                    total_pages = len(doc)
                    width = len(str(total_pages))  # 计算前导零位数

                    logger.info(f"PDF共{total_pages}页，开始渲染（DPI: {dpi}）")

                    # 计算缩放比例
                    zoom = dpi / 72.0

                    for page_num in range(total_pages):
                        # 检查取消
                        if cancel_event and cancel_event.is_set():
                            return ConversionResult(
                                success=False,
                                message=t("conversion.messages.operation_cancelled"),
                                error_code=ERROR_CODE_OPERATION_CANCELLED,
                            )

                        if progress_callback:
                            progress_callback(
                                t(
                                    "conversion.progress.rendering_page_progress",
                                    current=page_num + 1,
                                    total=total_pages,
                                )
                            )

                        page = doc.load_page(page_num)

                        # 渲染页面
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat, alpha=True)

                        # 生成文件名（带前导零）
                        image_filename = f"page_{str(page_num + 1).zfill(width)}.png"
                        image_path = str(temp_output_folder / image_filename)

                        # 保存图片
                        pix.save(image_path)

                        logger.debug(f"已保存: {image_filename} ({pix.width}x{pix.height})")

                logger.info(f"所有页面已渲染完成，共{total_pages}张图片")

                # 步骤4：移动文件夹到目标目录
                from docwen.utils.path_utils import ensure_unique_directory_path

                final_folder = ensure_unique_directory_path(str(Path(output_dir) / folder_name))
                from docwen.utils.workspace_manager import save_output_with_fallback

                saved_folder, _ = save_output_with_fallback(
                    str(temp_output_folder), final_folder, original_input_file=file_path
                )
                if saved_folder:
                    final_folder = saved_folder
                logger.info(f"文件夹已移动到: {final_folder}")

                # 步骤5：处理中间文件
                saved_items = []
                if should_keep_intermediates():
                    from docwen.utils.workspace_manager import save_intermediate_items

                    saved_items = save_intermediate_items(preprocess_result.intermediates, output_dir, move=True)
                    for _, saved_path in saved_items:
                        logger.info(f"保留中间PDF文件: {saved_path}")

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
                            message=t(
                                "conversion.messages.conversion_to_image_success", format="PNG", count=total_pages
                            ),
                            output_path=final_folder,
                            mask_input=config_manager.get_mask_manifest_input_path(),
                        )
                        write_manifest_json(final_folder, manifest)
                except Exception:
                    pass

                # 返回文件夹路径作为输出路径
                return ConversionResult(
                    success=True,
                    output_path=final_folder,
                    message=t("conversion.messages.conversion_to_image_success", format="PNG", count=total_pages),
                )

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except Exception as e:
            logger.error(f"执行 LayoutToPngStrategy 时出错: {e}", exc_info=True)
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
                message=t("conversion.messages.conversion_failed_with_error", error=str(e)),
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )


@register_conversion(CATEGORY_LAYOUT, "jpg")
class LayoutToJpgStrategy(BaseStrategy):
    """
    将版式文件的每一页转换为JPG图片的策略

    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 创建输出子文件夹（在临时目录）
    3. 使用PyMuPDF统一渲染所有页面为JPG图片（按指定DPI，处理透明度）
    4. 将子文件夹移动到目标目录
    5. 根据配置决定是否保留中间PDF文件

    输出结构：
    原文件名_时间戳_from原格式/
    ├── page_01.jpg
    ├── page_02.jpg
    └── page_03.jpg

    特点：
    - JPG不支持透明，透明背景会转为白色
    - 文件体积比PNG小

    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - OFD：先转为PDF再处理
    - CAJ：先转为PDF再处理
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行版式文件到JPG图片的转换

        参数:
            file_path: 输入的版式文件路径
            options: 转换选项字典，包含：
                - dpi: 渲染DPI（150/300/600）
                - actual_format: 实际文件格式
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数

        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        dpi = options.get("dpi", 150)
        actual_format = options.get("actual_format")
        cancel_event = options.get("cancel_event")

        if not actual_format:
            from docwen.utils.file_type_utils import detect_actual_file_format

            actual_format = detect_actual_file_format(file_path)

        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing"))

            logger.info(f"开始转换 {actual_format.upper()} 为JPG图片，DPI: {dpi}")

            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                if actual_format != "pdf" and progress_callback:
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

                # 步骤2：生成输出文件夹路径
                from docwen.utils.workspace_manager import get_output_directory

                output_dir = get_output_directory(file_path)

                from docwen.utils.path_utils import generate_output_path

                # 生成标准化文件夹名（使用原格式）
                description = f"from{actual_format.capitalize()}"
                folder_path_with_ext = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="jpg",
                )

                # 去掉.jpg扩展名，作为文件夹名
                folder_name = Path(folder_path_with_ext).stem
                temp_output_folder = Path(temp_dir) / folder_name
                temp_output_folder.mkdir(parents=True, exist_ok=True)

                logger.info(f"输出文件夹: {folder_name}")

                # 步骤3：打开PDF并渲染所有页面
                if progress_callback:
                    progress_callback(t("conversion.progress.rendering_pdf_to_jpg"))

                try:
                    import fitz  # PyMuPDF
                except ImportError:
                    msg = t("conversion.messages.missing_pymupdf")
                    return ConversionResult(
                        success=False, message=msg, error_code=ERROR_CODE_DEPENDENCY_MISSING, details=msg
                    )

                with fitz.open(pdf_path) as doc:
                    total_pages = len(doc)
                    width = len(str(total_pages))  # 计算前导零位数

                    logger.info(f"PDF共{total_pages}页，开始渲染（DPI: {dpi}）")

                    # 计算缩放比例
                    zoom = dpi / 72.0

                    for page_num in range(total_pages):
                        # 检查取消
                        if cancel_event and cancel_event.is_set():
                            return ConversionResult(
                                success=False,
                                message=t("conversion.messages.operation_cancelled"),
                                error_code=ERROR_CODE_OPERATION_CANCELLED,
                            )

                        if progress_callback:
                            progress_callback(
                                t(
                                    "conversion.progress.rendering_page_progress",
                                    current=page_num + 1,
                                    total=total_pages,
                                )
                            )

                        page = doc[page_num]

                        # 渲染页面（JPG不支持alpha通道，设置为False）
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat, alpha=False)

                        # 生成文件名（带前导零）
                        image_filename = f"page_{str(page_num + 1).zfill(width)}.jpg"
                        image_path = str(temp_output_folder / image_filename)

                        # 保存图片
                        pix.save(image_path)

                        logger.debug(f"已保存: {image_filename} ({pix.width}x{pix.height})")

                logger.info(f"所有页面已渲染完成，共{total_pages}张图片")

                # 步骤4：移动文件夹到目标目录
                from docwen.utils.path_utils import ensure_unique_directory_path

                final_folder = ensure_unique_directory_path(str(Path(output_dir) / folder_name))
                from docwen.utils.workspace_manager import save_output_with_fallback

                saved_folder, _ = save_output_with_fallback(
                    str(temp_output_folder), final_folder, original_input_file=file_path
                )
                if saved_folder:
                    final_folder = saved_folder
                logger.info(f"文件夹已移动到: {final_folder}")

                # 步骤5：处理中间文件
                saved_items = []
                if should_keep_intermediates():
                    from docwen.utils.workspace_manager import save_intermediate_items

                    saved_items = save_intermediate_items(preprocess_result.intermediates, output_dir, move=True)
                    for _, saved_path in saved_items:
                        logger.info(f"保留中间PDF文件: {saved_path}")

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
                            message=t(
                                "conversion.messages.conversion_to_image_success", format="JPG", count=total_pages
                            ),
                            output_path=final_folder,
                            mask_input=config_manager.get_mask_manifest_input_path(),
                        )
                        write_manifest_json(final_folder, manifest)
                except Exception:
                    pass

                # 返回文件夹路径作为输出路径
                return ConversionResult(
                    success=True,
                    output_path=final_folder,
                    message=t("conversion.messages.conversion_to_image_success", format="JPG", count=total_pages),
                )

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except Exception as e:
            logger.error(f"执行 LayoutToJpgStrategy 时出错: {e}", exc_info=True)
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
                success=False, message=t("conversion.messages.conversion_failed_with_error", error=str(e)), error=e
            )


@register_conversion(CATEGORY_LAYOUT, "tif")
class LayoutToTifStrategy(BaseStrategy):
    """
    将版式文件转换为TIF图片的策略（支持多页TIFF）

    转换流程：
    1. 预处理：CAJ/XPS/OFD → PDF（如果需要）
    2. 使用PyMuPDF渲染所有页面为图片
    3. 使用PIL保存为TIF（自动多页）
    4. 输出单个TIF文件（无需子文件夹）

    特点：
    - 自动判断单页/多页
    - 多页自动打包为单个TIFF文件
    - 无需子文件夹，直接输出文件
    - alpha=False，透明背景转为白色

    支持的输入格式：
    - PDF：直接处理
    - XPS：先转为PDF再处理
    - OFD：先转为PDF再处理
    - CAJ：先转为PDF再处理
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行版式文件到TIF图片的转换

        参数:
            file_path: 输入的版式文件路径
            options: 转换选项字典，包含：
                - dpi: 渲染DPI（150/300/600）
                - actual_format: 实际文件格式
                - cancel_event: 取消事件
            progress_callback: 进度更新回调函数

        返回:
            ConversionResult: 包含转换结果的对象
        """
        options = options or {}
        dpi = options.get("dpi", 150)
        actual_format = options.get("actual_format")
        cancel_event = options.get("cancel_event")

        if not actual_format:
            from docwen.utils.file_type_utils import detect_actual_file_format

            actual_format = detect_actual_file_format(file_path)

        try:
            if progress_callback:
                progress_callback(t("conversion.progress.preparing"))

            logger.info(f"开始转换 {actual_format.upper()} 为TIF图片，DPI: {dpi}")

            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                if actual_format != "pdf" and progress_callback:
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

                # 步骤2：生成输出文件路径（直接在临时目录，无需子文件夹）
                from docwen.utils.workspace_manager import get_output_directory

                output_dir = get_output_directory(file_path)

                from docwen.utils.path_utils import generate_output_path

                # 生成标准化TIF路径（在临时目录）
                description = f"from{actual_format.capitalize()}"
                tif_temp_path = generate_output_path(
                    file_path,
                    output_dir=temp_dir,
                    section="",
                    add_timestamp=True,
                    description=description,
                    file_type="tif",
                )

                logger.info(f"输出TIF文件: {Path(tif_temp_path).name}")

                # 步骤3：打开PDF并渲染所有页面
                if progress_callback:
                    progress_callback(t("conversion.progress.rendering_pdf_pages"))

                try:
                    import fitz  # PyMuPDF
                    from PIL import Image
                except ImportError as e:
                    logger.error(f"内部错误 - 模块导入失败: {e}", exc_info=True)
                    return ConversionResult(
                        success=False, message=t("conversion.messages.conversion_failed_with_error", error=str(e))
                    )

                page_image_paths = []

                with fitz.open(pdf_path) as doc:
                    total_pages = len(doc)
                    logger.info(f"PDF共{total_pages}页，开始渲染（DPI: {dpi}）")

                    # 计算缩放比例
                    zoom = dpi / 72.0

                    for page_num in range(total_pages):
                        # 检查取消
                        if cancel_event and cancel_event.is_set():
                            return ConversionResult(
                                success=False,
                                message=t("conversion.messages.operation_cancelled"),
                                error_code=ERROR_CODE_OPERATION_CANCELLED,
                            )

                        if progress_callback:
                            progress_callback(
                                t(
                                    "conversion.progress.rendering_page_progress",
                                    current=page_num + 1,
                                    total=total_pages,
                                )
                            )

                        page = doc[page_num]

                        # 渲染页面（TIF不需要透明通道，设置alpha=False）
                        mat = fitz.Matrix(zoom, zoom)
                        pix = page.get_pixmap(matrix=mat, alpha=False)

                        # 转换为PIL Image对象并落盘，降低多页时的内存占用
                        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                        page_image_path = str(Path(temp_dir) / f"page_{page_num + 1:04d}.png")
                        img.save(page_image_path, format="PNG", optimize=True)
                        with contextlib.suppress(Exception):
                            img.close()
                        page_image_paths.append(page_image_path)

                        logger.debug(f"已渲染第{page_num + 1}页 ({pix.width}x{pix.height})")

                # 检查取消
                if cancel_event and cancel_event.is_set():
                    return ConversionResult(
                        success=False,
                        message=t("conversion.messages.operation_cancelled"),
                        error_code=ERROR_CODE_OPERATION_CANCELLED,
                    )

                # 步骤4：保存为TIF（自动多页）
                if progress_callback:
                    progress_callback(t("conversion.progress.saving_tif_file"))

                if len(page_image_paths) > 1:
                    opened_images = []
                    try:
                        first = Image.open(page_image_paths[0])
                        opened_images.append(first)
                        rest = []
                        for p in page_image_paths[1:]:
                            im = Image.open(p)
                            opened_images.append(im)
                            rest.append(im)

                        first.save(
                            tif_temp_path, format="TIFF", save_all=True, append_images=rest, compression="tiff_lzw"
                        )
                        logger.info(f"多页TIF已保存: {Path(tif_temp_path).name}，共{len(page_image_paths)}页")
                    finally:
                        for im in opened_images:
                            with contextlib.suppress(Exception):
                                im.close()
                else:
                    with Image.open(page_image_paths[0]) as single:
                        single.save(tif_temp_path, format="TIFF", compression="tiff_lzw")
                    logger.info(f"单页TIF已保存: {Path(tif_temp_path).name}")

                # 步骤5：移动文件到目标目录
                final_tif_path = str(Path(output_dir) / Path(tif_temp_path).name)
                from docwen.utils.workspace_manager import save_output_with_fallback

                final_saved_path, _ = save_output_with_fallback(
                    tif_temp_path, final_tif_path, original_input_file=file_path
                )
                if final_saved_path:
                    final_tif_path = final_saved_path
                logger.info(f"TIF文件已移动到: {final_tif_path}")

                if should_keep_intermediates():
                    from docwen.utils.workspace_manager import save_intermediate_items

                    saved_items = save_intermediate_items(preprocess_result.intermediates, output_dir, move=True)
                    for _, saved_path in saved_items:
                        logger.info(f"保留中间PDF文件: {saved_path}")

                # 返回文件路径作为输出路径
                return ConversionResult(
                    success=True,
                    output_path=final_tif_path,
                    message=t("conversion.messages.conversion_to_tif_success", pages=total_pages),
                )

        except InterruptedError:
            return ConversionResult(
                success=False,
                message=t("conversion.messages.operation_cancelled"),
                error_code=ERROR_CODE_OPERATION_CANCELLED,
            )
        except Exception as e:
            logger.error(f"执行 LayoutToTifStrategy 时出错: {e}", exc_info=True)
            return ConversionResult(
                success=False, message=t("conversion.messages.conversion_failed_with_error", error=str(e)), error=e
            )
