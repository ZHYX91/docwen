"""
图片格式互转策略模块

支持多种图片格式之间的相互转换。

依赖：
- PIL: 图片处理
- formats.image: 格式转换函数
"""

import contextlib
import logging
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from PIL import Image

from docwen.translation import t
from docwen.services.error_codes import ERROR_CODE_CONVERSION_FAILED, ERROR_CODE_INVALID_INPUT
from docwen.services.result import ConversionResult
from docwen.utils.path_utils import generate_output_path

from .. import CATEGORY_IMAGE, register_conversion
from ..base_strategy import BaseStrategy
from .utils import get_image_format_description, is_multipage_tiff

logger = logging.getLogger(__name__)


@register_conversion(CATEGORY_IMAGE, CATEGORY_IMAGE)
class ImageFormatConversionStrategy(BaseStrategy):
    """
    通用图片格式转换策略

    支持以下格式的相互转换：
    - JPEG/JPG ↔ PNG ↔ GIF ↔ BMP ↔ TIFF ↔ WebP

    使用 Pillow 库进行格式转换，自动处理：
    - 透明度转换（RGBA → RGB，白色背景）
    - 颜色模式转换（P, LA 等 → RGB）
    - 多页TIFF（自动拆分每页分别输出）
    """

    def execute(
        self,
        file_path: str,
        options: dict[str, Any] | None = None,
        progress_callback: Callable[[str], None] | None = None,
    ) -> ConversionResult:
        """
        执行图片格式转换

        标准化工作流程：
        1. 在临时目录创建输入文件副本（使用 prepare_input_file）
        2. 从副本读取并转换格式
        3. 保存到最终输出目录
        4. 自动清理临时目录

        多页TIFF处理：
        - 检测到多页TIFF时，拆分每一页并分别转换
        - 每页使用标准化命名：原始名_page{N}_{时间戳}_from{格式}.{目标格式}
        - 所有页面共享同一时间戳

        Args:
            file_path: 输入图片文件路径
            options: 转换选项，包含actual_format和target_format
            progress_callback: 进度更新回调函数

        Returns:
            ConversionResult: 转换结果对象
        """
        try:
            # 获取实际格式
            actual_format = options.get("actual_format") if options else None
            if not actual_format:
                from docwen.utils.file_type_utils import detect_actual_file_format

                actual_format = detect_actual_file_format(file_path)

            # 从options中获取目标格式
            target_format = options.get("target_format") if options else None

            if not target_format:
                logger.warning("未能从选项中获取目标格式")
                msg = t("conversion.messages.conversion_failed_no_target")
                return ConversionResult(success=False, message=msg, error_code=ERROR_CODE_INVALID_INPUT, details=msg)

            logger.info(f"图片格式转换: {actual_format.upper()} → {target_format.upper()}")

            if progress_callback:
                progress_callback(
                    t(
                        "conversion.progress.converting_format_to_target",
                        format=actual_format.upper(),
                        target=target_format.upper(),
                    )
                )

            from docwen.utils.workspace_manager import get_output_directory

            output_dir = get_output_directory(file_path)

            compress_mode = options.get("compress_mode", "lossless") if options else "lossless"
            size_limit = options.get("size_limit") if options else None
            size_unit = options.get("size_unit", "KB") if options else "KB"

            # 使用临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 创建输入副本
                from docwen.utils.workspace_manager import prepare_input_file

                temp_input = prepare_input_file(file_path, temp_dir, actual_format)
                logger.debug(f"已创建输入副本: {Path(temp_input).name}")

                # 检测是否为多页TIFF
                if is_multipage_tiff(temp_input):
                    # 多页TIFF：拆分每一页并分别转换
                    return self._convert_multipage_tiff(
                        temp_input,
                        file_path,
                        actual_format,
                        target_format,
                        output_dir,
                        progress_callback,
                        compress_mode,
                        size_limit,
                        size_unit,
                    )
                else:
                    # 单页图片：直接转换
                    return self._convert_single_image(
                        temp_input,
                        file_path,
                        actual_format,
                        target_format,
                        output_dir,
                        progress_callback,
                        compress_mode,
                        size_limit,
                        size_unit,
                    )

        except Exception as e:
            logger.error(f"图片格式转换失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t("conversion.messages.conversion_failed_with_error", error=str(e)),
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )

    def _convert_single_image(
        self,
        temp_input: str,
        original_path: str,
        actual_format: str,
        target_format: str,
        output_dir: str,
        progress_callback: Callable[[str], None] | None = None,
        compress_mode: str = "lossless",
        size_limit: Any = None,
        size_unit: str = "KB",
    ) -> ConversionResult:
        """
        转换单页图片（支持压缩选项）

        Args:
            temp_input: 临时输入文件路径
            original_path: 原始文件路径
            actual_format: 实际格式
            target_format: 目标格式
            output_dir: 输出目录
            progress_callback: 进度回调

        Returns:
            ConversionResult: 转换结果
        """
        import tempfile

        # 使用新的convert_image函数，支持压缩选项
        from docwen.converter.formats.image import convert_image
        from docwen.utils.workspace_manager import write_temp_file_then_finalize

        # 策略层负责生成输出路径（使用原始文件路径）
        description = get_image_format_description(actual_format)
        output_path = generate_output_path(
            original_path,  # 使用原始文件路径生成文件名
            output_dir=output_dir,
            section="",
            add_timestamp=True,
            description=description,
            file_type=target_format,
        )

        compress_options: dict[str, Any] = {"compress_mode": compress_mode}

        if compress_mode == "limit_size":
            compress_options["size_limit"] = size_limit
            compress_options["size_unit"] = size_unit
            logger.info(f"使用压缩模式: {compress_options['size_limit']}{compress_options['size_unit']}")
        else:
            logger.info("使用最高质量模式")

        try:
            with tempfile.TemporaryDirectory() as temp_dir:

                def _writer(p: str) -> None:
                    convert_image(
                        source_path=temp_input,
                        target_format=target_format,
                        output_path=p,
                        options=compress_options,
                    )

                saved_path, _ = write_temp_file_then_finalize(
                    temp_dir=temp_dir,
                    target_path=output_path,
                    original_input_file=original_path,
                    writer=_writer,
                )

            if not saved_path:
                raise OSError("保存输出文件失败")

            logger.info(f"✓ 格式转换成功: {Path(saved_path).name}")

            # 标准化扩展名用于消息
            file_ext = self._normalize_extension(target_format)

            return ConversionResult(
                success=True,
                output_path=saved_path,
                message=t("conversion.messages.conversion_to_format_success", format=file_ext.upper()),
            )

        except Exception as e:
            logger.error(f"图片转换失败: {e}", exc_info=True)
            return ConversionResult(
                success=False,
                message=t("conversion.messages.conversion_failed_with_error", error=str(e)),
                error=e,
                error_code=ERROR_CODE_CONVERSION_FAILED,
                details=str(e) or None,
            )

    def _convert_multipage_tiff(
        self,
        temp_input: str,
        original_path: str,
        actual_format: str,
        target_format: str,
        output_dir: str,
        progress_callback: Callable[[str], None] | None = None,
        compress_mode: str = "lossless",
        size_limit: Any = None,
        size_unit: str = "KB",
    ) -> ConversionResult:
        """
        转换多页TIFF，拆分每一页并分别转换

        Args:
            temp_input: 临时输入文件路径
            original_path: 原始文件路径
            actual_format: 实际格式
            target_format: 目标格式
            output_dir: 输出目录
            progress_callback: 进度回调

        Returns:
            ConversionResult: 转换结果（返回第一页的路径）
        """
        output_files = []

        # 标准化文件扩展名
        file_ext = self._normalize_extension(target_format)
        pil_format = self._get_pil_format(target_format)

        try:
            import tempfile

            from docwen.utils.workspace_manager import finalize_output

            with tempfile.TemporaryDirectory() as temp_dir, Image.open(temp_input) as img:
                n_frames = getattr(img, "n_frames", 1)
                logger.info(f"检测到多页TIFF（共{n_frames}页），开始拆分转换")

                for i in range(n_frames):
                    img.seek(i)
                    page_num = i + 1

                    if progress_callback:
                        progress_callback(
                            t(
                                "conversion.progress.converting_page_to_format",
                                page=page_num,
                                format=file_ext.upper(),
                            )
                        )

                    description = get_image_format_description(actual_format)
                    final_page_output_path = generate_output_path(
                        original_path,
                        output_dir=output_dir,
                        section=f"page{page_num}",
                        add_timestamp=True,
                        description=description,
                        file_type=file_ext,
                    )
                    temp_page_output_path = str(Path(temp_dir) / Path(final_page_output_path).name)

                    logger.debug(f"第{page_num}页输出路径: {Path(final_page_output_path).name}")

                    frame = img.copy()
                    converted_frame = self._convert_image_mode(frame, target_format)

                    if compress_mode == "limit_size":
                        if size_limit is None:
                            raise ValueError("压缩模式下必须提供size_limit参数")
                        try:
                            size_limit_int = int(size_limit)
                        except (TypeError, ValueError):
                            raise ValueError("压缩模式下size_limit必须为整数") from None
                        if size_limit_int <= 0:
                            raise ValueError("压缩模式下size_limit必须为正整数")

                        if target_format in ["jpg", "jpeg", "webp"]:
                            from docwen.converter.formats.image.compression import (
                                bytes_to_kb,
                                compress_to_size,
                                get_file_size,
                            )

                            ok = compress_to_size(
                                converted_frame, temp_page_output_path, target_format, size_limit_int, size_unit
                            )
                            if not ok:
                                final_size = get_file_size(temp_page_output_path)
                                raise ValueError(
                                    f"无法将第{page_num}页压缩到目标大小: 目标={size_limit_int}{size_unit}, 实际={bytes_to_kb(final_size):.2f}KB"
                                )
                        else:
                            from docwen.converter.formats.image.compression import (
                                bytes_to_kb,
                                get_file_size,
                                should_compress,
                            )

                            save_kwargs = self._get_save_kwargs(target_format)
                            converted_frame.save(temp_page_output_path, format=pil_format, **save_kwargs)
                            final_size = get_file_size(temp_page_output_path)
                            if should_compress(final_size, size_limit_int, size_unit):
                                raise ValueError(
                                    f"{file_ext.upper()}格式无法可靠压缩到目标大小（第{page_num}页）: 目标={size_limit_int}{size_unit}, 实际={bytes_to_kb(final_size):.2f}KB"
                                )
                    else:
                        save_kwargs = self._get_save_kwargs(target_format)
                        converted_frame.save(temp_page_output_path, format=pil_format, **save_kwargs)

                    with contextlib.suppress(Exception):
                        converted_frame.close()
                    with contextlib.suppress(Exception):
                        frame.close()

                    saved_path, _ = finalize_output(
                        temp_page_output_path,
                        final_page_output_path,
                        original_input_file=original_path,
                    )
                    if not saved_path:
                        raise OSError("保存输出文件失败")

                    output_files.append(saved_path)
                    logger.info(f"✓ 第{page_num}页转换成功: {Path(saved_path).name}")

            logger.info(f"多页TIFF转换完成，共转换{len(output_files)}页")

            if progress_callback:
                progress_callback(t("conversion.progress.tiff_completed", count=len(output_files)))

            # 返回第一页的路径作为主输出
            return ConversionResult(
                success=True,
                output_path=output_files[0] if output_files else None,
                message=t(
                    "conversion.messages.conversion_to_format_pages_success",
                    format=file_ext.upper(),
                    count=len(output_files),
                ),
            )

        except Exception as e:
            logger.error(f"多页TIFF转换失败: {e}", exc_info=True)
            # 清理已创建的文件
            for output_file in output_files:
                try:
                    out_path = Path(output_file)
                    if out_path.exists():
                        out_path.unlink()
                        logger.debug(f"已删除: {output_file}")
                except Exception as cleanup_error:
                    logger.warning(f"清理文件失败: {output_file}, 错误: {cleanup_error}")
            raise

    def _convert_image_mode(self, img, target_format: str):
        """
        根据目标格式转换图片颜色模式

        Args:
            img: PIL Image对象
            target_format: 目标格式（jpeg, png, gif等）

        Returns:
            转换后的PIL Image对象
        """
        target_format = target_format.lower()

        # JPEG不支持透明度，需要转为RGB
        if target_format in ["jpeg", "jpg"]:
            if img.mode in ("RGBA", "LA", "P"):
                logger.debug(f"JPEG不支持透明度，{img.mode} → RGB（白色背景）")
                # 创建白色背景
                background = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                if img.mode in ("RGBA", "LA"):
                    background.paste(img, mask=img.split()[-1])  # 使用alpha通道
                else:
                    background.paste(img)
                return background
            elif img.mode != "RGB":
                logger.debug(f"{img.mode} → RGB")
                return img.convert("RGB")
            return img

        # GIF支持透明度但只有256色
        elif target_format == "gif":
            if img.mode not in ("P", "L"):
                logger.debug(f"{img.mode} → P（256色）")
                return img.convert("P")
            return img

        # BMP不支持透明度
        elif target_format == "bmp":
            if img.mode in ("RGBA", "LA", "P"):
                logger.debug(f"BMP不支持透明度，{img.mode} → RGB（白色背景）")
                background = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                if img.mode in ("RGBA", "LA"):
                    background.paste(img, mask=img.split()[-1])
                else:
                    background.paste(img)
                return background
            elif img.mode != "RGB":
                logger.debug(f"{img.mode} → RGB")
                return img.convert("RGB")
            return img

        # PNG和WebP支持透明度，保持RGBA
        elif target_format in ["png", "webp"]:
            if img.mode in ("P", "L", "LA"):
                logger.debug(f"{img.mode} → RGBA")
                return img.convert("RGBA")
            return img

        # TIFF支持多种模式，保持原样或转RGB
        elif target_format in ["tiff", "tif"]:
            if img.mode in ("P", "LA"):
                logger.debug(f"{img.mode} → RGB")
                return img.convert("RGB")
            return img

        # 默认：保持原样
        return img

    def _normalize_extension(self, format_name: str) -> str:
        """
        标准化文件扩展名（统一使用简短格式）

        Args:
            format_name: 格式名称

        Returns:
            标准化的扩展名
        """
        format_map = {
            "jpeg": "jpg",
            "jpg": "jpg",
            "tiff": "tif",
            "tif": "tif",
        }
        return format_map.get(format_name.lower(), format_name.lower())

    def _get_pil_format(self, format_name: str) -> str:
        """
        获取PIL库使用的格式名称

        Args:
            format_name: 格式名称

        Returns:
            PIL格式名称（大写）
        """
        # PIL使用JPEG和TIFF作为格式名称
        format_map = {
            "jpeg": "JPEG",
            "jpg": "JPEG",
            "tiff": "TIFF",
            "tif": "TIFF",
            "png": "PNG",
            "gif": "GIF",
            "bmp": "BMP",
            "webp": "WEBP",
        }
        return format_map.get(format_name.lower(), format_name.upper())

    def _get_save_kwargs(self, target_format: str) -> dict:
        """
        获取保存图片时的参数

        Args:
            target_format: 目标格式

        Returns:
            保存参数字典
        """
        target_format = target_format.lower()

        if target_format in ["jpeg", "jpg"]:
            return {"quality": 95, "optimize": True}
        elif target_format == "png":
            return {"optimize": True}
        elif target_format == "webp":
            return {"quality": 95, "method": 6}  # method=6 最高质量
        elif target_format in ["tiff", "tif"]:
            return {"compression": "tiff_lzw"}  # 使用LZW无损压缩
        else:
            return {}
