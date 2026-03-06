"""
图片格式转换核心模块

提供通用图片格式转换功能，支持：
- 多种图片格式互转（PNG/JPEG/WebP/BMP/GIF/TIFF）
- 最高质量模式（无损或接近无损）
- 压缩模式（限制文件大小）
"""

import logging
from pathlib import Path

from PIL import Image

from .compression import (
    bytes_to_kb,
    compress_to_size,
    get_file_size,
    get_save_params,
    should_compress,
)

logger = logging.getLogger(__name__)


def convert_image(source_path: str, target_format: str, output_path: str, options: dict | None = None) -> str:
    """
    转换图片格式（支持压缩选项）

    参数:
        source_path: 源文件路径（用于读取）
        target_format: 目标格式（小写，如'jpeg', 'png', 'webp'等）
        output_path: 输出文件路径（完整路径，由调用方提供）
        options: 转换选项，包括：
            - compress_mode: 'lossless' (最高质量) 或 'limit_size' (限制大小)
            - size_limit: 文件大小上限（数值，仅compress_mode='limit_size'时有效）
            - size_unit: 单位 'KB' 或 'MB'（仅compress_mode='limit_size'时有效）

    返回:
        str: 转换后的文件路径

    异常:
        RuntimeError: 转换失败时抛出
    """
    try:
        options = options or {}
        compress_mode = options.get("compress_mode", "lossless")

        logger.info(f"开始转换图片: {Path(source_path).name} → {target_format.upper()}")
        logger.debug(f"压缩模式: {compress_mode}")

        with Image.open(source_path) as img:
            logger.debug(f"原始图片: 模式={img.mode}, 尺寸={img.size}")

            output_dir = Path(output_path).parent
            if output_dir != Path():
                output_dir.mkdir(parents=True, exist_ok=True)

            if target_format in ["jpg", "jpeg"] and img.mode in ["RGBA", "LA", "P"]:
                logger.debug(f"JPEG不支持透明通道，将{img.mode}转换为RGB")
                background = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                if img.mode in ["RGBA", "LA"]:
                    background.paste(img, mask=img.split()[-1])
                else:
                    background.paste(img)
                img = background
            elif img.mode not in ["RGB", "RGBA", "L"]:
                logger.debug(f"将{img.mode}模式转换为RGB")
                img = img.convert("RGB")

            if compress_mode == "lossless":
                logger.info("使用最高质量模式")

                if target_format in ["jpg", "jpeg", "webp"]:
                    save_params = get_save_params(target_format, quality=95)
                    logger.debug(f"保存参数: {save_params}")
                    img.save(output_path, **save_params)
                else:
                    save_params = get_save_params(target_format)
                    logger.debug(f"保存参数: {save_params}")
                    img.save(output_path, **save_params)

            elif compress_mode == "limit_size":
                size_limit = options.get("size_limit")
                size_unit = options.get("size_unit", "KB")

                if size_limit is None:
                    raise ValueError("压缩模式下必须提供size_limit参数")
                try:
                    size_limit = int(size_limit)
                except (TypeError, ValueError):
                    raise ValueError("压缩模式下size_limit必须为整数") from None
                if size_limit <= 0:
                    raise ValueError("压缩模式下size_limit必须为正整数")

                logger.info(f"使用压缩模式: 目标大小={size_limit}{size_unit}")

                if target_format in ["jpg", "jpeg", "webp"]:
                    logger.info("开始根据目标大小搜索最优质量参数")
                    ok = compress_to_size(img, output_path, target_format, size_limit, size_unit)
                    if not ok:
                        final_size = get_file_size(output_path)
                        raise ValueError(
                            f"无法将图片压缩到目标大小: 目标={size_limit}{size_unit}, 实际={bytes_to_kb(final_size):.2f}KB"
                        )
                else:
                    logger.warning(f"{target_format.upper()}格式不支持有效压缩，使用无损保存后校验大小")
                    save_params = get_save_params(target_format)
                    img.save(output_path, **save_params)

                    final_size = get_file_size(output_path)
                    if should_compress(final_size, size_limit, size_unit):
                        raise ValueError(
                            f"{target_format.upper()}格式无法可靠压缩到目标大小: 目标={size_limit}{size_unit}, 实际={bytes_to_kb(final_size):.2f}KB"
                        )

            else:
                raise ValueError(f"未知的压缩模式: {compress_mode}")

        # 记录最终文件大小
        final_size = get_file_size(output_path)
        logger.info(f"转换完成: {Path(output_path).name} ({bytes_to_kb(final_size):.2f}KB)")

        return output_path

    except Exception as e:
        logger.error(f"图片转换失败: {e}", exc_info=True)
        raise RuntimeError(f"图片转换失败: {e}") from e
