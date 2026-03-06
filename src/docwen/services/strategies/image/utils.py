"""
图片策略辅助函数模块

提供图片转换策略所需的辅助函数：
- 格式描述生成
- 多页TIFF检测和页面提取

依赖：
- PIL: 图片处理
- path_utils: 输出路径生成
"""

import logging
from collections.abc import Callable
from pathlib import Path

from docwen.translation import t

logger = logging.getLogger(__name__)


def get_image_format_description(actual_format: str) -> str:
    """
    根据图片格式生成描述（使用真实格式）

    参数:
        actual_format: 图片的真实格式（如 'png', 'jpg', 'bmp'）

    返回:
        str: 格式描述，如 'fromPng', 'fromJpg' 等
    """
    format_map = {
        "png": "fromPng",
        "jpg": "fromJpg",
        "jpeg": "fromJpeg",
        "tif": "fromTif",
        "tiff": "fromTiff",
        "gif": "fromGif",
        "webp": "fromWebp",
        "bmp": "fromBmp",
        "heic": "fromHeic",
        "heif": "fromHeif",
    }

    return format_map.get(actual_format.lower() if actual_format else "", "fromImage")


def is_multipage_tiff(file_path: str) -> bool:
    """
    检测TIFF是否为多页

    参数:
        file_path: 图片文件路径

    返回:
        bool: 是否为多页TIFF
    """
    try:
        from PIL import Image

        from docwen.utils.file_type_utils import detect_actual_file_format

        actual_format = detect_actual_file_format(file_path)
        if actual_format != "tiff":
            return False

        with Image.open(file_path) as img:
            # 尝试获取帧数
            n_frames = getattr(img, "n_frames", 1)
            return n_frames > 1
    except Exception as e:
        logger.warning(f"检测TIFF页数失败: {e}")
        return False


def extract_tiff_pages(
    file_path: str,
    output_dir: str | None = None,
    actual_format: str = "tiff",
    progress_callback: Callable[[str], None] | None = None,
) -> list[tuple[int, str]]:
    """
    提取TIFF每一页为独立的PNG文件，使用标准化命名规则

    文件命名格式：{原始基础名}_page{N}_{时间戳}_from{格式}.png
    示例：报告_page1_20250109_123000_fromTiff.png

    参数:
        file_path: TIFF文件路径
        output_dir: 输出目录（必需）。拆分的PNG文件将保存到此目录
        actual_format: 实际文件格式（默认'tiff'），用于生成描述信息
        progress_callback: 进度回调函数（可选）

    返回:
        list: 文件路径列表 [(页码, 文件路径), ...]

    注意:
        - 所有页面共享同一个时间戳，确保同一批次拆分的文件时间戳一致
        - 为避免时间戳冲突，在循环外生成统一时间戳，然后手动构建文件名
    """
    if not output_dir:
        raise ValueError("extract_tiff_pages 需要指定 output_dir 参数")

    temp_files = []

    try:
        from PIL import Image

        with Image.open(file_path) as img:
            n_frames = getattr(img, "n_frames", 1)
            logger.info(f"TIFF文件共 {n_frames} 页，开始拆分")

            # 步骤1：生成统一时间戳（在循环外）
            from docwen.utils.path_utils import generate_named_output_path, generate_timestamp, strip_timestamp_suffix

            timestamp = generate_timestamp()
            logger.debug(f"生成统一时间戳: {timestamp}")

            base_name = Path(file_path).stem
            base_name_clean = strip_timestamp_suffix(base_name)
            logger.debug(f"清理后的基础文件名: {base_name_clean}")

            # 步骤3：生成描述
            description = f"from{actual_format.capitalize()}"

            # 步骤4：遍历每一页
            for i in range(n_frames):
                img.seek(i)
                page_num = i + 1

                # 更新进度
                if progress_callback:
                    progress_callback(t("conversion.progress.extracting_page", page=page_num))

                page_path = generate_named_output_path(
                    output_dir=output_dir,
                    base_name=base_name_clean,
                    section=f"page{page_num}",
                    add_timestamp=True,
                    description=description,
                    file_type="png",
                    timestamp_override=timestamp,
                )

                logger.debug(f"第 {page_num} 页输出路径: {Path(page_path).name}")

                # 转换并保存当前帧
                # RGBA模式转为RGB（白色背景）
                frame = img.copy()
                original_mode = frame.mode
                has_transparency = frame.mode in ["RGBA", "LA", "PA"] or (
                    frame.mode == "P" and frame.info.get("transparency") is not None
                )
                if has_transparency:
                    background = Image.new("RGB", frame.size, (255, 255, 255))
                    if frame.mode in ["P", "PA", "LA"]:
                        frame = frame.convert("RGBA")
                    background.paste(frame, mask=frame.split()[-1])
                    frame = background
                    logger.debug(f"第 {page_num} 页: {original_mode}转RGB（白色背景）")
                elif frame.mode != "RGB":
                    frame = frame.convert("RGB")
                    logger.debug(f"第 {page_num} 页: {original_mode}转RGB")

                # 保存为PNG格式
                frame.save(page_path)
                temp_files.append((page_num, page_path))
                logger.info(f"✓ 第 {page_num} 页提取成功: {Path(page_path).name}")

        logger.info(f"TIFF拆分完成，共提取 {len(temp_files)} 页")
        return temp_files

    except Exception as e:
        # 清理已创建的文件
        logger.error("TIFF拆分失败，清理已创建的文件")
        for _, temp_path in temp_files:
            try:
                temp_file = Path(temp_path)
                if temp_file.exists():
                    temp_file.unlink()
                    logger.debug(f"已删除: {temp_path}")
            except Exception as cleanup_error:
                logger.warning(f"清理文件失败: {temp_path}, 错误: {cleanup_error}")

        logger.error(f"提取TIFF页面失败: {e}", exc_info=True)
        raise
