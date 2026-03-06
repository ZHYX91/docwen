from __future__ import annotations

import logging
from pathlib import Path

from docwen.config.config_manager import config_manager
from docwen.utils.markdown_utils import format_base64_image_link

logger = logging.getLogger(__name__)


def build_base64_image_link(image_path: str, style: str = "wiki_embed") -> str:
    path = Path(image_path)
    if not path.exists():
        logger.warning("生成Base64链接失败，文件不存在: %s", image_path)
        return ""

    use_bytes: bytes | None = None
    mime_type: str | None = None

    if config_manager.get_export_base64_compress_enabled():
        threshold_kb = int(config_manager.get_export_base64_compress_threshold_kb())
        original_size = path.stat().st_size
        if original_size > threshold_kb * 1024:
            try:
                # 尝试压缩
                from docwen.converter.formats.image.compression import compress_file_to_bytes

                compressed_bytes = compress_file_to_bytes(
                    str(path),
                    target_format="JPEG",
                    target_size=threshold_kb,
                    unit="KB",
                )
                
                # 检查压缩结果
                if len(compressed_bytes) >= original_size:
                    logger.warning(
                        "Base64 压缩后体积未减小，回退到原图: %s (%sKB -> %sKB)",
                        path.name,
                        original_size // 1024,
                        len(compressed_bytes) // 1024,
                    )
                    # use_bytes 保持为 None，走下方原图逻辑
                else:
                    use_bytes = compressed_bytes
                    mime_type = "image/jpeg"
                    
                    if len(use_bytes) > threshold_kb * 1024:
                        logger.warning(
                            "Base64 图片压缩后仍超出阈值: %s (%sKB > %sKB)",
                            path.name,
                            len(use_bytes) // 1024,
                            threshold_kb,
                        )
            except Exception as e:
                logger.error(f"图片压缩失败，回退到原图: {path.name}, 错误: {e}")
                # use_bytes 保持为 None，走下方原图逻辑

    if use_bytes is None:
        ext = path.suffix.lower().lstrip(".")
        if ext in {"jpg", "jpeg"}:
            mime_type = "image/jpeg"
        elif ext:
            mime_type = f"image/{ext}"
        else:
            mime_type = "image/png"
        use_bytes = path.read_bytes()

    return format_base64_image_link(use_bytes, str(mime_type), path.name, style)
