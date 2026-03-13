"""
图片占位符共享处理模块
"""


def parse_image_payload(payload: str) -> tuple[str, int | None, int | None]:
    """
    解析图片占位符的 payload 部分

    兼容反转义：统一将 \\| 还原为 |，再按 | 切割。

    参数:
        payload: 占位符内容，如 "path\\|width\\|height" 或 "path|width|height"

    返回:
        (image_path, width, height) 三元组
    """
    payload = payload.replace("\\|", "|")
    parts = payload.split("|")
    path = parts[0].strip()
    width = int(parts[1].strip()) if len(parts) > 1 and parts[1].strip().isdigit() else None
    height = int(parts[2].strip()) if len(parts) > 2 and parts[2].strip().isdigit() else None
    return path, width, height
