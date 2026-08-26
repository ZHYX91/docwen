"""十六进制颜色转 rgba() 工具函数。"""

from __future__ import annotations


def _hex_to_rgba(color: str, alpha: int) -> str:
    """将十六进制颜色转换为 QSS 可用的 rgba() 文本。"""
    value = color.lstrip("#")
    if len(value) != 6:
        raise ValueError(f"Invalid color: {color}")
    red = int(value[0:2], 16)
    green = int(value[2:4], 16)
    blue = int(value[4:6], 16)
    return f"rgba({red}, {green}, {blue}, {alpha})"
