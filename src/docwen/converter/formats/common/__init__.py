"""
格式转换公共模块

提供办公软件检测和三级容错转换框架。

模块:
- detection: 软件可用性检测
- fallback: 三级容错转换框架
"""

from .detection import (
    OfficeSoftwareNotFoundError,
    check_office_availability,
)
from .fallback import (
    build_com_converters,
    convert_with_fallback,
    convert_with_libreoffice,
    try_com_conversion,
)

__all__ = [
    "OfficeSoftwareNotFoundError",
    "build_com_converters",
    # detection
    "check_office_availability",
    # fallback
    "convert_with_fallback",
    "convert_with_libreoffice",
    "try_com_conversion",
]
