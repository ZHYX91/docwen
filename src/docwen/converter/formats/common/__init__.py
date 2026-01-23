"""
格式转换公共模块

提供办公软件检测和三级容错转换框架。

模块:
- detection: 软件可用性检测
- fallback: 三级容错转换框架
"""

from .detection import (
    check_office_availability,
    OfficeSoftwareNotFoundError,
)

from .fallback import (
    convert_with_fallback,
    convert_with_libreoffice,
    try_com_conversion,
    build_com_converters,
)

__all__ = [
    # detection
    'check_office_availability',
    'OfficeSoftwareNotFoundError',
    # fallback
    'convert_with_fallback',
    'convert_with_libreoffice',
    'try_com_conversion',
    'build_com_converters',
]
