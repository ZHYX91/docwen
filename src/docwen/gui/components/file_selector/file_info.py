"""
文件信息类模块

封装单个文件的完整信息，包括：
- 基本信息：路径、文件名、大小
- 格式信息：实际格式、类别、扩展名
- 验证信息：是否支持、是否有效、警告消息
- 状态信息：选中状态、处理状态、输出路径

此类用于在批量文件列表中统一管理文件信息。
"""

import logging
from pathlib import Path

from docwen.i18n import t
from docwen.utils.file_type_utils import get_file_info
from docwen.utils.path_utils import get_file_size_formatted

logger = logging.getLogger(__name__)


class FileInfo:
    """
    文件信息类

    封装单个文件的完整信息：
    - 基本信息：路径、文件名、大小
    - 格式信息：实际格式、类别、扩展名
    - 验证信息：是否支持、是否有效、警告消息
    - 状态信息：选中状态、处理状态、输出路径

    此类用于在批量文件列表中统一管理文件信息。

    属性:
        file_path: 文件完整路径
        file_name: 文件名（不含路径）
        file_size: 格式化后的文件大小字符串
        actual_format: 实际文件格式（基于文件内容检测）
        actual_category: 实际文件类别
        extension: 文件扩展名
        extension_category: 基于扩展名的类别
        is_supported: 是否为支持的文件类型
        is_valid: 文件是否有效
        warning_message: 警告消息（如格式不匹配）
        is_selected: 是否被选中
        status: 处理状态 (pending/processing/completed/skipped/failed)
        output_path: 输出文件路径
        skip_reason: 跳过原因
        error_message: 失败错误信息
    """

    def __init__(self, file_path: str):
        """
        初始化文件信息

        读取文件的基本信息和实际格式，进行格式验证。

        参数:
            file_path: 文件的完整路径
        """
        self.file_path = file_path
        self.file_name = Path(file_path).name
        self.file_size = get_file_size_formatted(file_path)

        # 获取完整的文件信息
        file_info = get_file_info(file_path, t_func=t)
        self.actual_format = file_info["actual_format"]
        self.actual_category = file_info["actual_category"]
        self.extension = file_info["extension"]
        self.extension_category = file_info["extension_category"]
        self.is_supported = file_info["is_supported"]
        self.is_valid = file_info["is_valid"]
        self.warning_message = file_info["warning_message"]

        # 选中状态
        self.is_selected = False

        # 处理状态
        self.status = "pending"  # pending/processing/completed/skipped/failed
        self.output_path: str | None = None  # 输出文件路径

        # 跳过和失败的详细信息
        self.skip_reason: str | None = None  # 跳过原因（如："已是XLS格式"）
        self.error_message: str | None = None  # 失败原因（详细错误信息）

        logger.debug(
            f"创建FileInfo: {self.file_path}, 实际类别: {self.actual_category}, 实际格式: {self.actual_format}"
        )

    def to_dict(self) -> dict:
        """
        转换为字典格式

        将文件信息转换为字典，便于传递和序列化。
        兼容现有代码中使用字典的地方。

        返回:
            Dict: 包含所有关键文件信息的字典
        """
        return {
            "path": self.file_path,
            "size": self.file_size,
            "status": self.status,
            "output_path": self.output_path,
            "actual_category": self.actual_category,
            "actual_format": self.actual_format,
            "is_selected": self.is_selected,
            "warning_message": self.warning_message,
        }

    def __str__(self) -> str:
        """返回文件信息的字符串表示"""
        return f"FileInfo({self.file_path}, 类别:{self.actual_category}, 选中:{self.is_selected})"

    def __repr__(self) -> str:
        """返回文件信息的详细表示"""
        return f"FileInfo(path={self.file_path!r}, category={self.actual_category!r}, status={self.status!r})"
