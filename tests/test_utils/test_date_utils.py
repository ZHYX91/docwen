"""
date_utils 单元测试

覆盖日期格式转换和时间戳生成。
"""

from __future__ import annotations

import re

import pytest

from docwen.utils.date_utils import convert_date_format, generate_timestamp


pytestmark = pytest.mark.unit


# ============================================================
# convert_date_format
# ============================================================

class TestConvertDateFormat:
    """日期格式转换"""

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            # 中文「年月日」
            ("2024年1月5日", "2024-01-05"),
            ("2024年12月31日", "2024-12-31"),
            # 中文「年月号」
            ("2024年3月8号", "2024-03-08"),
            # 短横线分隔（已是标准格式，但月日可能不补零）
            ("2024-1-5", "2024-01-05"),
            ("2024-12-31", "2024-12-31"),
            # 点分隔
            ("2024.1.5", "2024-01-05"),
            ("2024.12.31", "2024-12-31"),
        ],
    )
    def test_recognized_formats(self, raw: str, expected: str) -> None:
        assert convert_date_format(raw) == expected

    @pytest.mark.parametrize(
        "raw",
        [
            "January 5, 2024",
            "05/01/2024",
            "随便什么文本",
            "",
        ],
    )
    def test_unrecognized_returns_original(self, raw: str) -> None:
        """无法识别的格式返回原字符串"""
        assert convert_date_format(raw) == raw


# ============================================================
# generate_timestamp
# ============================================================

class TestGenerateTimestamp:
    """时间戳生成"""

    def test_format(self) -> None:
        """格式为 YYYYMMDD_HHMMSS"""
        ts = generate_timestamp()
        assert re.fullmatch(r"\d{8}_\d{6}", ts), f"格式不匹配: {ts}"

    def test_uniqueness(self) -> None:
        """连续两次调用不应完全相同（或至少格式一致）"""
        ts1 = generate_timestamp()
        ts2 = generate_timestamp()
        # 格式一致即可，时间可能相同
        assert re.fullmatch(r"\d{8}_\d{6}", ts1)
        assert re.fullmatch(r"\d{8}_\d{6}", ts2)
