"""converter 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.utils.file_type_utils import detect_actual_file_format


pytestmark = pytest.mark.unit

@pytest.mark.integration
def test_detect_actual_file_format_for_markdown(project_root: Path) -> None:
    file_path = project_root / "tests" / "fixtures" / "files" / "sample.md"
    assert file_path.is_file()
    assert detect_actual_file_format(str(file_path)) == "md"
