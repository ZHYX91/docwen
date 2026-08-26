from __future__ import annotations

import re
from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

_LOGURU_PERCENT_FORMAT_PATTERN = re.compile(
    r"loguru_logger\.(?:trace|debug|info|success|warning|error|critical|exception)\s*"
    r"\(\s*(?:[rubf]|rb|br|rf|fr)?(?P<quote>['\"])(?:(?! (?P=quote)).|\\.)*?%[a-zA-Z]"
    r"(?:(?! (?P=quote)).|\\.)*?(?P=quote)\s*,",
    re.DOTALL | re.VERBOSE,
)


def test_repo_does_not_use_percent_style_formatting_with_loguru() -> None:
    project_root = Path(__file__).resolve().parents[2]
    violations: list[str] = []

    for relative_dir in ("src", "tests", "tools", "scripts"):
        for path in (project_root / relative_dir).rglob("*.py"):
            content = path.read_text(encoding="utf-8")
            for match in _LOGURU_PERCENT_FORMAT_PATTERN.finditer(content):
                line = content.count("\n", 0, match.start()) + 1
                violations.append(f"{path.relative_to(project_root)}:{line}")

    assert not violations, "发现 loguru_logger 使用了 '%' 风格占位符: " + ", ".join(violations)
