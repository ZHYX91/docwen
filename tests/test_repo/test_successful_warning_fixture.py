from __future__ import annotations

import hashlib
from pathlib import Path

import pytest
from docx import Document

pytestmark = pytest.mark.unit


def test_successful_warning_fixture_is_deterministic_and_semantically_nonempty(tmp_path: Path) -> None:
    from scripts.release.successful_warning_fixture import (
        SUCCESSFUL_WARNING_FIXTURE_SHA256,
        write_successful_warning_fixture,
    )

    first = tmp_path / "first.docx"
    second = tmp_path / "second.docx"

    first_hash = write_successful_warning_fixture(first)
    second_hash = write_successful_warning_fixture(second)

    assert first_hash == second_hash == SUCCESSFUL_WARNING_FIXTURE_SHA256
    assert first.read_bytes() == second.read_bytes()
    assert hashlib.sha256(first.read_bytes()).hexdigest() == SUCCESSFUL_WARNING_FIXTURE_SHA256
    paragraphs = [paragraph.text for paragraph in Document(first).paragraphs]
    assert len(paragraphs) == 10
    assert paragraphs[0] == "关于进一步规范公文处理工作的通知"
    assert paragraphs[-1] == "国务院办公厅　　　　2024年1月15日印发"
