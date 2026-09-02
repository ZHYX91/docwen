from __future__ import annotations

from pathlib import Path

import pytest
from tools.validation.source_family import read_source_text

pytestmark = pytest.mark.contract

ROOT = Path(__file__).resolve().parents[2]


def _read(relative: str) -> str:
    return read_source_text(ROOT / relative)


def test_shared_converter_serializes_only_its_own_request_state() -> None:
    converter = _read("packages/plugins/document/src/docwen_plugin_document/to_markdown/converter.py")
    plugin = _read("packages/plugins/document/src/docwen_plugin_document/plugin.py")

    assert "self._conversion_lock = RLock()" in converter
    public_convert = converter[converter.index("    def convert(") : converter.index("    def _convert_once(")]
    assert "with self._conversion_lock:" in public_convert
    assert "return self._convert_once(context)" in public_convert
    assert "return DocxToMarkdownConverter().convert(context)" in plugin


def test_same_instance_and_distinct_instance_concurrency_regressions_remain_owned() -> None:
    tests = _read("packages/plugins/document/tests/test_request_scoped_docx_policy_*.py")

    assert "test_shared_converter_serializes_requests_to_protect_instance_policy" in tests
    assert "test_warmed_document_plugin_keeps_parallel_request_policies_isolated" in tests
    assert "overlapped = request_b_parsing.wait" in tests
