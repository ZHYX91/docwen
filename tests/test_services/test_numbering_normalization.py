"""
序号归一化单元测试

覆盖 ConversionService 的序号归一化逻辑：
- default/read-config
- remove/keep/none
- scheme 校验与拼写错误提示
- 不支持的类别显式传参时报错
"""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from docwen.errors import InvalidInputError
from docwen.services.use_cases import _normalize_numbering_options


pytestmark = pytest.mark.unit


class _DummyConfigManager:
    def __init__(self) -> None:
        self._defaults = {
            "document": {
                "to_md_remove_numbering": True,
                "to_md_add_numbering": False,
                "to_md_default_scheme": "hierarchical_standard",
            },
            "text": {
                "to_docx_remove_numbering": True,
                "to_docx_add_numbering": False,
                "to_docx_default_scheme": "hierarchical_standard",
                "to_xlsx_remove_numbering": True,
                "to_xlsx_add_numbering": False,
                "to_xlsx_default_scheme": "hierarchical_standard",
            },
        }
        self._schemes = {"gongwen_standard": {}, "hierarchical_standard": {}}

    def get_conversion_defaults(self, section: str) -> dict:
        return dict(self._defaults.get(section, {}))

    def get_heading_schemes(self) -> dict:
        return dict(self._schemes)


def _ctx() -> SimpleNamespace:
    return SimpleNamespace(config_manager=_DummyConfigManager())


def test_document_to_md_default_from_config() -> None:
    options: dict = {}
    _normalize_numbering_options(options, category="document", target_format="md", action_type=None, ctx=_ctx())
    assert options["remove_numbering"] is True
    assert options["add_numbering"] is False
    assert options["numbering_scheme"] == ""


def test_document_to_md_keep_and_none() -> None:
    options: dict = {"clean_numbering": "keep", "add_numbering_mode": "none"}
    _normalize_numbering_options(options, category="document", target_format="md", action_type=None, ctx=_ctx())
    assert options["remove_numbering"] is False
    assert options["add_numbering"] is False
    assert options["numbering_scheme"] == ""


def test_document_to_md_remove_and_scheme() -> None:
    options: dict = {"clean_numbering": "remove", "add_numbering_mode": "gongwen_standard"}
    _normalize_numbering_options(options, category="document", target_format="md", action_type=None, ctx=_ctx())
    assert options["remove_numbering"] is True
    assert options["add_numbering"] is True
    assert options["numbering_scheme"] == "gongwen_standard"


def test_md_numbering_uses_text_docx_defaults() -> None:
    options: dict = {}
    _normalize_numbering_options(options, category="markdown", target_format=None, action_type="process_md_numbering", ctx=_ctx())
    assert options["remove_numbering"] is True
    assert options["add_numbering"] is False


def test_unsupported_category_rejects_user_provided_modes() -> None:
    options: dict = {"clean_numbering": "remove"}
    with pytest.raises(InvalidInputError):
        _normalize_numbering_options(options, category="layout", target_format="md", action_type=None, ctx=_ctx())


def test_add_numbering_suspected_typo_is_reported() -> None:
    options: dict = {"clean_numbering": "default", "add_numbering_mode": "defalt"}
    with pytest.raises(InvalidInputError) as ei:
        _normalize_numbering_options(options, category="document", target_format="md", action_type=None, ctx=_ctx())
    assert "疑似拼写错误" in str(ei.value)
