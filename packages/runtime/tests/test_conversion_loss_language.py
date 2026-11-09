"""Loss notices use the request output locale and a bounded catalog lookup."""

from __future__ import annotations

import pytest

from docwen_core.models.request import ConversionRequest
from docwen_runtime._execution_context import RuntimeExecutionContext
from docwen_runtime.config.task_language import output_loss_messages

pytestmark = pytest.mark.unit


@pytest.mark.parametrize(
    "locale",
    ["ar_SA", "de_DE", "en_US", "es_ES", "fr_FR", "ja_JP", "ko_KR", "pt_BR", "ru_RU", "vi_VN", "zh_CN", "zh_TW"],
)
def test_loss_language_catalog(locale):
    messages = output_loss_messages(locale)
    for key in ("formula_placeholder", "note_placeholder", "comment", "summary"):
        assert messages[f"conversion_loss.{key}"]


def test_loss_language_rejects_path():
    assert output_loss_messages("../../other") == output_loss_messages("zh_CN")


def test_request_language_overrides_saved_output_locale():
    context = object.__new__(RuntimeExecutionContext)
    context._request = ConversionRequest(
        request_id="locale",
        input_refs=[],
        target_format="docx",
        options={"locale": "en_US"},
        config_snapshot={"conversion": {"output_locale": "zh_CN"}},
    )
    assert context.loss_messages == output_loss_messages("en_US")
    context._request.options.clear()
    assert context.loss_messages == output_loss_messages("zh_CN")
