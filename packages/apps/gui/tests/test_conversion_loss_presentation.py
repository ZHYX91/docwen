"""The existing GUI diagnostic channel keeps occurrence and loss summary text."""

import pytest

from docwen_core.conversion_loss import ConversionLoss, LossCollector
from docwen_core.models.result import ConversionResult
from docwen_gui.execution_presenter import _result_warning_messages

pytestmark = pytest.mark.unit


def test_loss_diagnostics_remain_visible_without_document_comments():
    loss = ConversionLoss(
        "MD2DOCX-FORMULA-TEXT-FALLBACK",
        "Unsupported formula",
        "Placeholder retained",
        "line 4",
        "paragraph 2",
        r"\bad{x}",
    )
    collector = LossCollector()
    collector.add(loss)
    summary = collector.summary()
    assert summary is not None
    result = ConversionResult(task_id="loss", success=True, diagnostics=[loss.to_diagnostic(), summary])
    messages = _result_warning_messages(result)
    assert "Unsupported formula" in messages[0]
    assert "line 4" in messages[0] and "paragraph 2" in messages[0]
    assert messages[1] == summary.message
    assert "Comments were added" not in messages[1]
