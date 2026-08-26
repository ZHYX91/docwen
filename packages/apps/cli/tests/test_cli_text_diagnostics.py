"""Human-readable CLI projection for successful conversion warnings."""

from __future__ import annotations

import pytest

from docwen_core.models.result import ConversionDiagnostic, ConversionResult

pytestmark = pytest.mark.unit


def _warning_result() -> ConversionResult:
    return ConversionResult(
        task_id="warning-result",
        success=True,
        diagnostics=[
            ConversionDiagnostic(
                level="warning",
                message="缺少必需字段：成文日期",
                code="GONGWEN-NEEDS-REVIEW",
            ),
            ConversionDiagnostic(level="info", message="done", code="GONGWEN-OK"),
        ],
    )


def test_single_success_writes_warning_diagnostic_to_stderr(capsys) -> None:
    from docwen_cli.presenters.text_presenter import TextPresenter

    TextPresenter().present_single(_warning_result())

    captured = capsys.readouterr()
    assert captured.out == "转换成功\n"
    assert captured.err == "警告 [GONGWEN-NEEDS-REVIEW]: 缺少必需字段：成文日期\n"


def test_batch_success_warning_includes_input_file(capsys) -> None:
    from docwen_cli.presenters.text_presenter import TextPresenter

    TextPresenter().present_batch([_warning_result()], input_files=["rules.docx"])

    captured = capsys.readouterr()
    assert "  OK  rules.docx\n" in captured.out
    assert "总计: 1 文件  |  成功: 1  |  失败: 0" in captured.out
    assert captured.err == "警告 [GONGWEN-NEEDS-REVIEW]: rules.docx: 缺少必需字段：成文日期\n"


def test_quiet_mode_suppresses_success_and_warning_output(capsys) -> None:
    from docwen_cli.presenters.text_presenter import TextPresenter

    TextPresenter(quiet=True).present_single(_warning_result())

    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err == ""
