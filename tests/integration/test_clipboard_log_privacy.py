"""Failure paths must not copy clipboard content into either logging channel."""

import logging
from pathlib import Path
from typing import Any

import pytest

from docwen_application.controller import ApplicationController
from docwen_core.detection import inspect_utf8_markdown_snapshot
from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_gui.clipboard_inputs import ClipboardInputStore
from docwen_runtime._execution_context import _RuntimePluginLogger

pytestmark = [pytest.mark.integration, pytest.mark.pr_gate, pytest.mark.release_gate]


@pytest.mark.parametrize("failure", ["malformed-url", "invalid-formula"])
def test_real_clipboard_failures_do_not_log_authored_content(
    tmp_path: Path,
    round_trip_runtime: Any,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
    failure: str,
) -> None:
    sentinel = "CLIPBOARD_SENTINEL_9f6d"
    text = (
        f"# Invalid link\n\n![x](https://[{sentinel}]/private.png)\n"
        if failure == "malformed-url"
        else "# Invalid formula\n\n$\\begin{" + sentinel + "}$\n"
    )
    loggers: list[_RuntimePluginLogger] = []
    original_init = _RuntimePluginLogger.__init__

    def capture_logger(self, task_id: str) -> None:
        original_init(self, task_id)
        loggers.append(self)

    monkeypatch.setattr(_RuntimePluginLogger, "__init__", capture_logger)
    caplog.set_level(logging.DEBUG)
    store = ClipboardInputStore(tmp_path / "managed")
    try:
        snapshot = store.create(text, display_name_template="Clipboard {index}.md")
        source = Path(snapshot.path)
        inspection = inspect_utf8_markdown_snapshot(source)
        request = ConversionRequest(
            request_id="clipboard-private-failure",
            input_refs=[
                FileRef(
                    path=str(source),
                    format="markdown",
                    category="markdown",
                    metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
                )
            ],
            target_format="docx",
            output_policy=OutputPolicy(output_dir=str(tmp_path / "output")),
        )
        result = ApplicationController(runtime_port=round_trip_runtime).execute_single(request)
        if failure == "malformed-url":
            assert not result.success
            assert result.error is not None and result.error.diagnostic_code == "MD2DOCX-ERROR"
            assert any(item["level"] == "error" for logger in loggers for item in logger.messages)
        else:
            assert result.success, result.error
            assert "LaTeX to MathML failed" in caplog.text
        assert loggers, "The actual Runtime logger must be observed"
        assert sentinel not in caplog.text
        assert sentinel not in repr([logger.messages for logger in loggers])
        assert sentinel not in repr(result.diagnostics)
        assert sentinel not in repr(result.error)
        assert source.read_bytes() == text.encode("utf-8")
    finally:
        store.close()
