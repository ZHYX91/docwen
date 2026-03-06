"""GUI 逻辑单元测试。"""

from __future__ import annotations

import pytest

from docwen.gui.core.batch_event_adapter import adapt_batch_event_to_queue_messages
from docwen.services.batch import BatchEvent, FileResult

pytestmark = [pytest.mark.unit, pytest.mark.windows_only]


def test_adapter_maps_skipped_status_without_message_prefix_dependency() -> None:
    event = BatchEvent(
        type="file_finished",
        file_path="a.docx",
        index=1,
        total=1,
        result=FileResult(
            file_path="a.docx",
            success=True,
            status="skipped",
            message="Already DOCX",
        ),
    )

    msgs = adapt_batch_event_to_queue_messages(event)
    assert msgs == [("set_file_status", "a.docx", "skipped", None, "Already DOCX", None)]


def test_adapter_maps_completed_status() -> None:
    event = BatchEvent(
        type="file_finished",
        file_path="a.docx",
        index=1,
        total=1,
        result=FileResult(
            file_path="a.docx",
            success=True,
            status="completed",
            output_path="out.docx",
            message="ok",
        ),
    )

    msgs = adapt_batch_event_to_queue_messages(event)
    assert msgs == [("set_file_status", "a.docx", "completed", "out.docx", None, None)]
