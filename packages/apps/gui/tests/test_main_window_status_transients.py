"""Regressions for MainWindow task-status transient lifecycle."""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.unit


@pytest.mark.parametrize("terminal_kind", ["completed", "failed", "cancelled"])
def test_terminal_status_clears_persistent_main_window_work_transients(
    main_window,
    terminal_kind: str,
) -> None:
    from docwen_gui.i18n import t

    info_vm = main_window._info_area_vm
    info_vm.set_task_summary(
        operation_id="task-1",
        current_file="rules.docx",
        completed_count=1,
        total_count=1,
        state="success",
        tone="warning",
    )
    main_window._on_status_message_changed(t("main_window.task_processing_prefix") + " rules.docx")
    main_window._on_status_message_changed(t("main_window.task_progress_prefix") + "90% Finalizing output")
    assert "processing:main-window" in info_vm._transient_messages
    assert "progress:main-window" in info_vm._transient_messages

    if terminal_kind == "completed":
        terminal_message = t("main_window.task_completed_status")
    elif terminal_kind == "failed":
        terminal_message = t("main_window.task_failed_prefix") + "boom"
    else:
        terminal_message = t("main_window.task_cancelled_status")
    main_window._on_status_message_changed(terminal_message)

    assert "processing:main-window" not in info_vm._transient_messages
    assert "progress:main-window" not in info_vm._transient_messages
    assert "processing:main-window" not in info_vm._pending_transient_updates
    assert "progress:main-window" not in info_vm._pending_transient_updates
    assert "terminal:main-window" not in info_vm._transient_messages
    assert "error:main-window" not in info_vm._transient_messages
    assert info_vm.status_source == "task"
    assert info_vm.status_tone == "warning"
