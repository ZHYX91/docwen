"""Batch-related dialog functions.

Provides the ``show_batch_add_failed_dialog`` function that displays a
detailed error dialog when some files fail to be added to the batch list.

Placed in ``widgets/`` as these are GUI-level dialog utilities.
"""

from __future__ import annotations

import logging

from PySide6.QtWidgets import QWidget

from docwen_gui.dialogs.feedback import warn
from docwen_gui.i18n import t as _t

logger = logging.getLogger(__name__)


def show_batch_add_failed_dialog(
    parent: QWidget | None,
    failed_files: list[tuple[str, str]],
) -> None:
    """Show a dialog listing files that failed to be added to the batch.

    Args:
        parent: Parent widget for the dialog.
        failed_files: List of ``(file_path, reason)`` tuples.
    """
    if not failed_files:
        return

    title = _t("main_window.batch_add_failed_title", "Batch Add — Failed Files")
    reason_label = _t("main_window.batch_add_failed_reason", "Reason")
    detail_blocks = [f"{file_path}\n  {reason_label}: {reason}" for file_path, reason in failed_files]
    details = "\n\n".join(detail_blocks)

    logger.info("Showing batch add failed dialog for %d files", len(failed_files))

    count = len(failed_files)
    message = _t(
        "main_window.batch_add_failed_message",
        "{count} file(s) could not be added to the batch list.",
        count=count,
    )
    warn(title, message, details=details, parent=parent)
