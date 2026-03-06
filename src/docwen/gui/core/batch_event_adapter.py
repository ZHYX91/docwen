from __future__ import annotations

from pathlib import Path

from docwen.i18n import t
from docwen.services.batch import BatchEvent
from docwen.services.error_codes import ERROR_CODE_SKIPPED_SAME_FORMAT


def adapt_batch_event_to_queue_messages(event: BatchEvent) -> list[tuple]:
    if event.type == "file_started":
        return [
            ("set_file_status", event.file_path, "processing", None, None, None),
            (
                "set_progress",
                t(
                    "components.status_bar.processing_file",
                    current=event.index,
                    total=event.total,
                    filename=Path(event.file_path).name,
                ),
            ),
        ]

    if event.type == "file_finished" and event.result:
        r = event.result
        message = r.message or ""
        if r.success:
            if r.status == "skipped":
                if r.error_code == ERROR_CODE_SKIPPED_SAME_FORMAT:
                    fmt = (r.details or "").strip()
                    if fmt:
                        message = t("components.status_bar.skipped_same_format", format=fmt)
                    else:
                        message = t("components.file_dialogs.skip_reason_unknown")
                elif not message:
                    message = t("components.file_dialogs.skip_reason_unknown")
                return [("set_file_status", r.file_path, "skipped", None, message, None)]
            return [("set_file_status", r.file_path, "completed", r.output_path, None, None)]

        if r.status == "cancelled":
            return [
                ("set_file_status", r.file_path, "failed", None, None, t("conversion.messages.operation_cancelled"))
            ]
        return [
            (
                "set_file_status",
                r.file_path,
                "failed",
                None,
                None,
                message or t("conversion.messages.conversion_failed_check_log"),
            )
        ]

    return []
