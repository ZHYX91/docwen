"""One read-only projection of task outcomes and session notices."""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication

from docwen_gui.i18n import t
from docwen_gui.styles.theme_semantics import get_status_theme_class, get_theme_class_color

from .info_area_vm import InfoAreaViewModel
from .task_history import TaskHistory


def activity_status_label(status: str) -> str:
    return {
        "info": t("activity.info"),
        "warning": t("activity.warning"),
        "pending": t("components.file_drop.status.pending"),
        "processing": t("components.file_drop.status.processing"),
        "completed": t("components.file_drop.status.completed"),
        "failed": t("components.file_drop.status.failed"),
        "skipped": t("components.file_drop.status.skipped"),
        "cancelled": t("components.file_drop.status.cancelled"),
    }[status]


@dataclass(frozen=True)
class ActivityRecord:
    key: str
    operation_id: str
    timestamp: datetime
    status: str
    source_path: str
    output_path: str
    operation: str
    details: str
    output_paths: tuple[str, ...] = ()

    @property
    def status_label(self) -> str:
        return activity_status_label(self.status)


class ActivityRecordsModel(QAbstractTableModel):
    """The view never copies execution options or credentials into diagnostics."""

    records_changed = Signal()

    def __init__(self, history: TaskHistory, feedback: InfoAreaViewModel, parent=None) -> None:
        super().__init__(parent)
        self.history = history
        self.feedback = feedback
        self.records: list[ActivityRecord] = []
        history.changed.connect(self.refresh)
        feedback.state_changed.connect(self.refresh)
        self.refresh()

    def refresh(self) -> None:
        notices = self.feedback.history_rows
        known_operations = {key for key, _ in self.history.entries}
        records = []
        for operation_id, task in self.history.entries:
            target = str(task.context.get("target_format") or "").upper()
            operation = t("conversion_panel.convert") + (f" → {target}" if target else "")
            action = str(task.context.get("action_name") or "")
            if action and action != "convert":
                if "proofread" in action:
                    operation = t("conversion_panel.document.proofread_button")
                elif "merge" in action:
                    operation = t("activity.merge") + (f" → {target}" if target else "")
                elif "split" in action:
                    operation = t("activity.split")
                elif "number" in action:
                    operation = t("activity.numbering")
                elif "validate" in action:
                    operation = t("activity.validate")
            for path in task.paths:
                if not path:
                    continue
                outcome = task.outcomes.get(path)
                output = outcome.output_path if outcome else ""
                outputs = outcome.output_paths if outcome else ()
                matching = [
                    row
                    for row in notices
                    if row.operation_id == operation_id
                    and (
                        (
                            row.file_path
                            and Path(row.file_path).as_posix() in {path, Path(output).as_posix() if output else ""}
                        )
                        or len(task.paths) == 1
                    )
                ]
                warnings = outcome.warnings if outcome else ()
                messages = list(
                    dict.fromkeys(row.message for row in matching if not warnings or row.message_type != "warning")
                )
                has_warning = bool(warnings) or any(row.message_type == "warning" for row in matching)
                details = [f"{t('activity.input')}: {path}"]
                if output:
                    details.extend(f"{t('activity.output')}: {path}" for path in outputs or (output,))
                if outcome and outcome.error_message:
                    details.append(outcome.error_message)
                if outcome and outcome.skip_reason:
                    details.append(outcome.skip_reason)
                details.extend(warnings)
                details.extend(messages)
                records.append(
                    ActivityRecord(
                        f"{operation_id}:{path}",
                        operation_id,
                        outcome.updated_at if outcome else task.created_at,
                        "warning"
                        if outcome and outcome.status == "completed" and has_warning
                        else outcome.status
                        if outcome
                        else "pending",
                        path,
                        output,
                        operation,
                        "\n\n".join(dict.fromkeys(details)),
                        outputs,
                    )
                )
        for row in notices:
            if row.operation_id in known_operations:
                task = self.history.get(row.operation_id)
                # Unscoped batch warnings are task notices, not failures of
                # every file in the operation. Per-file details stay scoped.
                if task is None or len(task.paths) == 1 or row.file_path or row.message_type != "warning":
                    continue
            status = {"success": "completed", "danger": "failed", "warning": "warning"}.get(row.message_type, "info")
            details = row.message
            if row.repeat_count > 1:
                details += "\n" + t("info_area.history_repeated", count=row.repeat_count)
            path = row.file_path or row.navigate_file_path
            if path:
                details += "\n\n" + path
            records.append(
                ActivityRecord(
                    f"notice:{row.created_at.isoformat()}:{row.message}",
                    row.operation_id,
                    row.created_at,
                    status,
                    "",
                    path if row.show_location else "",
                    t("activity.info"),
                    details,
                )
            )
        if records == self.records:
            return
        self.beginResetModel()
        self.records = records
        self.endResetModel()
        self.records_changed.emit()

    @property
    def failed_count(self) -> int:
        return sum(row.status == "failed" for row in self.records)

    def clear_finished(self) -> None:
        self.feedback.clear_history()
        self.history.clear_finished()

    def rowCount(self, parent=None) -> int:
        return 0 if parent is not None and parent.isValid() else len(self.records)

    def columnCount(self, parent=None) -> int:
        return 0 if parent is not None and parent.isValid() else 4

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        record = self.records[index.row()]
        if role == Qt.ItemDataRole.ForegroundRole and index.column() == 1:
            theme = "dark" if QApplication.palette().base().color().lightness() < 128 else "light"
            tone = record.status if record.status in {"info", "warning"} else get_status_theme_class(record.status)
            return QColor(get_theme_class_color(tone, theme))
        if role == Qt.ItemDataRole.UserRole + 1:
            return (record.timestamp.isoformat(), record.status_label, record.source_path.casefold(), record.operation)[
                index.column()
            ]
        if role == Qt.ItemDataRole.UserRole:
            return record
        if role == Qt.ItemDataRole.ToolTipRole:
            return record.details
        if role == Qt.ItemDataRole.DisplayRole:
            return (
                record.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                record.status_label,
                Path(record.source_path).name if record.source_path else "—",
                record.operation,
            )[index.column()]
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return (t("activity.time"), t("activity.status"), t("activity.file"), t("activity.operation"))[section]
        return None
