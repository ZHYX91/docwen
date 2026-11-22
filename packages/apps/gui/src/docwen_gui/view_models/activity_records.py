"""One read-only projection of task outcomes and session notices."""

from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QAbstractTableModel, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication

from docwen_gui.diagnostics import DiagnosticSummary
from docwen_gui.i18n import t
from docwen_gui.path_identity import display_path
from docwen_gui.styles.theme_semantics import get_status_theme_class, get_theme_class_color

from .info_area_vm import HistoryRowData, InfoAreaViewModel
from .task_history import TaskHistory, TaskRecord


def activity_status_label(status: str) -> str:
    match status:
        case "info":
            return t("activity.info")
        case "not_started":
            return t("activity.not_started")
        case "warning":
            return f"⚠ {t('activity.warning')}"
        case "pending":
            return t("components.file_drop.status.pending")
        case "processing":
            return t("components.file_drop.status.processing")
        case "completed":
            return t("components.file_drop.status.completed")
        case "failed":
            return t("components.file_drop.status.failed")
        case "skipped":
            return t("components.file_drop.status.skipped")
        case "cancelled":
            return t("components.file_drop.status.cancelled")
        case _:
            raise KeyError(status)


def _notices_by_source(task: TaskRecord, notices: list[HistoryRowData]) -> dict[str, list[HistoryRowData]]:
    paths = task.paths
    if not notices:
        return {}
    if len(paths) == 1:
        return {paths[0]: notices}

    # A notice can refer to either the input or the primary output. Shared
    # outputs may belong to several sources. Preserve the original notice order.
    sources_by_location: dict[str, set[str]] = defaultdict(set)
    for path in paths:
        sources_by_location[path].add(path)
        outcome = task.outcomes.get(path)
        if outcome and outcome.output_path:
            sources_by_location[Path(outcome.output_path).as_posix()].add(path)
    by_source: dict[str, list[HistoryRowData]] = defaultdict(list)
    for row in notices:
        if row.file_path:
            for source in sources_by_location.get(Path(row.file_path).as_posix(), ()):
                by_source[source].append(row)
    return by_source


def _notice_details(message: str, message_type: str) -> str:
    if message_type == "warning":
        return f"{activity_status_label('warning')}: {message}"
    return message


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
    source_label: str = ""
    source_location_available: bool = True
    output_paths: tuple[str, ...] = ()
    diagnostic: DiagnosticSummary | None = None
    warning_messages: tuple[str, ...] = ()

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
        feedback.history_changed.connect(self.refresh)
        self.refresh()

    def refresh(self) -> None:
        notices = self.feedback.history_rows
        notices_by_operation: dict[str, list[HistoryRowData]] = defaultdict(list)
        for row in notices:
            if row.event_kind == "notice":
                notices_by_operation[row.operation_id].append(row)
        entries = self.history.entries
        known_operations = {key for key, _ in entries}
        records = []
        for operation_id, task in entries:
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
            source_labels = task.context.get("source_labels", {})
            synthetic_paths = {
                Path(item).as_posix() for item in task.context.get("synthetic_input_paths", []) if isinstance(item, str)
            }
            by_source = _notices_by_source(task, notices_by_operation.get(operation_id, []))
            for path in task.paths:
                if not path:
                    continue
                outcome = task.outcomes.get(path)
                source_label = str(source_labels.get(path) or "") if isinstance(source_labels, dict) else ""
                source_location_available = path not in synthetic_paths
                output = outcome.output_path if outcome else ""
                outputs = outcome.output_paths if outcome else ()
                matching = by_source.get(path, [])
                warnings = outcome.warnings if outcome else ()
                messages = list(
                    dict.fromkeys(
                        _notice_details(row.details_message or row.message, row.message_type)
                        for row in matching
                        if not warnings or row.message_type != "warning"
                    )
                )
                has_warning = bool(warnings) or any(row.message_type == "warning" for row in matching)
                input_display = source_label or display_path(path)
                details = [f"{t('activity.input')}: {input_display}"]
                if output:
                    details.extend(f"{t('activity.output')}: {display_path(path)}" for path in outputs or (output,))
                if outcome and outcome.error_message:
                    details.append(outcome.error_message)
                if outcome and outcome.skip_reason:
                    details.append(outcome.skip_reason)
                if summary := task.context.get("request_summary"):
                    details.append(str(summary))
                if outcome:
                    details.extend(outcome.local_details)
                details.extend(_notice_details(message, "warning") for message in warnings)
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
                        source_label,
                        source_location_available,
                        outputs,
                        outcome.diagnostic if outcome else DiagnosticSummary(status="pending"),
                        tuple(
                            dict.fromkeys(
                                (*warnings, *(row.message for row in matching if row.message_type == "warning"))
                            )
                        ),
                    )
                )
        for row in notices:
            if row.event_kind != "notice":
                continue
            if row.operation_id in known_operations:
                task = self.history.get(row.operation_id)
                # Unscoped batch warnings are task notices, not failures of
                # every file in the operation. Per-file details stay scoped.
                if task is None or len(task.paths) == 1 or row.file_path or row.message_type != "warning":
                    continue
            status = {"success": "completed", "danger": "failed", "warning": "warning"}.get(row.message_type, "info")
            if row.diagnostic is not None and row.diagnostic.status == "not_started":
                status = "not_started"
            details = _notice_details(row.details_message or row.message, row.message_type)
            if row.repeat_count > 1:
                details += "\n" + t("info_area.history_repeated", count=row.repeat_count)
            path = row.file_path or row.navigate_file_path
            if path:
                details += "\n\n" + display_path(path)
            output_notice = bool(row.navigate_file_path)
            records.append(
                ActivityRecord(
                    f"notice:{row.created_at.isoformat()}:{row.message}",
                    row.operation_id,
                    row.created_at,
                    status,
                    "" if output_notice else path,
                    path if row.show_location and output_notice else "",
                    row.operation or t("activity.info"),
                    details,
                    diagnostic=row.diagnostic or DiagnosticSummary(status=status),
                    warning_messages=(row.message,) if row.message_type == "warning" else (),
                )
            )
        if records == self.records:
            return
        if len(records) == len(self.records) and all(
            old.key == new.key for old, new in zip(self.records, records, strict=True)
        ):
            # Existing task rows keep their identity while results arrive.
            # Let Qt update and re-sort just the affected range.
            changed = [i for i, (old, new) in enumerate(zip(self.records, records, strict=True)) if old != new]
            self.records = records
            self.dataChanged.emit(self.index(changed[0], 0), self.index(changed[-1], self.columnCount() - 1))
        else:
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
        if role == Qt.ItemDataRole.UserRole:
            return record
        if role == Qt.ItemDataRole.ToolTipRole:
            return record.details
        sort = role == Qt.ItemDataRole.UserRole + 1
        if sort or role == Qt.ItemDataRole.DisplayRole:
            column = index.column()
            if column == 0:
                return record.timestamp.isoformat() if sort else record.timestamp.strftime("%Y-%m-%d %H:%M:%S")
            if column == 1:
                return record.status_label
            if column == 2:
                name = record.source_label or Path(record.source_path).name
                return name.casefold() if sort else name or "—"
            return record.operation
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return (t("activity.time"), t("activity.status"), t("activity.file"), t("activity.operation"))[section]
        return None
