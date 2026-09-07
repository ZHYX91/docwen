"""Searchable, modeless activity window with an integrated details pane."""

from pathlib import Path

from PySide6.QtCore import QModelIndex, QSize, QSortFilterProxyModel, Qt, Signal
from PySide6.QtGui import QPalette
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QHeaderView,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QStyle,
    QStyledItemDelegate,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from docwen_gui.i18n import t
from docwen_gui.styles.design_tokens import Sizing, Spacing
from docwen_gui.styles.theme_semantics import apply_theme_class
from docwen_gui.view_models.activity_records import ActivityRecord, ActivityRecordsModel, activity_status_label
from docwen_gui.widgets.panel_card import ChoiceGroup
from docwen_gui.widgets.value_controls import ScrollSafeComboBox


class _ActivityFilter(QSortFilterProxyModel):
    query = ""
    status = ""
    operation = ""
    operation_id = ""
    source_path = ""

    def lessThan(self, left: QModelIndex, right: QModelIndex) -> bool:
        if left.column() == 2:

            def file_key(index: QModelIndex) -> tuple[str, str]:
                path = str(index.data(self.sortRole()) or "")
                return Path(path).name.casefold(), path.casefold()

            return file_key(left) < file_key(right)
        if left.column() == 0 and left.data(self.sortRole()) == right.data(self.sortRole()):
            # A clock tick can contain several outcomes; preserve arrival order.
            return left.row() < right.row()
        return super().lessThan(left, right)

    def filterAcceptsRow(self, source_row: int, source_parent: QModelIndex) -> bool:
        model = self.sourceModel()
        if not isinstance(model, ActivityRecordsModel):
            return False
        record = model.records[source_row]
        return (
            (not self.operation_id or record.operation_id == self.operation_id)
            and (not self.source_path or record.source_path == self.source_path)
            and (not self.status or record.status == self.status)
            and (not self.operation or record.operation == self.operation)
            and (
                not self.query
                or self.query
                in "\n".join(
                    (record.source_path, record.output_path, record.details, record.operation, record.status_label)
                ).casefold()
            )
        )


class _ActivityDelegate(QStyledItemDelegate):
    def sizeHint(self, option, index) -> QSize:
        hint = super().sizeHint(option, index)
        hint.setHeight(max(hint.height(), Sizing.CONTROL_HEIGHT, option.fontMetrics.height() + 2 * Spacing.CONTROL_GAP))
        return hint

    def initStyleOption(self, option, index) -> None:
        super().initStyleOption(option, index)
        if option.state & QStyle.StateFlag.State_Selected:
            # Semantic status colours must yield to the theme's selection contrast.
            palette = QApplication.palette()
            for group in (QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive, QPalette.ColorGroup.Disabled):
                color = palette.color(group, QPalette.ColorRole.HighlightedText)
                option.palette.setColor(group, QPalette.ColorRole.Text, color)
                option.palette.setColor(group, QPalette.ColorRole.HighlightedText, color)


class ActivityRecordsDialog(QDialog):
    location_requested = Signal(str, bool)

    def __init__(self, model: ActivityRecordsModel, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("activityRecordsDialog")
        self.setWindowTitle(t("activity.title"))
        self.setModal(False)
        self.resize(800, 600)
        self.setMinimumSize(420, 360)
        self._model = model
        self._selected_key = ""
        self._resetting = False
        self._proxy = _ActivityFilter(self)
        self._proxy.setSourceModel(model)
        self._proxy.setSortRole(Qt.ItemDataRole.UserRole + 1)
        self._proxy.setSortCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            Spacing.CARD_PADDING, Spacing.CARD_PADDING, Spacing.CARD_PADDING, Spacing.CARD_PADDING
        )
        layout.setSpacing(Spacing.GROUP_GAP)
        self.search = QLineEdit(self)
        self.search.setPlaceholderText(t("activity.search"))
        self.search.setAccessibleName(t("activity.search"))
        self.search.setClearButtonEnabled(True)
        layout.addWidget(self.search)
        filters = QHBoxLayout()
        filters.setSpacing(Spacing.CONTROL_GAP)
        self.status_filter = ScrollSafeComboBox(self)
        self.status_filter.setAccessibleName(t("activity.status"))
        self.status_filter.addItem(t("activity.all_statuses"), "")
        for status in ("failed", "warning", "completed", "processing", "pending", "skipped", "cancelled", "info"):
            label = activity_status_label(status)
            self.status_filter.addItem(label, status)
        self.operation_filter = ScrollSafeComboBox(self)
        self.operation_filter.setAccessibleName(t("activity.operation"))
        filters.addWidget(self.status_filter, 1)
        filters.addWidget(self.operation_filter, 1)
        layout.addLayout(filters)
        splitter = QSplitter(Qt.Orientation.Vertical, self)
        self.table = QTableView(splitter)
        self.table.setAccessibleName(t("activity.title"))
        self.table.setModel(self._proxy)
        self.table.setItemDelegate(_ActivityDelegate(self.table))
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSortingEnabled(True)
        self.table.sortByColumn(0, Qt.SortOrder.DescendingOrder)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().hide()
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 175)
        self.table.setColumnWidth(1, 95)
        self.table.setColumnWidth(2, 210)
        self.details = QPlainTextEdit(splitter)
        self.details.setReadOnly(True)
        self.details.setAccessibleName(t("activity.details"))
        self.details.setPlaceholderText(t("activity.select_record"))
        splitter.setSizes([300, 160])
        layout.addWidget(splitter, 1)
        self.outputs = ScrollSafeComboBox(self)
        self.outputs.setAccessibleName(t("file_locations.show_outputs"))
        self.outputs.hide()
        layout.addWidget(self.outputs)
        actions = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, self)
        actions.button(QDialogButtonBox.StandardButton.Close).setText(t("common.close"))
        selection_actions = ChoiceGroup(self, responsive=True, spacing=Spacing.CONTROL_GAP)
        self.open_source = QPushButton(t("file_locations.input"), selection_actions)
        self.open_output = QPushButton(t("file_locations.output"), selection_actions)
        self.copy = QPushButton(t("common.copy"), selection_actions)
        for button in (self.open_source, self.open_output, self.copy):
            button.setObjectName("secondaryActionButton")
            apply_theme_class(button, "secondary")
            selection_actions.content_layout.addWidget(button, 1)
        layout.addWidget(selection_actions)
        self.clear = actions.addButton(t("activity.clear_finished"), QDialogButtonBox.ButtonRole.ResetRole)
        for button in actions.buttons():
            button.setObjectName("secondaryActionButton")
            apply_theme_class(button, "secondary")
        layout.addWidget(actions)
        actions.rejected.connect(self.close)
        self.open_source.clicked.connect(lambda: self._open_selected(False))
        self.open_output.clicked.connect(lambda: self._open_selected(True))
        self.copy.clicked.connect(lambda: QApplication.clipboard().setText(self.details.toPlainText()))
        self.clear.clicked.connect(model.clear_finished)
        self.search.textChanged.connect(self._change_filters)
        self.status_filter.currentIndexChanged.connect(self._change_filters)
        self.operation_filter.currentIndexChanged.connect(self._change_filters)
        self.table.selectionModel().currentChanged.connect(self._show_selected)
        model.modelAboutToBeReset.connect(self._before_reset)
        model.records_changed.connect(self._refresh)
        self._refresh()

    def _before_reset(self) -> None:
        self._resetting = True

    def _refresh(self) -> None:
        selected_operation = self.operation_filter.currentData() or ""
        self.operation_filter.blockSignals(True)
        self.operation_filter.clear()
        self.operation_filter.addItem(t("activity.all_operations"), "")
        for operation in sorted({row.operation for row in self._model.records}):
            self.operation_filter.addItem(operation, operation)
        self.operation_filter.setCurrentIndex(max(0, self.operation_filter.findData(selected_operation)))
        self.operation_filter.blockSignals(False)
        self._apply_filters()
        self._resetting = False
        self._show_selected()
        self.table.resizeColumnToContents(0)
        self.table.resizeColumnToContents(1)

    def _change_filters(self, *_args) -> None:
        # An explicit filter edit returns to the full history scope.
        self._proxy.operation_id = ""
        self._proxy.source_path = ""
        self._apply_filters()

    def _apply_filters(self, *_args) -> None:
        self._proxy.query = self.search.text().strip().casefold()
        self._proxy.status = self.status_filter.currentData() or ""
        self._proxy.operation = self.operation_filter.currentData() or ""
        self._proxy.invalidate()
        for index in range(self._proxy.rowCount()):
            item = self._proxy.index(index, 0)
            record = item.data(Qt.ItemDataRole.UserRole)
            if isinstance(record, ActivityRecord) and record.key == self._selected_key:
                self.table.setCurrentIndex(item)
                self._show_selected()
                return
        if self._proxy.rowCount():
            self.table.setCurrentIndex(self._proxy.index(0, 0))
        self._show_selected()

    def _selected(self) -> ActivityRecord | None:
        value = self.table.currentIndex().data(Qt.ItemDataRole.UserRole)
        return value if isinstance(value, ActivityRecord) else None

    def _show_selected(self, *_args) -> None:
        if self._resetting:
            return
        record = self._selected()
        targeted = bool(self._proxy.operation_id)
        target_exists = any(
            row.operation_id == self._proxy.operation_id
            and (not self._proxy.source_path or row.source_path == self._proxy.source_path)
            for row in self._model.records
        )
        self.details.setPlaceholderText(
            t("activity.record_unavailable") if targeted and not target_exists else t("activity.select_record")
        )
        if record:
            self._selected_key = record.key
        text = record.details if record else ""
        if self.details.toPlainText() != text:
            self.details.setPlainText(text)
        self.details.setProperty("activityStatus", record.status if record else "")
        self.copy.setEnabled(bool(text))
        self.open_source.setEnabled(bool(record and record.source_path))
        self.open_output.setEnabled(bool(record and record.output_path))
        selected_output = self.outputs.currentData()
        self.outputs.blockSignals(True)
        self.outputs.clear()
        if record:
            for path in record.output_paths or ((record.output_path,) if record.output_path else ()):
                self.outputs.addItem(Path(path).name, path)
        self.outputs.setCurrentIndex(max(0, self.outputs.findData(selected_output)))
        self.outputs.blockSignals(False)
        self.outputs.setVisible(self.outputs.count() > 1)

    def _open_selected(self, output: bool) -> None:
        record = self._selected()
        if record:
            path = str(self.outputs.currentData() or record.output_path) if output else record.source_path
            if path:
                self.location_requested.emit(path, True)

    def show_records(self, *, failures_only: bool = False, operation_id: str = "", source_path: str = "") -> None:
        if failures_only or operation_id:
            self.search.clear()
            self.operation_filter.setCurrentIndex(0)
            self.status_filter.setCurrentIndex(self.status_filter.findData("failed") if failures_only else 0)
            self._selected_key = next(
                (
                    row.key
                    for row in self._model.records
                    if row.operation_id == operation_id and (not source_path or row.source_path == source_path)
                ),
                "",
            )
        self._proxy.operation_id = operation_id
        self._proxy.source_path = source_path if operation_id else ""
        self._apply_filters()
        self.show()
        self.raise_()
        self.activateWindow()
