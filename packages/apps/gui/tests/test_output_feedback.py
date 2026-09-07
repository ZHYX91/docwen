"""Output identity and state actions remain tied to their completed task."""

import pytest
from PySide6.QtCore import Qt

from docwen_core.models.artifact import ArtifactManifest
from docwen_core.models.result import ConversionResult
from docwen_gui.dialogs.activity_records import ActivityRecordsDialog
from docwen_gui.view_models.activity_records import ActivityRecordsModel
from docwen_gui.view_models.batch_list_vm import BatchFileEntry, BatchListViewModel
from docwen_gui.view_models.info_area_vm import InfoAreaViewModel
from docwen_gui.view_models.output_files import result_output_paths
from docwen_gui.view_models.task_history import TaskHistory
from docwen_gui.widgets.batch_list import BatchEntryItemWidget, BatchList
from docwen_gui.widgets.info_area import InfoArea

pytestmark = pytest.mark.gui


def test_batch_entry_card_fits_viewport_with_list_spacing(qtbot, tmp_path):
    source = tmp_path / "项目记录.md"
    source.write_text("body", encoding="utf-8")
    vm = BatchListViewModel()
    vm.add_files([str(source)])
    view = BatchList(vm)
    qtbot.addWidget(view)
    view.resize(400, 600)
    view.show()
    qtbot.waitUntil(lambda: bool(view.findChildren(BatchEntryItemWidget)))
    entry = view.findChildren(BatchEntryItemWidget)[0]
    listing = entry._list_widget
    assert listing is not None
    qtbot.waitUntil(lambda: entry.width() > 0)
    assert entry.geometry().right() < listing.viewport().width()


def test_single_item_batch_keeps_completion_counts(qtbot):
    vm = InfoAreaViewModel()
    vm.set_mode("batch")
    vm.set_task_summary(operation_id="batch", state="success", total_count=1, completed_count=1, batch=True)
    view = InfoArea(vm)
    qtbot.addWidget(view)
    view.show()
    assert "1/1" in vm.status_summary_text


def test_result_outputs_exclude_resources_and_keep_primary_first():
    artifacts = [
        ArtifactManifest("csv2", "primary", "/out/two.csv", "two.csv", "text/csv"),
        ArtifactManifest("image", "image", "/out/a.png", "a.png", "image/png"),
        ArtifactManifest("manifest", "manifest", "/out/node.json", "node.json"),
        ArtifactManifest("csv1", "primary", "/out/one.csv", "one.csv", "text/csv", is_primary=True),
        ArtifactManifest("attachment", "auxiliary", "/out/attachment.md", "attachment.md", "text/markdown"),
    ]
    result = ConversionResult(task_id="outputs", success=True, artifacts=artifacts)
    assert result_output_paths(result) == ("/out/one.csv", "/out/two.csv", "/out/attachment.md")


@pytest.mark.parametrize("secondary_kind", ["image", "auxiliary"])
def test_page_and_frame_images_are_outputs_not_extracted_resources(secondary_kind):
    artifacts = [
        ArtifactManifest("first", "image", "/out/page1.png", "page1.png", "image/png", is_primary=True),
        ArtifactManifest("second", secondary_kind, "/out/page2.png", "page2.png", "image/png"),
        ArtifactManifest("third", secondary_kind, "/out/page3.png", "page3.png", "image/png"),
    ]
    result = ConversionResult(task_id="pages", success=True, artifacts=artifacts)
    assert result_output_paths(result) == tuple(a.staging_path for a in artifacts)


def test_single_result_location_and_count_use_completed_output(qtbot):
    vm = InfoAreaViewModel()
    view = InfoArea(vm)
    qtbot.addWidget(view)
    paths = ("/out/项目记录_人员明细.csv", "/out/项目记录_部门.csv")
    vm.set_task_summary(operation_id="finished", state="success", output_path=paths[0], output_paths=paths)
    view.show()
    row = view._output_row
    located, details = [], []
    vm.location_requested.connect(located.append)
    vm.task_guide_action_requested.connect(lambda action, target: details.append((action, target)))
    row.location_button.click()
    row.count_button.click()
    assert located == [paths[0]]
    assert details == [("view_outputs", "")]
    assert row.count_button.isVisible()
    assert "2" in row.count_button.text()
    assert not view.is_guide_row_visible
    vm.set_mode("batch")
    assert not row.isVisible()
    vm.set_mode("single")
    assert row.isVisible()
    vm.begin_task(operation_id="next", current_file="next.md", total_count=1)
    assert not row.isVisible()


@pytest.mark.parametrize(
    ("status", "output", "action"),
    [
        ("completed", "/out/result.md", "open_output_location"),
        ("failed", "", "show_error_details"),
        ("skipped", "", "show_skip_details"),
    ],
)
def test_batch_status_actions_are_keyboard_accessible(qtbot, status, output, action):
    entry = BatchFileEntry("/input.docx", "input.docx", "docx", "document", status=status, output_path=output)
    view = BatchEntryItemWidget(entry)
    qtbot.addWidget(view)
    view.show()
    observed = []
    view.action_requested.connect(lambda key, path: observed.append((key, path)))
    view.status_button.setFocus()
    qtbot.keyClick(view.status_button, Qt.Key.Key_Space)
    assert observed == [(action, entry.file_path)]
    assert not view.status_button.icon().isNull()
    if output:
        view.output_row.location_button.click()
        assert observed[-1] == ("open_output_location", entry.file_path)


def test_multiple_outputs_remain_available_after_list_changes(qtbot):
    history, feedback = TaskHistory(), InfoAreaViewModel()
    history.remember({"request_id": "op", "file_path": "/input.md"})
    outputs = ("/out/one.csv", "/out/two.csv")
    history.record("op", "/input.md", "completed", output_path=outputs[0], output_paths=outputs)
    model = ActivityRecordsModel(history, feedback)
    view = ActivityRecordsDialog(model)
    qtbot.addWidget(view)
    view.show_records(operation_id="op")
    assert view.outputs.count() == 2
    assert all(path in view.details.toPlainText() for path in outputs)
    opened = []
    view.location_requested.connect(lambda path, parent: opened.append((path, parent)))
    view.outputs.setCurrentIndex(1)
    view.open_output.click()
    assert opened == [(outputs[1], True)]


def test_retry_clears_previous_outputs_before_new_failure(qapp):
    vm = BatchListViewModel()
    path = "C:/input.docx"
    vm.add_files([path], file_resolver=lambda _: {"detected_format": "docx", "workflow_category": "document"})
    vm.set_file_status(path, "completed", output_path="C:/old.docx", output_paths=("C:/old.docx", "C:/old2.docx"))
    vm.set_file_status(path, "processing", operation_id="new")
    vm.set_file_status(path, "failed", error_message="Worker failed", operation_id="new")
    entry = vm.get_file_entry(path)
    assert entry is not None
    assert not entry.output_path and not entry.output_paths


def test_targeted_history_never_falls_back_to_an_unrelated_record(qtbot):
    history, feedback = TaskHistory(), InfoAreaViewModel()
    history.remember({"request_id": "new", "file_path": "/new.md"})
    history.record("new", "/new.md", "completed", output_path="/new.docx")
    dialog = ActivityRecordsDialog(ActivityRecordsModel(history, feedback))
    qtbot.addWidget(dialog)
    dialog.show_records(operation_id="cleared", source_path="/old.md")
    assert dialog._proxy.rowCount() == 0
    assert dialog._selected() is None
    assert not dialog.details.toPlainText()
    assert not dialog.open_output.isEnabled()
    dialog.show_records()
    assert dialog._proxy.rowCount() == 1


def test_batch_summary_has_no_implicit_output_navigation(qtbot):
    vm = InfoAreaViewModel()
    view = InfoArea(vm)
    qtbot.addWidget(view)
    vm.set_task_summary(
        state="success",
        batch=True,
        total_count=2,
        completed_count=2,
        output_path="/first.docx",
        navigate_file_path="/first.docx",
        navigation_kind="output",
    )
    assert not vm.status_action_target
    assert not view._status_summary_label.isEnabled()
    assert view._output_row.isHidden()
