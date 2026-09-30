"""GUI clipboard Markdown input contracts: paste, output, retry, and safe presentation."""

from __future__ import annotations

from pathlib import Path

import pytest
from PySide6.QtCore import QMimeData, Qt, QUrl
from PySide6.QtWidgets import QApplication

from docwen_core.models import FILE_INSPECTION_METADATA_KEY
from docwen_core.models.request import OutputPolicy
from docwen_gui.dialogs.activity_records import ActivityRecordsDialog
from docwen_gui.main_window import MainWindow
from docwen_gui.path_identity import normalize_path
from docwen_gui.view_models.activity_records import ActivityRecordsModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.gui


@pytest.fixture
def clipboard_window(qapp: QApplication, qtbot, tmp_path: Path):
    window = MainWindow(
        view_model=MainWindowViewModel(controller=None),
        clipboard_input_root=tmp_path / "clipboard-inputs",
    )
    qtbot.addWidget(window)
    window.resize(720, 620)
    window.show()
    qapp.processEvents()
    yield window
    window.close()


def _paste_text(window: MainWindow, qapp: QApplication, qtbot, text: str) -> str:
    qapp.clipboard().setText(text)
    qtbot.mouseClick(window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: bool(window.view_model.files))
    selected = window.view_model.selected_file
    assert selected is not None
    return selected.path


def test_single_paste_freezes_exact_utf8_and_replaces_only_after_admission(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    first_text = "  ---\ntitle: 保留\n---\n\n```text\n  keep  \n```\n"
    first_path = _paste_text(clipboard_window, qapp, qtbot, first_text)
    first = Path(first_path)
    assert clipboard_window._clipboard_store is not None
    descriptor = clipboard_window._clipboard_store.descriptor(first_path)
    assert descriptor is not None
    assert first.read_bytes() == first_text.encode("utf-8")
    assert descriptor.display_name in clipboard_window._input_area_vm.selection_message
    assert first_path not in clipboard_window._input_area_vm.selection_message
    assert clipboard_window.input_area._open_location_button.isHidden()

    qapp.clipboard().setText("new clipboard value")
    assert first.read_bytes() == first_text.encode("utf-8")

    second_text = "Ordinary unmarked text is still Markdown workflow input."
    second_path = _paste_text(clipboard_window, qapp, qtbot, second_text)
    assert second_path != first_path
    assert len(clipboard_window.view_model.files) == 1
    assert Path(second_path).read_bytes() == second_text.encode("utf-8")
    assert not first.exists()


def test_batch_paste_appends_independent_visible_snapshots(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    clipboard_window._input_area_vm.set_mode("batch")
    qapp.processEvents()

    for text in ("# First\n", "# Second\n"):
        qapp.clipboard().setText(text)
        qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
        qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)

    qtbot.waitUntil(lambda: len(clipboard_window.view_model.files) == 2)
    refs = clipboard_window.view_model.files
    entries = [clipboard_window._batch_list_vm.get_file_entry(ref.path) for ref in refs]
    assert clipboard_window._clipboard_store is not None
    for entry, ref in zip(entries, refs, strict=True):
        assert entry is not None
        descriptor = clipboard_window._clipboard_store.descriptor(ref.path)
        assert descriptor is not None
        assert entry.file_name == descriptor.display_name
        assert entry.source_location_available is False
    assert [Path(ref.path).read_text(encoding="utf-8") for ref in refs] == ["# First\n", "# Second\n"]


def test_file_clipboard_prefers_real_file_over_url_text(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    tmp_path: Path,
) -> None:
    source = tmp_path / "copied.md"
    source.write_text("# Copied file\n", encoding="utf-8")
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(source))])
    mime.setText(f"file:///{source.as_posix()}")
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(clipboard_window.view_model.files) == 1)

    selected = clipboard_window.view_model.selected_file
    assert selected is not None
    assert Path(selected.path).resolve() == source.resolve()
    assert clipboard_window._clipboard_store is None


def test_single_file_clipboard_rejects_multiple_and_preserves_current_input(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    tmp_path: Path,
) -> None:
    original = _paste_text(clipboard_window, qapp, qtbot, "# Keep current\n")
    first = tmp_path / "first.md"
    second = tmp_path / "second.md"
    first.write_text("# First\n", encoding="utf-8")
    second.write_text("# Second\n", encoding="utf-8")
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(first)), QUrl.fromLocalFile(str(second))])
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qapp.processEvents()

    assert [ref.path for ref in clipboard_window.view_model.files] == [original]
    assert clipboard_window.view_model.selected_file is not None
    assert clipboard_window.view_model.selected_file.path == original
    assert clipboard_window._input_area_vm.selection_tone == "warning"


def test_single_file_clipboard_counts_original_list_before_filtering(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original = _paste_text(clipboard_window, qapp, qtbot, "# Keep current\n")
    valid = tmp_path / "valid.md"
    missing = tmp_path / "missing.bin"
    valid.write_text("# Valid\n", encoding="utf-8")
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(valid)), QUrl.fromLocalFile(str(missing))])
    qapp.clipboard().setMimeData(mime)
    monkeypatch.setattr("docwen_gui.dialogs.feedback.confirm", lambda *_a, **_k: False)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qapp.processEvents()
    assert [ref.path for ref in clipboard_window.view_model.files] == [original]
    assert clipboard_window._input_area_vm.selection_tone == "warning"


def test_single_folder_clipboard_keeps_input_when_batch_switch_declined(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original = _paste_text(clipboard_window, qapp, qtbot, "# Keep current\n")
    folder = tmp_path / "folder"
    folder.mkdir()
    (folder / "inside.md").write_text("# Inside\n", encoding="utf-8")
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(folder))])
    qapp.clipboard().setMimeData(mime)
    monkeypatch.setattr("docwen_gui.dialogs.feedback.confirm", lambda *_a, **_k: False)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qapp.processEvents()
    assert clipboard_window._input_area_vm.mode == "single"
    assert [ref.path for ref in clipboard_window.view_model.files] == [original]


def test_switch_to_batch_add_uses_captured_clipboard_file_list(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original = _paste_text(clipboard_window, qapp, qtbot, "# Keep current\n")
    first = tmp_path / "first.md"
    second = tmp_path / "second.csv"
    first.write_text("# First\n", encoding="utf-8")
    second.write_text("name,value\nA,1\n", encoding="utf-8")
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(first)), QUrl.fromLocalFile(str(second))])
    qapp.clipboard().setMimeData(mime)

    def accept_and_change_clipboard(*_args, **_kwargs):
        qapp.clipboard().setText("# changed after capture\n")
        return True
    monkeypatch.setattr("docwen_gui.dialogs.feedback.confirm", accept_and_change_clipboard)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(clipboard_window.view_model.files) == 3)
    assert clipboard_window._input_area_vm.mode == "batch"
    assert {Path(ref.path).resolve() for ref in clipboard_window.view_model.files} == {
        Path(original).resolve(),
        first.resolve(),
        second.resolve(),
    }


def test_batch_file_clipboard_adds_mixed_real_files(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    tmp_path: Path,
) -> None:
    clipboard_window._input_area_vm.set_mode("batch")
    markdown = tmp_path / "notes.md"
    spreadsheet = tmp_path / "data.csv"
    markdown.write_text("# Notes\n", encoding="utf-8")
    spreadsheet.write_text("name,value\nA,1\n", encoding="utf-8")
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(markdown)), QUrl.fromLocalFile(str(spreadsheet))])
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(clipboard_window.view_model.files) == 2)

    refs = clipboard_window.view_model.files
    assert {Path(ref.path).resolve() for ref in refs} == {markdown.resolve(), spreadsheet.resolve()}
    assert {ref.category for ref in refs} == {"markdown", "spreadsheet"}
    assert clipboard_window._clipboard_store is None


def test_batch_file_clipboard_deduplicates_and_reports_partial_unavailable_inputs(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    tmp_path: Path,
) -> None:
    clipboard_window._input_area_vm.set_mode("batch")
    valid = tmp_path / "valid.md"
    blocked = tmp_path / "blocked.bin"
    missing = tmp_path / "missing.pdf"
    valid.write_text("# Valid\n", encoding="utf-8")
    blocked.write_bytes(b"\x00\x01\x02\x03")
    mime = QMimeData()
    mime.setUrls(
        [
            QUrl.fromLocalFile(str(valid)),
            QUrl.fromLocalFile(str(valid)),
            QUrl.fromLocalFile(str(blocked)),
            QUrl.fromLocalFile(str(missing)),
        ]
    )
    qapp.clipboard().setMimeData(mime)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(clipboard_window.view_model.files) == 1)
    assert Path(clipboard_window.view_model.files[0].path).resolve() == valid.resolve()
    assert "2" in clipboard_window._input_area_vm.selection_message
    detail = clipboard_window._input_area_vm.selection_detail
    assert str(missing) in detail
    assert "blocked" in detail.lower() or "support" in detail.lower() or "内容" in detail


def test_nonlocal_url_clipboard_still_uses_text_fallback(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    text = "https://example.invalid/report"
    mime = QMimeData()
    mime.setUrls([QUrl(text)])
    mime.setText(text)
    qapp.clipboard().setMimeData(mime)

    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(clipboard_window.view_model.files) == 1)

    selected = clipboard_window.view_model.selected_file
    assert selected is not None
    assert Path(selected.path).read_text(encoding="utf-8") == text
    assert clipboard_window._clipboard_store is not None


def test_empty_or_non_text_clipboard_never_creates_input(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    qapp.clipboard().setText(" \t\r\n ")
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qapp.processEvents()
    assert clipboard_window.view_model.files == []

    mime = QMimeData()
    mime.setData("application/octet-stream", b"binary")
    qapp.clipboard().setMimeData(mime)
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qapp.processEvents()
    assert clipboard_window.view_model.files == []
    assert clipboard_window._clipboard_store is None


def test_source_output_policy_cancel_custom_and_mixed_batch(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    synthetic = _paste_text(clipboard_window, qapp, qtbot, "# Clipboard\n")
    regular = tmp_path / "regular.md"
    regular.write_text("# Regular\n", encoding="utf-8")

    monkeypatch.setattr("docwen_gui.main_window.QFileDialog.getExistingDirectory", lambda *_a, **_k: "")
    assert clipboard_window._prepare_clipboard_output_policy([synthetic], "single", OutputPolicy()) is None

    custom = tmp_path / "custom"
    custom_policy = OutputPolicy(output_dir=str(custom))
    monkeypatch.setattr(
        "docwen_gui.main_window.QFileDialog.getExistingDirectory",
        lambda *_a, **_k: pytest.fail("custom output must not ask again"),
    )
    assert clipboard_window._prepare_clipboard_output_policy([synthetic], "single", custom_policy) is custom_policy
    for unsafe in (
        OutputPolicy(output_dir=str(Path(synthetic).parent)),
        OutputPolicy(output_path=str(Path(synthetic).parent / "result.docx")),
    ):
        assert clipboard_window._prepare_clipboard_output_policy([synthetic], "single", unsafe) is None

    persistent = tmp_path / "persistent"
    monkeypatch.setattr(
        "docwen_gui.main_window.QFileDialog.getExistingDirectory",
        lambda *_a, **_k: str(persistent),
    )
    mixed = clipboard_window._prepare_clipboard_output_policy(
        [str(regular), synthetic],
        "batch",
        OutputPolicy(),
    )
    assert mixed is not None
    assert mixed.output_dir is None
    assert mixed.for_input(str(regular)).output_dir is None
    assert mixed.for_input(synthetic).output_dir == str(persistent.resolve())


def test_unwritable_snapshot_root_rejects_paste_without_changing_selection(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original = _paste_text(clipboard_window, qapp, qtbot, "# Retained input\n")

    def unavailable_store():
        raise PermissionError("test-only denied root")

    monkeypatch.setattr(clipboard_window, "_clipboard_store_for_paste", unavailable_store)
    qapp.clipboard().setText("# Replacement\n")
    qtbot.mouseClick(clipboard_window.input_area.paste_button, Qt.MouseButton.LeftButton)
    qapp.processEvents()
    assert [ref.path for ref in clipboard_window.view_model.files] == [original]
    assert Path(original).read_text(encoding="utf-8") == "# Retained input\n"


def test_failed_retry_reuses_original_snapshot_not_current_clipboard(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original = "# Original clipboard snapshot\n"
    path = _paste_text(clipboard_window, qapp, qtbot, original)
    normalized = normalize_path(path)
    assert clipboard_window._clipboard_store is not None
    descriptor = clipboard_window._clipboard_store.descriptor(path)
    assert descriptor is not None

    context = {
        "request_id": "clipboard-retry",
        "file_path": normalized,
        "target_format": "docx",
        "action_name": "",
        "options": {},
        "synthetic_input_paths": [normalized],
        "source_labels": {normalized: descriptor.display_name},
    }
    clipboard_window._task_history.remember(context)
    clipboard_window._task_history.record("clipboard-retry", normalized, "failed")
    clipboard_window._info_area_vm.set_task_summary(operation_id="clipboard-retry", state="failed")

    clipboard_window.view_model.remove_file(path)
    qapp.processEvents()
    assert Path(path).is_file()
    qapp.clipboard().setText("# Changed clipboard\n")

    calls: list[dict[str, object]] = []
    monkeypatch.setattr(clipboard_window._workflow, "single", lambda **kwargs: calls.append(kwargs))
    clipboard_window._retry_failed_request()
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(calls) == 1)

    assert Path(path).read_text(encoding="utf-8") == original
    assert calls[0]["file_path"] == normalized


def test_missing_retry_snapshot_does_not_read_new_clipboard(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = _paste_text(clipboard_window, qapp, qtbot, "# Original\n")
    normalized = normalize_path(path)
    assert clipboard_window._clipboard_store is not None
    descriptor = clipboard_window._clipboard_store.descriptor(path)
    assert descriptor is not None
    clipboard_window._task_history.remember(
        {
            "request_id": "missing-clipboard-retry",
            "file_path": normalized,
            "target_format": "docx",
            "action_name": "",
            "options": {},
            "synthetic_input_paths": [normalized],
            "source_labels": {normalized: descriptor.display_name},
        }
    )
    clipboard_window._task_history.record("missing-clipboard-retry", normalized, "failed")
    clipboard_window._info_area_vm.set_task_summary(operation_id="missing-clipboard-retry", state="failed")
    Path(path).unlink()
    qapp.clipboard().setText("# Replacement must not be used\n")
    calls: list[dict[str, object]] = []
    monkeypatch.setattr(clipboard_window._workflow, "single", lambda **kwargs: calls.append(kwargs))

    clipboard_window._retry_failed_request()
    qapp.processEvents()

    assert calls == []


def test_running_state_locks_all_input_mutation_controls(clipboard_window: MainWindow, qapp: QApplication) -> None:
    clipboard_window._action_area_vm.show_cancel()
    qapp.processEvents()
    controls = (
        clipboard_window.input_area.add_button,
        clipboard_window.input_area.paste_button,
        clipboard_window.input_area.clear_button,
        clipboard_window.input_area.single_mode_button,
        clipboard_window.input_area.batch_mode_button,
    )
    assert all(not control.isEnabled() for control in controls)

    clipboard_window._action_area_vm.hide_cancel()
    qapp.processEvents()
    assert all(control.isEnabled() for control in controls)


def test_activity_records_show_clipboard_label_without_backing_path(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
) -> None:
    path = _paste_text(clipboard_window, qapp, qtbot, "# Private backing path must stay hidden\n")
    normalized = normalize_path(path)
    assert clipboard_window._clipboard_store is not None
    descriptor = clipboard_window._clipboard_store.descriptor(path)
    assert descriptor is not None
    clipboard_window._task_history.remember(
        {
            "request_id": "activity-clipboard",
            "file_path": normalized,
            "target_format": "docx",
            "action_name": "",
            "options": {},
            "synthetic_input_paths": [normalized],
            "source_labels": {normalized: descriptor.display_name},
        }
    )
    clipboard_window._task_history.record("activity-clipboard", normalized, "failed", "Conversion failed")

    model = ActivityRecordsModel(clipboard_window._task_history, clipboard_window._info_area_vm)
    dialog = ActivityRecordsDialog(model)
    qtbot.addWidget(dialog)
    dialog.show_records(operation_id="activity-clipboard")
    qapp.processEvents()

    assert model.records[0].source_label == descriptor.display_name
    assert normalized not in model.records[0].details
    assert dialog.table.model().index(0, 2).data() == descriptor.display_name
    assert not dialog.open_source.isEnabled()


def test_failed_html_clipboard_retry_restores_original_synthetic_inspection(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    text = "<html><body>literal clipboard body</body></html>\n"
    path = _paste_text(clipboard_window, qapp, qtbot, text)
    normalized = normalize_path(path)
    original_ref = clipboard_window.view_model.files[0]
    original_fact = dict(original_ref.metadata[FILE_INSPECTION_METADATA_KEY])
    assert original_fact["detection_method"] == "synthetic_markdown"
    assert clipboard_window._clipboard_store is not None
    descriptor = clipboard_window._clipboard_store.descriptor(path)
    assert descriptor is not None

    context = {
        "request_id": "clipboard-html-retry",
        "file_path": normalized,
        "target_format": "docx",
        "action_name": "",
        "options": {},
        "input_refs": [original_ref.to_dict()],
        "synthetic_input_paths": [normalized],
        "source_labels": {normalized: descriptor.display_name},
    }
    clipboard_window._task_history.remember(context)
    clipboard_window._task_history.record("clipboard-html-retry", normalized, "failed")
    clipboard_window._info_area_vm.set_task_summary(operation_id="clipboard-html-retry", state="failed")
    clipboard_window.view_model.remove_file(path)
    qapp.processEvents()
    assert Path(path).is_file()

    qapp.clipboard().setText("# New clipboard must not replace retry bytes\n")
    calls: list[dict[str, object]] = []
    monkeypatch.setattr(clipboard_window._workflow, "single", lambda **kwargs: calls.append(kwargs))

    clipboard_window._retry_failed_request()
    qtbot.waitUntil(lambda: not clipboard_window.view_model.inspection_busy)
    qtbot.waitUntil(lambda: len(calls) == 1)

    restored_ref = clipboard_window.view_model.files[0]
    assert restored_ref.path == normalized
    assert restored_ref.metadata[FILE_INSPECTION_METADATA_KEY] == original_fact
    assert Path(path).read_text(encoding="utf-8") == text
    assert calls[0]["file_path"] == normalized


def test_tampered_failed_clipboard_snapshot_is_rejected_before_retry(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = _paste_text(clipboard_window, qapp, qtbot, "<html>original</html>\n")
    normalized = normalize_path(path)
    original_ref = clipboard_window.view_model.files[0]
    assert clipboard_window._clipboard_store is not None
    descriptor = clipboard_window._clipboard_store.descriptor(path)
    assert descriptor is not None
    clipboard_window._task_history.remember(
        {
            "request_id": "clipboard-tamper-retry",
            "file_path": normalized,
            "target_format": "docx",
            "action_name": "",
            "options": {},
            "input_refs": [original_ref.to_dict()],
            "synthetic_input_paths": [normalized],
            "source_labels": {normalized: descriptor.display_name},
        }
    )
    clipboard_window._task_history.record("clipboard-tamper-retry", normalized, "failed")
    clipboard_window._info_area_vm.set_task_summary(operation_id="clipboard-tamper-retry", state="failed")
    clipboard_window.view_model.remove_file(path)
    qapp.processEvents()

    Path(path).write_text("<html>tampered</html>\n", encoding="utf-8")
    qapp.clipboard().setText("# Replacement clipboard\n")
    calls: list[dict[str, object]] = []
    monkeypatch.setattr(clipboard_window._workflow, "single", lambda **kwargs: calls.append(kwargs))

    clipboard_window._retry_failed_request()
    qapp.processEvents()

    assert calls == []
    assert clipboard_window.view_model.files == []


def test_main_window_source_location_receiver_rejects_synthetic_input(
    clipboard_window: MainWindow,
    qapp: QApplication,
    qtbot,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = _paste_text(clipboard_window, qapp, qtbot, "# Synthetic\n")
    calls: list[tuple[str, bool]] = []
    monkeypatch.setattr(
        clipboard_window,
        "_open_path",
        lambda target, *, open_parent=False: calls.append((target, open_parent)) or True,
    )

    clipboard_window._handle_batch_entry_action("open_source_location", path)
    clipboard_window._open_location(path)

    assert calls == []
