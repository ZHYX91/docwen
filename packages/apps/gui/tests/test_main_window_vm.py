"""Model-state tests for MainWindowViewModel.

These tests validate that the ViewModel is the source of truth for
observable state and that signals fire correctly.  No QApplication
is needed for the ViewModel itself.
"""

from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

from docwen_core.detection import (
    OOXML_SIGNATURE_INFO_METADATA_KEY,
    OOXML_SIGNATURE_VALIDATION_UNAVAILABLE,
)
from docwen_core.models import FILE_INSPECTION_METADATA_KEY, AdmissionDecision, FileInspection
from docwen_gui.i18n import t
from docwen_gui.view_models.main_window_vm import DEFAULT_MODE, MainWindowViewModel

pytestmark = pytest.mark.unit


# ── Helpers ────────────────────────────────────────────────────────────


def _test_file_inspector(path: str) -> FileInspection:
    """Use Core for real fixtures and an explicit fake for state-only paths."""

    import docwen_core.detection as detection

    source = Path(path)
    if source.is_file():
        return detection.inspect_file(path)
    suffix = source.suffix.lower().lstrip(".")
    detected_format = "markdown" if suffix in {"md", "markdown"} else suffix or "txt"
    category = {
        "pdf": "layout",
        "xlsx": "spreadsheet",
        "xls": "spreadsheet",
        "png": "image",
        "jpg": "image",
        "jpeg": "image",
        "md": "markdown",
        "markdown": "markdown",
        "txt": "markdown",
    }.get(detected_format, "document")
    payload = {
        "file_path": str(source),
        "size_bytes": 0,
        "mtime_ns": 0,
        "extension": source.suffix.lower(),
        "declared_format": detected_format,
        "declared_category": category,
        "detected_format": detected_format,
        "detected_category": category,
        "workflow_category": category,
        "detection_method": "unknown",
        "confidence": "unverified",
        "structure_status": "unverified",
        "relation": "exact_match",
        "decision": "allow",
        "declared_supported": True,
        "detected_supported": True,
        "warning_code": "",
        "warning_message": "",
        "reason_code": "",
        "reason_message": "",
        "warnings": [],
        "ooxml_signature": {},
    }
    return FileInspection.from_dict(payload)


@pytest.fixture
def vm() -> MainWindowViewModel:
    """Create a ViewModel with no controller (limited mode)."""
    return MainWindowViewModel(controller=None, file_inspector=_test_file_inspector)


# ── Initial state ──────────────────────────────────────────────────────


class TestInitialState:
    def test_default_mode_is_single(self, vm: MainWindowViewModel) -> None:
        assert vm.mode == "single"
        assert vm.mode == DEFAULT_MODE

    def test_no_files_on_start(self, vm: MainWindowViewModel) -> None:
        assert vm.files == []
        assert vm.has_files is False

    def test_default_status_message(self, vm: MainWindowViewModel) -> None:
        assert vm.status_message == t("common.ready")

    def test_no_controller_by_default(self, vm: MainWindowViewModel) -> None:
        assert vm.controller is None


# ── Mode changes ───────────────────────────────────────────────────────


class TestModeChanges:
    def test_set_mode_valid(self, vm: MainWindowViewModel) -> None:
        signal_values: list[str] = []
        vm.mode_changed.connect(signal_values.append)
        vm.set_mode("batch")
        assert vm.mode == "batch"
        assert signal_values == ["batch"]

    def test_set_mode_invalid_raises(self, vm: MainWindowViewModel) -> None:
        with pytest.raises(ValueError, match="Invalid mode"):
            vm.set_mode("invalid_mode")

    def test_set_same_mode_no_signal(self, vm: MainWindowViewModel) -> None:
        signal_values: list[str] = []
        vm.mode_changed.connect(signal_values.append)
        vm.set_mode("single")  # already default
        assert signal_values == []

    def test_mode_changed_emits_after_releasing_mutex(self, vm: MainWindowViewModel) -> None:
        lock_available: list[bool] = []

        def probe_lock(_mode: str) -> None:
            acquired = vm._mutex.tryLock()
            lock_available.append(acquired)
            if acquired:
                vm._mutex.unlock()

        vm.mode_changed.connect(probe_lock)

        vm.set_mode("batch")

        assert lock_available == [True]


# ── File management ────────────────────────────────────────────────────


class TestFileManagement:
    def test_add_files_emits_signal(self, vm: MainWindowViewModel) -> None:
        signals: list[list] = []
        vm.files_changed.connect(signals.append)
        vm.add_files(["/tmp/a.docx", "/tmp/b.pdf"])
        assert len(signals) == 1
        assert len(signals[0]) == 2

    def test_files_changed_emits_after_releasing_mutex(self, vm: MainWindowViewModel) -> None:
        lock_available: list[bool] = []

        def probe_lock(_files: list) -> None:
            acquired = vm._mutex.tryLock()
            lock_available.append(acquired)
            if acquired:
                vm._mutex.unlock()

        vm.files_changed.connect(probe_lock)

        vm.add_files(["/tmp/a.docx"])

        assert lock_available == [True]

    def test_clear_files_emits_after_releasing_mutex(self, vm: MainWindowViewModel) -> None:
        lock_available: list[bool] = []

        def probe_lock() -> None:
            acquired = vm._mutex.tryLock()
            lock_available.append(acquired)
            if acquired:
                vm._mutex.unlock()

        vm.add_files(["/tmp/a.docx"])
        vm.files_cleared.connect(probe_lock)

        vm.clear_files()

        assert lock_available == [True]

    def test_add_duplicate_path_is_deduped(self, vm: MainWindowViewModel) -> None:
        signals: list[list] = []
        vm.files_changed.connect(signals.append)
        vm.add_files(["/tmp/a.docx"])
        assert len(signals) == 1
        vm.add_files(["/tmp/a.docx", "/tmp/b.pdf"])
        assert len(signals) == 2
        assert len(signals[1]) == 2  # only b.pdf added, a.docx skipped

    def test_remove_file_emits_signal(self, vm: MainWindowViewModel) -> None:
        signals: list[list] = []
        vm.files_changed.connect(signals.append)
        vm.add_files(["/tmp/a.docx", "/tmp/b.pdf"])
        assert len(signals) == 1
        vm.remove_file("/tmp/a.docx")
        assert len(signals) == 2
        assert len(signals[1]) == 1
        assert signals[1][0].path == "/tmp/b.pdf"

    def test_remove_nonexistent_no_signal(self, vm: MainWindowViewModel) -> None:
        signals: list[list] = []
        vm.files_changed.connect(signals.append)
        vm.add_files(["/tmp/a.docx"])
        assert len(signals) == 1
        vm.remove_file("/tmp/nonexistent.docx")
        assert len(signals) == 1  # no new signal

    def test_clear_files(self, vm: MainWindowViewModel) -> None:
        change_signals: list[list] = []
        clear_signals: list = []
        vm.files_changed.connect(change_signals.append)
        vm.files_cleared.connect(lambda: clear_signals.append(True))
        vm.add_files(["/tmp/a.docx"])
        vm.clear_files()
        assert vm.files == []
        assert vm.has_files is False
        assert len(clear_signals) == 1

    def test_add_files_is_not_limited_by_preview_scan_limit(self, vm: MainWindowViewModel) -> None:
        paths = [f"/tmp/file_{i}.docx" for i in range(250)]
        vm.add_files(paths)
        assert len(vm.files) == 250

    def test_invalid_paths_skipped(self, vm: MainWindowViewModel) -> None:
        """FileRef creation may fail for empty/nonsensical paths; ViewModel skips them."""
        vm.add_files([""])
        assert vm.files == []

    def test_production_inspector_rejects_missing_path(self, tmp_path) -> None:
        vm = MainWindowViewModel(controller=None)

        outcome = vm.add_files([str(tmp_path / "missing.docx")])

        assert outcome.added == ()
        assert len(outcome.rejected) == 1
        assert vm.files == []

    def test_structurally_invalid_office_container_is_rejected_before_list_entry(
        self, vm: MainWindowViewModel, tmp_path
    ) -> None:
        source = tmp_path / "ordinary.docx"
        with ZipFile(source, "w", ZIP_DEFLATED) as archive:
            archive.writestr("hello.txt", "not an OOXML package")

        outcome = vm.add_files([str(source)])

        assert outcome.added == ()
        assert len(outcome.rejected) == 1
        assert vm.files == []

    def test_signed_ooxml_admission_preserves_typed_warning_and_frozen_fact(
        self,
        vm: MainWindowViewModel,
        tmp_path,
        monkeypatch,
    ) -> None:
        source = tmp_path / "signed.docx"
        source.write_bytes(b"probe")
        warning = (
            f"[{OOXML_SIGNATURE_VALIDATION_UNAVAILABLE}] OOXML signature material "
            "was detected; intact and tampered inputs cannot be distinguished."
        )

        inspection_payload = {
            "actual_format": "docx",
            "detected_format": "docx",
            "actual_category": "document",
            "detected_category": "document",
            "workflow_category": "document",
            "warning_message": warning,
            "ooxml_signature": {
                "state": "complete",
                "signature_part_count": 1,
                "marker_count": 8,
                "reason": "complete structural signature graph detected",
                "format": "docx",
            },
        }
        inspection = SimpleNamespace(
            detected_format="docx",
            workflow_category="document",
            warning_message=warning,
            decision=AdmissionDecision.ALLOW_WITH_WARNING,
            reason_message="",
            ooxml_signature=inspection_payload["ooxml_signature"],
            to_dict=lambda: dict(inspection_payload),
        )

        monkeypatch.setattr("docwen_core.detection.inspect_file", lambda _path: inspection)

        vm.add_files([str(source)])

        assert len(vm.files) == 1
        assert vm.files[0].warning_message == warning
        assert vm.files[0].metadata[OOXML_SIGNATURE_INFO_METADATA_KEY]["state"] == "complete"
        assert vm.files[0].metadata[FILE_INSPECTION_METADATA_KEY]["warning_message"] == warning

    @pytest.mark.parametrize(
        ("suffix", "actual_format"),
        [
            (".txt", "txt"),
            (".md", "markdown"),
            (".txt", "markdown"),
            (".md", "txt"),
            (".markdown", "txt"),
        ],
    )
    def test_detected_txt_and_markdown_share_text_generation_workflow(
        self,
        vm: MainWindowViewModel,
        tmp_path,
        monkeypatch,
        suffix: str,
        actual_format: str,
    ) -> None:
        source = tmp_path / f"note{suffix}"
        source.write_text("plain or markdown text", encoding="utf-8")
        workflow_category = "markdown"
        payload = {
            "detected_format": actual_format,
            "detected_category": "document" if actual_format == "txt" else "markdown",
            "workflow_category": workflow_category,
            "warning_message": "",
        }
        inspection = SimpleNamespace(
            detected_format=actual_format,
            workflow_category=workflow_category,
            warning_message="",
            decision=AdmissionDecision.ALLOW,
            reason_message="",
            ooxml_signature={},
            to_dict=lambda: dict(payload),
        )
        monkeypatch.setattr("docwen_core.detection.inspect_file", lambda _path: inspection)

        vm.add_files([str(source)])
        ref = vm.files[0]
        vm.set_selected_file(ref)

        assert ref.format == actual_format
        assert ref.category == "markdown"
        assert vm.ui_projection.right_panel_slot.value == "template"


# ── Status message ─────────────────────────────────────────────────────


class TestStatusMessage:
    def test_set_status_emits_signal(self, vm: MainWindowViewModel) -> None:
        signals: list[str] = []
        vm.status_message_changed.connect(signals.append)
        vm.set_status_message("Processing...")
        assert vm.status_message == "Processing..."
        assert signals == ["Processing..."]

    def test_set_same_status_no_signal(self, vm: MainWindowViewModel) -> None:
        signals: list[str] = []
        vm.status_message_changed.connect(signals.append)
        vm.set_status_message(t("common.ready"))  # same as default
        assert signals == []

    def test_status_changed_emits_after_releasing_mutex(self, vm: MainWindowViewModel) -> None:
        lock_available: list[bool] = []

        def probe_lock(_message: str) -> None:
            acquired = vm._mutex.tryLock()
            lock_available.append(acquired)
            if acquired:
                vm._mutex.unlock()

        vm.status_message_changed.connect(probe_lock)

        vm.set_status_message("Processing...")

        assert lock_available == [True]


# ── Task event handling ────────────────────────────────────────────────


class TestTaskEventHandling:
    def test_task_started(self, vm: MainWindowViewModel) -> None:
        vm.begin_execution_telemetry("t1", ("t1",))
        vm.on_task_event("task_started", {"task_id": "t1", "message": "Converting"})
        assert vm.current_task_id == "t1"
        assert vm.status_message.startswith(t("main_window.task_processing_prefix"))

    def test_task_progress(self, vm: MainWindowViewModel) -> None:
        vm.begin_execution_telemetry("t1", ("t1",))
        vm.on_task_event("task_started", {"task_id": "t1", "message": "Converting"})
        vm.on_task_event("task_progress", {"task_id": "t1", "percent": 75.0, "message": "Table 3/4"})
        assert vm.status_message.startswith(t("main_window.task_progress_prefix"))
        assert "75%" in vm.status_message

    @pytest.mark.parametrize("event_type", ["task_completed", "task_failed", "task_cancelled"])
    def test_runtime_terminal_only_closes_live_telemetry(
        self,
        vm: MainWindowViewModel,
        event_type: str,
    ) -> None:
        summaries: list[dict] = []
        vm.task_summary_changed.connect(summaries.append)
        vm.begin_execution_telemetry("t1", ("t1",))
        vm.on_task_event("task_started", {"task_id": "t1", "message": "working"})
        live_status = vm.status_message
        vm.on_task_event(event_type, {"task_id": "t1", "message": "terminal"})
        assert vm.current_task_id is None
        assert vm.status_message == live_status
        assert summaries == []

    def test_late_runtime_identity_cannot_replace_or_clear_current_execution(self, vm: MainWindowViewModel) -> None:
        vm.begin_execution_telemetry("operation-a", ("operation-a",))
        vm.on_task_event("task_started", {"task_id": "operation-a", "message": "old"})
        vm.begin_execution_telemetry("operation-b", ("operation-b",))
        vm.on_task_event("task_started", {"task_id": "operation-b", "message": "current"})
        current_status = vm.status_message

        vm.on_task_event("task_completed", {"task_id": "operation-a"})
        vm.on_task_event("task_started", {"task_id": "operation-a", "message": "stale"})

        assert vm.current_task_id == "operation-b"
        assert vm.status_message == current_status

    def test_terminal_child_cannot_clear_newer_active_batch_child(self, vm: MainWindowViewModel) -> None:
        vm.begin_execution_telemetry("batch", ("batch-0", "batch-1"))
        vm.on_task_event("task_started", {"task_id": "batch-0", "message": "first"})
        vm.on_task_event("task_completed", {"task_id": "batch-0"})
        vm.on_task_event("task_started", {"task_id": "batch-1", "message": "second"})
        current_status = vm.status_message

        vm.on_task_event("task_failed", {"task_id": "batch-0", "message": "late"})
        vm.on_task_event("task_started", {"task_id": "batch-0", "message": "late start"})

        assert vm.current_task_id == "batch-1"
        assert vm.status_message == current_status

    @pytest.mark.parametrize(
        ("status", "expected_status"),
        [
            ("completed", "completed"),
            ("partial", "completed"),
            ("failed", "failed"),
            ("cancelled", "cancelled"),
        ],
    )
    def test_authoritative_execution_summary_publishes_once(
        self,
        vm: MainWindowViewModel,
        status: str,
        expected_status: str,
    ) -> None:
        summaries: list[dict] = []
        vm.task_summary_changed.connect(summaries.append)
        vm.publish_execution_summary(status, {"task_id": "t1", "message": "boom"})

        if expected_status == "completed":
            assert vm.status_message == t("main_window.task_completed_status")
        elif expected_status == "failed":
            assert vm.status_message.startswith(t("main_window.task_failed_prefix"))
        else:
            assert vm.status_message == t("main_window.task_cancelled_status")
        assert summaries == [{"task_id": "t1", "message": "boom", "status": status}]

    def test_unknown_event_type_ignored(self, vm: MainWindowViewModel) -> None:
        """Unknown event types should not crash the VM."""
        vm.on_task_event("unknown_event", {"task_id": "t1"})
        # should not raise, should not change status
        assert vm.status_message == t("common.ready")

    def test_ipc_missing_file_status_is_localized(self, vm: MainWindowViewModel, tmp_path) -> None:
        missing_path = tmp_path / "missing.docx"
        vm.handle_ipc_command("add_file", str(missing_path))

        assert vm.status_message == t("info_area.ipc_file_missing", path=str(missing_path))

    def test_ipc_add_publishes_admitted_file_before_activation(self, vm: MainWindowViewModel, tmp_path) -> None:
        source = tmp_path / "renamed.docx"
        source.write_bytes(b"%PDF-1.4\n% IPC warning probe\n")
        events: list[str] = []
        received_warnings: list[str] = []
        vm.files_changed.connect(lambda _refs: events.append("files"))
        vm.ipc_file_received.connect(
            lambda _path: (
                events.append("ipc"),
                received_warnings.append(vm.files[-1].warning_message),
            )
        )
        vm.window_activation_requested.connect(lambda: events.append("activate"))

        vm.handle_ipc_command("add_file", str(source))

        assert events == ["files", "ipc", "activate"]
        assert received_warnings and received_warnings[0]


# ── Shutdown ───────────────────────────────────────────────────────────


class TestShutdown:
    def test_request_shutdown_emits_signal(self, vm: MainWindowViewModel) -> None:
        signals: list = []
        vm.shutdown_requested.connect(lambda: signals.append(True))
        vm.request_shutdown()
        assert len(signals) == 1

    def test_request_shutdown_leaves_controller_stop_to_window_owner(self) -> None:
        ctrl = MagicMock()
        vm = MainWindowViewModel(controller=ctrl)
        vm.request_shutdown()
        ctrl.stop.assert_not_called()


# ── Title updates ──────────────────────────────────────────────────────


class TestTitleUpdates:
    def test_title_with_files(self, vm: MainWindowViewModel) -> None:
        titles: list[str] = []
        vm.title_changed.connect(titles.append)
        vm.add_files(["/tmp/a.docx"])
        assert len(titles) >= 1
        assert "[1 file]" in titles[-1]

    def test_title_after_clear(self, vm: MainWindowViewModel) -> None:
        titles: list[str] = []
        vm.title_changed.connect(titles.append)
        vm.add_files(["/tmp/a.docx"])
        vm.clear_files()
        assert "DocWen" in titles[-1]
        assert len(titles[-1]) > 0
