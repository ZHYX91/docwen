"""Model-state tests for InputAreaViewModel.

These tests validate that the ViewModel is the source of truth for
InputArea state and that signals fire correctly.  No QApplication
is needed for the ViewModel itself.
"""

from pathlib import Path

import pytest

from docwen_core.models import FILE_INSPECTION_METADATA_KEY, AdmissionDecision, FileInspection
from docwen_core.models.file_ref import FileRef
from docwen_gui.view_models.input_area_vm import _BATCH_SCAN_LIMIT, _TEXT_PAYLOAD_MAX_PATHS, InputAreaViewModel
from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.gui


# ── Fixtures ──────────────────────────────────────────────────────────


@pytest.fixture
def main_vm() -> MainWindowViewModel:
    return MainWindowViewModel(controller=None)


@pytest.fixture
def vm(main_vm: MainWindowViewModel) -> InputAreaViewModel:
    """Create an InputAreaViewModel wired to a MainWindowViewModel."""
    return InputAreaViewModel(main_vm=main_vm)


# ── Initial state ─────────────────────────────────────────────────────


class TestInitialState:
    def test_default_mode_is_single(self, vm: InputAreaViewModel) -> None:
        assert vm.mode == "single"

    def test_no_selection_message_initially(self, vm: InputAreaViewModel) -> None:
        assert vm.selection_message == ""
        assert vm.selection_detail == ""
        assert vm.selection_tone == "secondary"

    def test_syncs_mode_from_main_vm(self, main_vm: MainWindowViewModel) -> None:
        main_vm.set_mode("batch")
        vm = InputAreaViewModel(main_vm=main_vm)
        assert vm.mode == "batch"


# ── Mode changes ──────────────────────────────────────────────────────


class TestModeChanges:
    def test_set_mode_valid(self, vm: InputAreaViewModel) -> None:
        signals: list[str] = []
        vm.mode_changed.connect(signals.append)
        vm.set_mode("batch")
        assert vm.mode == "batch"
        assert signals == ["batch"]

    def test_set_mode_invalid_raises(self, vm: InputAreaViewModel) -> None:
        with pytest.raises(ValueError, match="Invalid mode"):
            vm.set_mode("invalid")

    def test_set_same_mode_no_signal(self, vm: InputAreaViewModel) -> None:
        signals: list[str] = []
        vm.mode_changed.connect(signals.append)
        vm.set_mode("single")  # already default
        assert signals == []

    def test_set_mode_syncs_to_main_vm(self, vm: InputAreaViewModel, main_vm: MainWindowViewModel) -> None:
        vm.set_mode("batch")
        assert main_vm.mode == "batch"


# ── File addition: single mode ────────────────────────────────────────


class TestSingleModeAddFiles:
    def test_add_single_valid_file(self, vm: InputAreaViewModel, tmp_path) -> None:
        file_path = tmp_path / "test.docx"
        file_path.write_text("content")
        signals: list[list] = []
        vm.files_added.connect(signals.append)
        vm.add_files([str(file_path)])
        assert len(signals) == 1
        assert signals[0] == [str(file_path)]
        assert vm.selection_detail == str(tmp_path)

    def test_format_warning_remains_final_feedback_and_file_is_inspected_once(
        self,
        tmp_path,
    ) -> None:
        file_path = tmp_path / "renamed.docx"
        file_path.write_bytes(b"%PDF-1.4\n")
        calls: list[str] = []

        def inspect(path: str) -> FileInspection:
            calls.append(path)
            warning = "File extension (docx) does not match actual content (pdf)"
            return FileInspection.from_dict(
                {
                    "file_path": path,
                    "declared_format": "docx",
                    "declared_category": "document",
                    "detected_format": "pdf",
                    "detected_category": "layout",
                    "workflow_category": "layout",
                    "decision": AdmissionDecision.REQUIRE_EXPLICIT_ACCEPTANCE.value,
                    "warning_message": warning,
                    "reason_message": warning,
                }
            )

        main_vm = MainWindowViewModel(controller=None, file_inspector=inspect)
        vm = InputAreaViewModel(main_vm=main_vm)

        vm.add_files([str(file_path)])

        assert calls == [str(file_path)]
        assert vm.selection_tone == "warning"
        assert vm.selection_message == "File extension (docx) does not match actual content (pdf)"
        assert vm.selection_detail == str(tmp_path)

    def test_supported_content_with_unknown_suffix_reaches_explicit_acceptance_state(
        self, vm: InputAreaViewModel, main_vm: MainWindowViewModel, tmp_path
    ) -> None:
        file_path = tmp_path / "renamed.bin"
        file_path.write_bytes(b"%PDF-1.4\n% deterministic probe\n")
        added: list[list[str]] = []
        vm.files_added.connect(added.append)

        vm.add_files([str(file_path)])

        assert added == [[str(file_path)]]
        assert len(main_vm.files) == 1
        assert main_vm.files[0].format == "pdf"
        assert main_vm.files[0].category == "layout"
        assert vm.selection_tone == "warning"

    def test_external_selection_sync_updates_visual_state_without_readding(
        self, vm: InputAreaViewModel, tmp_path
    ) -> None:
        file_path = tmp_path / "ipc.docx"
        file_path.write_text("content")
        added: list[list[str]] = []
        vm.files_added.connect(added.append)

        vm.sync_selection([FileRef(path=str(file_path), format="docx", category="document")])

        assert added == []
        assert "ipc.docx" in vm.selection_message
        assert vm.selection_detail == str(tmp_path)
        assert vm.selection_tone == "success"

    @pytest.mark.parametrize(
        "ref",
        [
            FileRef(
                path="warning.docx",
                format="pdf",
                category="layout",
                warning_message="The filename and content formats differ.",
            ),
            FileRef(
                path="reason.docx",
                format="pdf",
                category="layout",
                metadata={FILE_INSPECTION_METADATA_KEY: {"reason_message": "Re-check the detected file format."}},
            ),
        ],
    )
    def test_external_selection_sync_preserves_file_ref_admission_message(
        self,
        vm: InputAreaViewModel,
        ref: FileRef,
    ) -> None:
        vm.sync_selection([ref])

        assert vm.selection_tone == "warning"
        assert vm.selection_message in {
            "The filename and content formats differ.",
            "Re-check the detected file format.",
        }

    def test_add_folder_in_single_rejected(self, vm: InputAreaViewModel, tmp_path) -> None:
        folder = tmp_path / "subdir"
        folder.mkdir()
        signals: list[list] = []
        msg_signals: list[tuple] = []
        vm.files_added.connect(signals.append)
        vm.selection_message_changed.connect(lambda m, t: msg_signals.append((m, t)))
        vm.add_files([str(folder)])
        assert len(signals) == 0
        assert len(msg_signals) == 1
        assert msg_signals[0][1] == "warning"

    def test_add_zero_files_in_single_rejected(self, vm: InputAreaViewModel) -> None:
        signals: list[list] = []
        vm.files_added.connect(signals.append)
        vm.add_files([])
        assert len(signals) == 0

    def test_add_multiple_files_in_single_rejected(self, vm: InputAreaViewModel, tmp_path) -> None:
        f1 = tmp_path / "a.docx"
        f2 = tmp_path / "b.docx"
        f1.write_text("a")
        f2.write_text("b")
        signals: list[list] = []
        msg_signals: list[tuple] = []
        vm.files_added.connect(signals.append)
        vm.selection_message_changed.connect(lambda m, t: msg_signals.append((m, t)))
        vm.add_files([str(f1), str(f2)])
        assert len(signals) == 0
        assert len(msg_signals) == 1
        assert msg_signals[0][1] == "warning"


# ── File addition: batch mode ─────────────────────────────────────────


class TestBatchModeAddFiles:
    @pytest.fixture(autouse=True)
    def setup_batch(self, vm: InputAreaViewModel) -> None:
        vm.set_mode("batch")

    def test_add_batch_files(self, vm: InputAreaViewModel, tmp_path) -> None:
        f1 = tmp_path / "a.docx"
        f2 = tmp_path / "b.docx"
        f1.write_text("a")
        f2.write_text("b")
        signals: list[list] = []
        vm.files_added.connect(signals.append)
        vm.add_files([str(f1), str(f2)])
        assert len(signals) == 1
        assert len(signals[0]) == 2

    def test_add_folder_recursive(self, vm: InputAreaViewModel, tmp_path) -> None:
        folder = tmp_path / "subdir"
        folder.mkdir()
        f1 = folder / "test.docx"
        f1.write_text("content")
        signals: list[list] = []
        vm.files_added.connect(signals.append)
        vm.add_files([str(folder)])
        assert len(signals) == 1
        assert str(f1) in signals[0]

    def test_add_folder_recursive_uses_stable_path_order(self, vm: InputAreaViewModel, tmp_path) -> None:
        folder = tmp_path / "unsorted"
        nested = folder / "b-nested"
        nested.mkdir(parents=True)
        paths = [
            nested / "m.txt",
            folder / "z.txt",
            folder / "A.txt",
        ]
        for path in paths:
            path.write_text("content")
        vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        signals: list[list[str]] = []
        vm.files_added.connect(signals.append)

        vm.add_files([str(folder)])

        assert signals == [sorted((str(path) for path in paths), key=lambda path: path.casefold())]

    def test_add_folder_reports_partial_skips(self, vm: InputAreaViewModel, tmp_path) -> None:
        folder = tmp_path / "mixed"
        nested = folder / "nested"
        nested.mkdir(parents=True)
        supported = folder / "keep.txt"
        unsupported = nested / "ignore.bin"
        supported.write_text("1")
        unsupported.write_text("2")
        vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        signals: list[list[str]] = []
        msg_signals: list[tuple[str, str]] = []
        vm.files_added.connect(signals.append)
        vm.selection_message_changed.connect(lambda message, tone: msg_signals.append((message, tone)))

        vm.add_files([str(folder)])

        assert signals == [[str(supported)]]
        assert msg_signals[-1][1] == "warning"
        assert "1" in msg_signals[-1][0]

    def test_add_folder_recursive_is_not_limited_by_preview_scan_limit(
        self,
        vm: InputAreaViewModel,
        tmp_path,
    ) -> None:
        folder = tmp_path / "large"
        folder.mkdir()
        total = _BATCH_SCAN_LIMIT + 5
        for index in range(total):
            (folder / f"doc-{index:03}.txt").write_text("content")
        vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        signals: list[list[str]] = []
        vm.files_added.connect(signals.append)

        vm.add_files([str(folder)])

        assert len(signals) == 1
        assert len(signals[0]) == total

    def test_add_folder_prunes_tool_directories_case_insensitively(
        self,
        vm: InputAreaViewModel,
        tmp_path,
    ) -> None:
        kept = tmp_path / "keep.txt"
        hidden = tmp_path / ".GIT" / "hidden.txt"
        dependency = tmp_path / "node_modules" / "dependency.txt"
        for path in (kept, hidden, dependency):
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("content")
        vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        signals: list[list[str]] = []
        vm.files_added.connect(signals.append)

        vm.add_files([str(tmp_path)])

        assert signals == [[str(kept)]]

    def test_add_empty_batch_no_signal(self, vm: InputAreaViewModel) -> None:
        signals: list[list] = []
        vm.files_added.connect(signals.append)
        vm.add_files([])
        assert len(signals) == 0

    def test_no_supported_files_warning(self, vm: InputAreaViewModel, tmp_path) -> None:
        """When all files are unsupported extensions, emit warning."""
        # Use a custom file filter that rejects everything
        vm.file_filter = lambda p: False
        f1 = tmp_path / "test.docx"
        f1.write_text("content")
        signals: list[list] = []
        msg_signals: list[tuple] = []
        vm.files_added.connect(signals.append)
        vm.selection_message_changed.connect(lambda m, t: msg_signals.append((m, t)))
        vm.add_files([str(f1)])
        assert len(signals) == 0
        assert len(msg_signals) == 1
        assert msg_signals[0][1] == "warning"


# ── Drag preview ─────────────────────────────────────────────────────


class TestDragPreview:
    def test_single_supported_file_preview(self, vm: InputAreaViewModel, tmp_path) -> None:
        file_path = tmp_path / "hover.docx"
        file_path.write_text("content")

        preview = vm.build_drag_preview([str(file_path)])

        assert preview.added_count == 1
        assert preview.skipped_count == 0
        assert preview.tone == "info"
        assert file_path.name in preview.message
        assert preview.tooltip == str(file_path)

    def test_single_heif_file_preview_is_supported(self, vm: InputAreaViewModel, tmp_path) -> None:
        file_path = tmp_path / "photo.heif"
        file_path.write_bytes(b"heif")

        preview = vm.build_drag_preview([str(file_path)])

        assert preview.added_count == 1
        assert preview.skipped_count == 0
        assert preview.tone == "info"

    def test_single_unknown_suffix_preview_defers_to_content_admission(self, vm: InputAreaViewModel, tmp_path) -> None:
        file_path = tmp_path / "data.xml"
        file_path.write_text("<root />")

        preview = vm.build_drag_preview([str(file_path)])

        assert preview.added_count == 1
        assert preview.skipped_count == 0
        assert preview.tone == "info"
        assert file_path.name in preview.message

    def test_batch_folder_preview_counts_supported_and_skipped(self, vm: InputAreaViewModel, tmp_path) -> None:
        folder = tmp_path / "input"
        nested = folder / "nested"
        nested.mkdir(parents=True)
        supported = folder / "keep.txt"
        unsupported = nested / "ignore.bin"
        supported.write_text("1")
        unsupported.write_text("2")
        vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        vm.set_mode("batch")

        preview = vm.build_drag_preview([str(folder)])

        assert preview.added_count == 1
        assert preview.skipped_count == 1
        assert preview.has_recursive_scan is True
        assert preview.tone == "info"
        assert "1" in preview.message
        assert "ignore.bin" in preview.tooltip

    def test_batch_folder_preview_marks_large_scan_degraded(self, vm: InputAreaViewModel, tmp_path) -> None:
        folder = tmp_path / "large"
        folder.mkdir()
        for index in range(_BATCH_SCAN_LIMIT + 5):
            (folder / f"doc-{index:03}.txt").write_text("content")
        vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        vm.set_mode("batch")

        preview = vm.build_drag_preview([str(folder)])

        assert preview.has_recursive_scan is True
        assert preview.has_degraded_preview is True
        assert preview.added_count == _BATCH_SCAN_LIMIT
        assert str(_BATCH_SCAN_LIMIT) in preview.message

    def test_batch_folder_preview_warns_when_everything_is_skipped(self, vm: InputAreaViewModel, tmp_path) -> None:
        folder = tmp_path / "unsupported"
        folder.mkdir()
        unsupported = folder / "ignore.bin"
        unsupported.write_text("content")
        vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        vm.set_mode("batch")

        preview = vm.build_drag_preview([str(folder)])

        assert preview.added_count == 0
        assert preview.skipped_count == 1
        assert preview.tone == "warning"
        assert "ignore.bin" in preview.tooltip

    def test_batch_folder_preview_omits_ignored_tree_without_skip_noise(
        self,
        vm: InputAreaViewModel,
        tmp_path,
    ) -> None:
        kept = tmp_path / "keep.txt"
        hidden = tmp_path / "__PYCACHE__" / "hidden.txt"
        kept.write_text("content")
        hidden.parent.mkdir()
        hidden.write_text("content")
        vm.file_filter = lambda path: Path(path).suffix.lower() == ".txt"
        vm.set_mode("batch")

        preview = vm.build_drag_preview([str(tmp_path)])

        assert preview.added_count == 1
        assert preview.skipped_count == 0
        assert str(kept) in preview.tooltip
        assert "hidden.txt" not in preview.tooltip


# ── Clear files ───────────────────────────────────────────────────────


class TestClearFiles:
    def test_clear_emits_signal(self, vm: InputAreaViewModel) -> None:
        signals: list = []
        vm.files_cleared.connect(lambda: signals.append(True))
        vm.clear_files()
        assert len(signals) == 1

    def test_clear_resets_selection_message(self, vm: InputAreaViewModel) -> None:
        msg_signals: list[tuple] = []
        vm.selection_message_changed.connect(lambda m, t: msg_signals.append((m, t)))
        vm.clear_files()
        assert msg_signals == [("", "secondary")]


# ── Drag-and-drop text payload parsing ────────────────────────────────


class TestTextPayloadParsing:
    def test_single_path(self) -> None:
        result = InputAreaViewModel.extract_paths_from_text_payload("/home/user/test.docx")
        # The path typically won't exist in tests, so it returns empty
        assert isinstance(result, list)

    def test_braced_paths(self, tmp_path) -> None:
        f1 = tmp_path / "my file.docx"
        f1.write_text("content")
        text = f"{{{f1}}}"
        result = InputAreaViewModel.extract_paths_from_text_payload(text)
        assert str(f1) in result

    def test_space_separated_multiple(self, tmp_path) -> None:
        f1 = tmp_path / "a.docx"
        f2 = tmp_path / "b.xlsx"
        f1.write_text("a")
        f2.write_text("b")
        text = f"{f1} {f2}"
        result = InputAreaViewModel.extract_paths_from_text_payload(text)
        assert str(f1) in result
        assert str(f2) in result

    def test_url_scheme(self, tmp_path) -> None:
        f1 = tmp_path / "file.docx"
        f1.write_text("content")
        from PySide6.QtCore import QUrl

        url = QUrl.fromLocalFile(str(f1))
        text = url.toString()
        result = InputAreaViewModel.extract_paths_from_text_payload(text)
        assert str(f1) in result

    def test_quoted_paths(self, tmp_path) -> None:
        f1 = tmp_path / "file.docx"
        f1.write_text("content")
        text = f'"{f1}"'
        result = InputAreaViewModel.extract_paths_from_text_payload(text)
        assert str(f1) in result

    def test_deduplication(self, tmp_path) -> None:
        f1 = tmp_path / "file.docx"
        f1.write_text("content")
        text = f"{f1} {f1}"
        result = InputAreaViewModel.extract_paths_from_text_payload(text)
        assert result.count(str(f1)) == 1

    def test_empty_text(self) -> None:
        result = InputAreaViewModel.extract_paths_from_text_payload("")
        assert result == []

    def test_none_text(self) -> None:
        result = InputAreaViewModel.extract_paths_from_text_payload(None)  # type: ignore[arg-type]
        assert result == []

    def test_non_absolute_paths_ignored(self) -> None:
        result = InputAreaViewModel.extract_paths_from_text_payload("relative/path.docx")
        assert result == []

    def test_nonexistent_file_ignored(self) -> None:
        result = InputAreaViewModel.extract_paths_from_text_payload("/nonexistent/file.docx")
        assert result == []

    def test_os_error_from_invalid_path_token_is_ignored(self, tmp_path, monkeypatch) -> None:
        invalid_path = tmp_path / ("x" * 5000)
        original_exists = Path.exists

        def guarded_exists(path: Path) -> bool:
            if path == invalid_path:
                raise OSError("synthetic path lookup failure")
            return original_exists(path)

        monkeypatch.setattr(Path, "exists", guarded_exists)

        assert InputAreaViewModel.extract_paths_from_text_payload(str(invalid_path)) == []

    def test_max_paths_limit(self, tmp_path) -> None:
        """Ensure we cap at _TEXT_PAYLOAD_MAX_PATHS."""
        files = []
        for i in range(_TEXT_PAYLOAD_MAX_PATHS + 20):
            f = tmp_path / f"file_{i}.docx"
            f.write_text(f"content_{i}")
            files.append(f)
        text = " ".join(str(f) for f in files)
        result = InputAreaViewModel.extract_paths_from_text_payload(text)
        assert len(result) <= _TEXT_PAYLOAD_MAX_PATHS


# ── MIME URL extraction ──────────────────────────────────────────────


class TestMimeUrlExtraction:
    def test_local_file_urls(self, tmp_path) -> None:
        f1 = tmp_path / "test.docx"
        f1.write_text("content")
        from PySide6.QtCore import QUrl

        urls = [QUrl.fromLocalFile(str(f1))]
        result = InputAreaViewModel.extract_urls_from_mime_data(urls)
        # Normalize paths for cross-platform comparison
        result_norm = [str(Path(p)) for p in result]
        assert str(Path(str(f1))) in result_norm

    def test_remote_urls_filtered(self) -> None:
        from PySide6.QtCore import QUrl

        urls = [QUrl("https://example.com/file.docx")]
        result = InputAreaViewModel.extract_urls_from_mime_data(urls)
        assert result == []


# ── File dialog filter ────────────────────────────────────────────────


class TestFileDialogFilter:
    def test_filter_contains_all_files(self, vm: InputAreaViewModel) -> None:
        f = vm.build_file_dialog_filter()
        assert "All Files" in f or "*.*" in f

    def test_filter_contains_supported_types(self, vm: InputAreaViewModel) -> None:
        f = vm.build_file_dialog_filter()
        assert ";;" in f  # multiple filter groups

    def test_filter_includes_all_core_markdown_aliases(self, vm: InputAreaViewModel) -> None:
        assert "*.md" in vm.build_file_dialog_filter()
        assert "*.markdown" in vm.build_file_dialog_filter()


# ── Request add/folder dialog ─────────────────────────────────────────


class TestRequestDialogs:
    def test_request_add_dialog_does_not_crash(self, vm: InputAreaViewModel) -> None:
        vm.request_add_dialog()

    def test_request_add_dialog_force_batch(self, vm: InputAreaViewModel) -> None:
        vm.request_add_dialog(force_batch_mode=True)
        assert vm.mode == "batch"

    def test_request_add_folder_dialog(self, vm: InputAreaViewModel) -> None:
        vm.request_add_folder_dialog()

    def test_request_add_folder_dialog_force_batch(self, vm: InputAreaViewModel) -> None:
        vm.request_add_folder_dialog(force_batch_mode=True)
        assert vm.mode == "batch"


# ── File filter override ──────────────────────────────────────────────


class TestFileFilter:
    def test_custom_filter_used(self, vm: InputAreaViewModel, tmp_path) -> None:
        f1 = tmp_path / "test.xyz"
        f1.write_text("content")
        vm.file_filter = lambda p: True  # accept everything
        signals: list[list] = []
        vm.files_added.connect(signals.append)
        vm.add_files([str(f1)])
        assert len(signals) == 1

    def test_custom_filter_rejects(self, vm: InputAreaViewModel, tmp_path) -> None:
        f1 = tmp_path / "test.docx"
        f1.write_text("content")
        vm.file_filter = lambda p: False  # reject everything
        signals: list[list] = []
        vm.files_added.connect(signals.append)
        vm.add_files([str(f1)])
        assert len(signals) == 0
