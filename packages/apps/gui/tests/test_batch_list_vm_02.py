"""Focused tests split from test_batch_list_vm.py."""

from __future__ import annotations

from ._batch_list_vm_support import (
    CATEGORY_ORDER,
    BatchFileEntry,
    BatchListViewModel,
    _add_synthetic,
    pytest,
)

pytestmark = pytest.mark.unit
from ._batch_list_vm_support import (
    sample_entries as sample_entries,
)
from ._batch_list_vm_support import (
    vm as vm,
)


class TestFilter:
    def test_default_filter_all(self, vm: BatchListViewModel) -> None:
        assert vm.active_filter == "all"

    def test_set_filter(self, vm: BatchListViewModel) -> None:
        vm.set_status_filter("failed")
        assert vm.active_filter == "failed"

    def test_invalid_filter_defaults_to_all(self, vm: BatchListViewModel) -> None:
        vm.set_status_filter("invalid")
        assert vm.active_filter == "all"

    def test_set_same_filter_no_change(self, vm: BatchListViewModel) -> None:
        vm.set_status_filter("all")
        assert vm.active_filter == "all"

    def test_filter_signal(self, vm: BatchListViewModel) -> None:
        signals: list[str] = []
        vm.filter_changed.connect(lambda f: signals.append(f))
        vm.set_status_filter("failed")
        assert "failed" in signals

    def test_filter_matching(self, vm: BatchListViewModel) -> None:
        entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            status="failed",
        )
        vm._entries["/test/doc.docx"] = ("document", entry)
        vm.set_status_filter("failed")
        assert vm._entry_matches_filter(entry) is True
        vm.set_status_filter("all")
        assert vm._entry_matches_filter(entry) is True
        vm.set_status_filter("completed")
        assert vm._entry_matches_filter(entry) is False

    def test_cancelled_filter_matches_cancelled_entries(self, vm: BatchListViewModel) -> None:
        entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            status="cancelled",
        )
        vm._entries["/test/doc.docx"] = ("document", entry)

        vm.set_status_filter("cancelled")

        assert vm.active_filter == "cancelled"
        assert vm._entry_matches_filter(entry) is True

    def test_cancelled_entries_do_not_count_as_failed_retry_targets(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/sheet.xlsx"])
        vm.set_file_status("/test/doc.docx", "cancelled", error_message="Task was cancelled")
        vm.set_file_status("/test/sheet.xlsx", "failed", error_message="error")

        assert vm.get_failed_file_count() == 1
        assert vm.get_failed_files() == ["/test/sheet.xlsx"]

    def test_focus_failed_items(self, vm: BatchListViewModel) -> None:
        vm.set_status_filter("all")
        vm.focus_failed_items()
        assert vm.active_filter == "failed"


class TestFailedFiles:
    def test_no_failed_empty(self, vm: BatchListViewModel) -> None:
        assert vm.get_failed_files() == []
        assert vm.get_failed_file_count() == 0

    def test_get_failed_files(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/sheet.xlsx"])
        vm.set_file_status("/test/doc.docx", "failed", error_message="error")
        vm.set_file_status("/test/sheet.xlsx", "completed")
        failed = vm.get_failed_files()
        assert len(failed) == 1
        assert "/test/doc.docx" in failed

    def test_get_failed_by_category(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/sheet.xlsx"])
        vm.set_file_status("/test/doc.docx", "failed", error_message="error")
        assert len(vm.get_failed_files("document")) == 1
        assert len(vm.get_failed_files("spreadsheet")) == 0

    def test_reset_failed(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        vm.set_file_status("/test/doc.docx", "failed", error_message="error")
        reset = vm.reset_failed_files(["/test/doc.docx"])
        assert len(reset) == 1
        entry = vm.get_file_entry("/test/doc.docx")
        assert entry is not None
        assert entry.status == "pending"
        assert not entry.error_message  # should be cleared


class TestTabs:
    def test_activate_tab(self, vm: BatchListViewModel) -> None:
        assert vm.activate_tab("spreadsheet") is True
        assert vm.current_category == "spreadsheet"

    def test_activate_invalid_tab(self, vm: BatchListViewModel) -> None:
        assert vm.activate_tab("nonexistent") is False
        assert vm.current_category == CATEGORY_ORDER[0]

    def test_activate_same_tab(self, vm: BatchListViewModel) -> None:
        vm.activate_tab("spreadsheet")
        assert vm.activate_tab("spreadsheet") is False  # no change signal

    def test_tab_changed_signal(self, vm: BatchListViewModel) -> None:
        signals: list[str] = []
        vm.current_category_changed.connect(lambda c: signals.append(c))
        vm.activate_tab("document")
        assert "document" in signals


class TestSort:
    def test_set_sort_state(self, vm: BatchListViewModel) -> None:
        vm.set_sort_state("name", False)
        assert vm.sort_key == "name"
        assert vm.sort_ascending is False

    def test_invalid_sort_key(self, vm: BatchListViewModel) -> None:
        vm.set_sort_state("invalid")
        assert vm.sort_key == "custom"

    def test_sort_signal(self, vm: BatchListViewModel) -> None:
        signals: list[tuple] = []
        vm.sort_changed.connect(lambda k, a: signals.append((k, a)))
        vm.set_sort_state("size", False)
        assert ("size", False) in signals

    def test_get_sort_state(self, vm: BatchListViewModel) -> None:
        assert vm.get_sort_state() == ("custom", True)
        vm.set_sort_state("name", False)
        assert vm.get_sort_state() == ("name", False)


class TestSelection:
    def test_no_selection_on_empty(self, vm: BatchListViewModel) -> None:
        assert vm.get_current_file() is None

    def test_get_current_file(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        current = vm.get_current_file("document")
        assert current == "/test/doc.docx"

    def test_locate_file_entry(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/sheet.xlsx"])
        found, cat = vm.locate_file_entry("/test/sheet.xlsx")
        assert found is True
        assert cat == "spreadsheet"
        assert vm.current_category == "spreadsheet"

    def test_locate_nonexistent(self, vm: BatchListViewModel) -> None:
        found, cat = vm.locate_file_entry("/nonexistent")
        assert found is False
        assert cat is None

    def test_visible_count_respects_filter(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/doc2.docx"])
        vm.set_file_status("/test/doc.docx", "completed")
        # With filter "all", visible count is 2
        assert vm.get_visible_count_for_category("document") == 2
        vm.set_status_filter("completed")
        assert vm.get_visible_count_for_category("document") == 1


class TestRetryTargets:
    def test_no_targets(self, vm: BatchListViewModel) -> None:
        selected, category = vm.build_retry_targets("document", None)
        assert selected == []
        assert category == []

    def test_category_failed(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/doc2.docx"])
        vm.set_file_status("/test/doc.docx", "failed", error_message="error")
        vm.set_file_status("/test/doc2.docx", "completed")
        _selected, category = vm.build_retry_targets("document", None)
        assert category == ["/test/doc.docx"]

    def test_selected_failed(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        vm.set_file_status("/test/doc.docx", "failed", error_message="error")
        selected, category = vm.build_retry_targets("document", "/test/doc.docx")
        assert selected == ["/test/doc.docx"]
        assert category == ["/test/doc.docx"]


class TestReorder:
    def test_reorder_manual(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.docx", "/test/b.docx"])
        vm.reorder_manual("document", ["/test/b.docx", "/test/a.docx"])
        ordered = vm._get_ordered_paths_for_category("document")
        assert ordered == ["/test/b.docx", "/test/a.docx"]

    def test_reorder_switches_to_custom(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.docx", "/test/b.docx"])
        vm.set_sort_state("name", True)
        assert vm.sort_key == "name"
        vm.reorder_manual("document", ["/test/b.docx", "/test/a.docx"])
        assert vm.sort_key == "custom"

    def test_reorder_emits_sort_changed_when_switching_to_custom(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.docx", "/test/b.docx"])
        vm.set_sort_state("name", True)
        signals: list[tuple[str, bool]] = []
        vm.sort_changed.connect(lambda key, ascending: signals.append((key, ascending)))

        vm.reorder_manual("document", ["/test/b.docx", "/test/a.docx"])

        assert ("custom", True) in signals


class TestBatchFileEntry:
    def test_defaults(self) -> None:
        entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
        )
        assert entry.status == "pending"
        assert entry.output_path is None
        assert entry.error_message is None
        assert entry.skip_reason is None
        assert entry.warning_message is None
        assert entry.operation_id is None
        assert entry.size_bytes == 0

    def test_full_entry(self) -> None:
        entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            warning_message="Warning: font substitution",
            size_bytes=2048,
            status="completed",
            output_path="/output/doc.md",
            skip_reason=None,
            error_message=None,
            operation_id="op-1",
        )
        assert entry.size_bytes == 2048
        assert entry.warning_message == "Warning: font substitution"
        assert entry.operation_id == "op-1"
        assert entry.output_path == "/output/doc.md"

    def test_unknown_status_is_rejected_at_construction(self) -> None:
        with pytest.raises(ValueError, match="status must be one of"):
            BatchFileEntry(
                file_path="/test/doc.docx",
                file_name="doc.docx",
                detected_format="docx",
                workflow_category="document",
                status="mystery",
            )


class TestAggregateFileCollection:
    """Tests for get_aggregate_file_list and has_aggregate_targets."""

    def test_merge_pdfs_collects_layout_files(self, vm: BatchListViewModel) -> None:
        _add_synthetic(
            vm,
            [
                "/test/a.pdf",  # layout
                "/test/b.pdf",  # layout
                "/test/doc.docx",  # document — not layout
            ],
        )
        files = vm.get_aggregate_file_list("merge_pdfs")
        assert len(files) == 2
        assert "/test/a.pdf" in files
        assert "/test/b.pdf" in files
        assert "/test/doc.docx" not in files

    def test_merge_pdfs_excludes_non_pdf_layout_files(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.pdf", "/test/b.ofd", "/test/c.pdf"])

        assert vm.get_aggregate_file_list("merge_pdfs") == ["/test/a.pdf", "/test/c.pdf"]

    def test_merge_pdfs_respects_manual_order(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.pdf", "/test/b.pdf", "/test/c.pdf"])
        vm.reorder_manual("layout", ["/test/c.pdf", "/test/a.pdf", "/test/b.pdf"])

        assert vm.get_aggregate_file_list("merge_pdfs") == [
            "/test/c.pdf",
            "/test/a.pdf",
            "/test/b.pdf",
        ]

    def test_merge_tables_respects_manual_order(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/base.xlsx", "/test/collect-a.xlsx", "/test/collect-b.xlsx"])
        vm.reorder_manual("spreadsheet", ["/test/collect-b.xlsx", "/test/base.xlsx", "/test/collect-a.xlsx"])

        assert vm.get_aggregate_file_list("merge_tables") == [
            "/test/collect-b.xlsx",
            "/test/base.xlsx",
            "/test/collect-a.xlsx",
        ]

    def test_merge_images_respects_manual_order(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.png", "/test/b.jpg", "/test/c.tiff"])
        vm.reorder_manual("image", ["/test/c.tiff", "/test/a.png", "/test/b.jpg"])

        assert vm.get_aggregate_file_list("merge_images_to_tiff") == [
            "/test/c.tiff",
            "/test/a.png",
            "/test/b.jpg",
        ]

    def test_merge_tables_collects_spreadsheet_files(self, vm: BatchListViewModel) -> None:
        _add_synthetic(
            vm,
            [
                "/test/t1.xlsx",  # spreadsheet
                "/test/t2.xlsx",  # spreadsheet
                "/test/img.png",  # image — not spreadsheet
            ],
        )
        files = vm.get_aggregate_file_list("merge_tables")
        assert len(files) == 2

    def test_merge_images_collects_image_files(self, vm: BatchListViewModel) -> None:
        _add_synthetic(
            vm,
            [
                "/test/a.png",
                "/test/b.jpg",
                "/test/c.tiff",
                "/test/doc.docx",  # not image
            ],
        )
        files = vm.get_aggregate_file_list("merge_images_to_tiff")
        assert len(files) == 3

    def test_unknown_action_returns_empty(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.pdf", "/test/b.pdf"])
        files = vm.get_aggregate_file_list("unknown_action")
        assert files == []

    def test_no_matching_files_returns_empty(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/sheet.xlsx"])
        files = vm.get_aggregate_file_list("merge_pdfs")  # no layout files
        assert files == []

    def test_has_aggregate_targets_two_or_more(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.pdf", "/test/b.pdf"])
        assert vm.has_aggregate_targets("merge_pdfs") is True

    def test_has_aggregate_targets_insufficient(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/only.pdf"])
        assert vm.has_aggregate_targets("merge_pdfs") is False

    def test_has_aggregate_targets_none(self, vm: BatchListViewModel) -> None:
        assert vm.has_aggregate_targets("merge_pdfs") is False
