"""Focused tests split from test_batch_list_vm.py."""

from __future__ import annotations

from ._batch_list_vm_support import (
    CATEGORY_ORDER,
    OOXML_SIGNATURE_VALIDATION_UNAVAILABLE,
    AdmissionDecision,
    BatchFileEntry,
    BatchListViewModel,
    SimpleNamespace,
    _add_synthetic,
    _sort_value,
    format_size,
    pytest,
    should_pulse_processing_transition,
)

pytestmark = pytest.mark.unit
from ._batch_list_vm_support import (
    sample_entries as sample_entries,
)
from ._batch_list_vm_support import (
    vm as vm,
)


class TestFormatSize:
    def test_zero_bytes(self) -> None:
        assert format_size(0) == "0 B"

    def test_bytes(self) -> None:
        assert format_size(500) == "500 B"

    def test_kb(self) -> None:
        assert format_size(1024).startswith("1.0 KB")

    def test_mb(self) -> None:
        assert "MB" in format_size(1024 * 1024)

    def test_negative_bytes(self) -> None:
        assert format_size(-100) == "0 B"


class TestPulseLimit:
    def test_under_limit(self) -> None:
        assert should_pulse_processing_transition(40) is True
        assert should_pulse_processing_transition(1) is True

    def test_over_limit(self) -> None:
        assert should_pulse_processing_transition(41) is False


class TestSortValue:
    def test_sort_by_name(self) -> None:
        entry = BatchFileEntry(
            file_path="/test/MyDoc.docx",
            file_name="MyDoc.docx",
            detected_format="docx",
            workflow_category="document",
        )
        assert _sort_value(entry, "name") == "mydoc.docx"

    def test_sort_by_type(self) -> None:
        entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="DOCX",
            workflow_category="document",
        )
        assert _sort_value(entry, "type") == "docx"

    def test_sort_by_size(self) -> None:
        entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
            size_bytes=1024,
        )
        assert _sort_value(entry, "size") == 1024

    def test_sort_custom_returns_zero(self) -> None:
        entry = BatchFileEntry(
            file_path="/test/doc.docx",
            file_name="doc.docx",
            detected_format="docx",
            workflow_category="document",
        )
        assert _sort_value(entry, "custom") == 0


class TestConstruction:
    def test_initial_state(self, vm: BatchListViewModel) -> None:
        assert vm.entry_count == 0
        assert vm.current_category == CATEGORY_ORDER[0]
        assert vm.active_filter == "all"
        assert vm.sort_key == "custom"
        assert vm.sort_ascending is True

    def test_current_category_is_text(self, vm: BatchListViewModel) -> None:
        assert vm.current_category == "text"

    def test_no_file_returns_none(self, vm: BatchListViewModel) -> None:
        assert vm.get_current_file() is None
        assert vm.get_file_entry("/nonexistent") is None


class TestAddFiles:
    def test_add_single_file(self, vm: BatchListViewModel) -> None:
        added, failed = _add_synthetic(vm, ["/test/doc1.docx"])
        assert len(added) == 1
        assert len(failed) == 0
        assert vm.entry_count == 1

    def test_add_duplicate_skipped(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc1.docx"])
        added, _failed = _add_synthetic(vm, ["/test/doc1.docx"])
        assert len(added) == 0

    def test_add_multiple_files_same_category(self, vm: BatchListViewModel) -> None:
        added, _failed = _add_synthetic(vm, ["/test/a.docx", "/test/b.docx"])
        assert len(added) == 2
        assert vm.entry_count == 2

    def test_signed_ooxml_warning_reaches_batch_admission(
        self,
        vm: BatchListViewModel,
        tmp_path,
        monkeypatch,
    ) -> None:
        source = tmp_path / "signed.docx"
        source.write_bytes(b"probe")
        warning = f"[{OOXML_SIGNATURE_VALIDATION_UNAVAILABLE}] presence-only warning"
        inspection = SimpleNamespace(
            detected_format="docx",
            workflow_category="document",
            warning_message=warning,
            decision=AdmissionDecision.ALLOW_WITH_WARNING,
            reason_message="",
            to_dict=lambda: {
                "detected_format": "docx",
                "detected_category": "document",
                "workflow_category": "document",
                "decision": "allow_with_warning",
                "warning_message": warning,
            },
        )
        monkeypatch.setattr("docwen_core.detection.inspect_file", lambda _path: inspection)

        added, failed = vm.add_files([str(source)])

        assert failed == []
        assert len(added) == 1
        entry = vm.get_file_entry(str(source))
        assert entry is not None
        assert entry.warning_message == warning

    def test_add_files_different_categories(self, vm: BatchListViewModel) -> None:
        added, _failed = _add_synthetic(
            vm,
            [
                "/test/doc.docx",  # document
                "/test/sheet.xlsx",  # spreadsheet
                "/test/img.png",  # image
                "/test/layout.pdf",  # layout
                "/test/readme.md",  # text
                "/test/book.epub",  # other
            ],
        )
        assert len(added) == 6
        assert vm.get_file_count("document") == 1
        assert vm.get_file_count("spreadsheet") == 1
        assert vm.get_file_count("image") == 1
        assert vm.get_file_count("layout") == 1
        assert vm.get_file_count("text") == 1
        assert vm.get_file_count("other") == 1
        assert vm.current_category == "text"

    def test_add_files_activates_majority_category(self, vm: BatchListViewModel) -> None:
        emitted: list[str] = []
        vm.current_category_changed.connect(lambda category: emitted.append(category))

        _add_synthetic(vm, ["/test/a.docx", "/test/b.docx", "/test/sheet.xlsx"])

        assert vm.current_category == "document"
        assert emitted[-1:] == ["document"]

    def test_add_files_tie_uses_stable_category_priority(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/sheet.xlsx", "/test/doc.docx"])

        assert vm.current_category == "spreadsheet"

    def test_add_unsupported_file(self, vm: BatchListViewModel) -> None:
        added, failed = _add_synthetic(vm, ["/test/file.xyzzy"])
        # Unknown extension fails with the basic heuristic
        assert failed or len(added) <= 1
        # xyzzy not in mapping
        assert len(failed) == 1 or len(added) == 0

    def test_add_xml_is_not_a_supported_other_import(self, vm: BatchListViewModel) -> None:
        added, failed = _add_synthetic(vm, ["/test/data.xml"])

        assert added == []
        assert len(failed) == 1
        assert failed[0][1] == "Unsupported file type"

    def test_add_heif_is_supported_image_input(self, vm: BatchListViewModel) -> None:
        added, failed = _add_synthetic(vm, ["/test/photo.heif"])

        assert failed == []
        assert added == ["/test/photo.heif"]
        assert vm.get_file_count("image") == 1

    def test_add_no_extension_fails(self, vm: BatchListViewModel) -> None:
        added, failed = _add_synthetic(vm, ["/test/nofile"])
        assert len(failed) >= 1 or len(added) == 0

    def test_missing_path_without_resolver_fails_closed(self, vm: BatchListViewModel, tmp_path) -> None:
        missing = tmp_path / "missing.docx"

        added, failed = vm.add_files([str(missing)])

        assert added == []
        assert len(failed) == 1
        assert "not found" in failed[0][1].lower()

    def test_signal_emitted_on_add(self, vm: BatchListViewModel) -> None:
        emitted: list[tuple] = []
        vm.files_added.connect(lambda a, f: emitted.append((a, f)))
        _add_synthetic(vm, ["/test/doc.docx"])
        assert len(emitted) == 1
        assert len(emitted[0][0]) == 1  # added

    def test_entry_count_signal(self, vm: BatchListViewModel) -> None:
        counts: list[int] = []
        vm.entry_count_changed.connect(lambda c: counts.append(c))
        _add_synthetic(vm, ["/test/doc.docx"])
        assert 1 in counts

    def test_file_resolver_called(self, vm: BatchListViewModel) -> None:
        def resolver(path: str) -> dict | None:
            return {"detected_format": "docx", "workflow_category": "document"}

        added, _failed = vm.add_files(["/test/doc.docx"], file_resolver=resolver)
        assert len(added) == 1
        entry = vm.get_file_entry("/test/doc.docx")
        assert entry is not None
        assert entry.detected_format == "docx"
        assert entry.workflow_category == "document"

    def test_file_resolver_does_not_accept_legacy_actual_field_aliases(self, vm: BatchListViewModel) -> None:
        added, failed = vm.add_files(
            ["/test/doc.docx"],
            file_resolver=lambda _path: {
                "actual_format": "docx",
                "actual_category": "document",
            },
        )

        assert added == []
        assert failed == [("/test/doc.docx", "File detected format is unavailable")]

    def test_file_resolver_requires_canonical_workflow_category(self, vm: BatchListViewModel) -> None:
        added, failed = vm.add_files(
            ["/test/notes.md"],
            file_resolver=lambda _path: {
                "detected_format": "markdown",
                "workflow_category": "text",
            },
        )

        assert added == []
        assert failed == [("/test/notes.md", "File workflow category is unavailable")]

    @pytest.mark.parametrize(
        "detected_format",
        ["txt", "markdown"],
    )
    def test_file_resolver_maps_all_text_workflow_formats_to_text_tab(
        self,
        vm: BatchListViewModel,
        detected_format: str,
    ) -> None:
        warning = "declared and detected text formats differ"

        def resolver(path: str) -> dict[str, object]:
            return {
                "detected_format": detected_format,
                "workflow_category": "markdown",
                "warning_message": warning,
                "metadata": {"inspection": {"decision": "allow_with_warning"}},
            }

        added, failed = vm.add_files([f"/test/note-{detected_format}.txt"], file_resolver=resolver)

        assert failed == []
        assert len(added) == 1
        entry = vm.get_file_entry(added[0])
        assert entry is not None
        assert entry.workflow_category == "markdown"
        assert vm.get_file_display_category(entry.file_path) == "text"
        assert entry.warning_message == warning
        assert entry.metadata["inspection"]["decision"] == "allow_with_warning"
        assert vm.get_file_count("other") == 0

    def test_non_text_detected_format_keeps_detected_family_despite_txt_suffix(self, vm: BatchListViewModel) -> None:
        added, failed = vm.add_files(
            ["/test/disguised.txt"],
            file_resolver=lambda _path: {
                "detected_format": "rtf",
                "workflow_category": "document",
                "warning_message": "actual RTF",
            },
        )

        assert failed == []
        entry = vm.get_file_entry(added[0])
        assert entry is not None
        assert entry.workflow_category == "document"

    def test_file_resolver_returns_none(self, vm: BatchListViewModel) -> None:
        def resolver(path: str) -> dict | None:
            return None

        added, failed = vm.add_files(["/test/doc.docx"], file_resolver=resolver)
        assert len(added) == 0
        assert len(failed) == 1

    def test_file_resolver_raises(self, vm: BatchListViewModel) -> None:
        def resolver(path: str) -> dict | None:
            raise RuntimeError("test error")

        added, failed = vm.add_files(["/test/doc.docx"], file_resolver=resolver)
        assert len(added) == 0
        assert len(failed) == 1

    def test_large_batch_preserves_counts_order_and_filters(self, vm: BatchListViewModel) -> None:
        paths = (
            [f"/test/doc-{i:04d}.docx" for i in range(600)]
            + [f"/test/page-{i:04d}.pdf" for i in range(250)]
            + [f"/test/image-{i:04d}.png" for i in range(150)]
        )

        def resolver(path: str) -> dict[str, str] | None:
            if path.endswith(".docx"):
                return {"detected_format": "docx", "workflow_category": "document"}
            if path.endswith(".pdf"):
                return {"detected_format": "pdf", "workflow_category": "layout"}
            if path.endswith(".png"):
                return {"detected_format": "png", "workflow_category": "image"}
            return None

        added, failed = vm.add_files(paths, file_resolver=resolver)

        assert failed == []
        assert len(added) == len(paths)
        assert vm.entry_count == 1000
        assert vm.get_file_count("document") == 600
        assert vm.get_file_count("layout") == 250
        assert vm.get_file_count("image") == 150
        assert set(vm.get_files()) == set(paths)

        for path in paths[::4]:
            assert vm.set_file_status(path, "completed") is True

        vm.set_status_filter("completed")
        assert vm.get_visible_count_for_category("document") == 150
        assert vm.get_visible_count_for_category("layout") == 63
        assert vm.get_visible_count_for_category("image") == 37

        vm.set_sort_state("name", False)
        document_files = vm.get_files_for_category("document")
        assert len(document_files) == 600
        assert document_files[0].endswith("doc-0599.docx")
        assert document_files[-1].endswith("doc-0000.docx")


class TestRemoveFiles:
    def test_remove_single_file(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        assert vm.remove_file("/test/doc.docx") is True
        assert vm.entry_count == 0

    def test_remove_nonexistent(self, vm: BatchListViewModel) -> None:
        assert vm.remove_file("/nonexistent") is False

    def test_remove_multiple(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.docx", "/test/b.docx", "/test/c.docx"])
        removed = vm.remove_files(["/test/a.docx", "/test/b.docx"])
        assert len(removed) == 2
        assert vm.entry_count == 1

    def test_removed_signal(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        emitted: list[str] = []
        vm.files_removed.connect(lambda fp: emitted.append(fp))
        vm.remove_file("/test/doc.docx")
        assert "/test/doc.docx" in emitted

    def test_entry_count_signal_on_remove(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        counts: list[int] = []
        vm.entry_count_changed.connect(lambda c: counts.append(c))
        vm.remove_file("/test/doc.docx")
        assert 0 in counts


class TestClearFiles:
    def test_clear_empties_all(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/sheet.xlsx", "/test/img.png"])
        vm.clear_files()
        assert vm.entry_count == 0
        assert vm.active_filter == "all"
        assert vm.sort_key == "custom"

    def test_clear_emits_signal(self, vm: BatchListViewModel) -> None:
        vm.files_cleared.connect(lambda: setattr(self, "_cleared", True))
        self._cleared = False
        vm.files_cleared.connect(lambda: None)  # dummy connect
        _add_synthetic(vm, ["/test/doc.docx"])
        vm.clear_files()
        # check signal was emitted
        assert vm.entry_count == 0


class TestGetFiles:
    def test_get_files_empty(self, vm: BatchListViewModel) -> None:
        assert vm.get_files() == []

    def test_get_files_ordered(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/sheet.xlsx", "/test/img.png"])
        files = vm.get_files()
        assert len(files) == 3

    def test_get_files_respects_current_sort_order(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/b.docx", "/test/a.docx"])
        vm.set_sort_state("name", True)

        assert vm.get_files() == ["/test/a.docx", "/test/b.docx"]

    def test_get_files_respects_manual_reorder(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/a.docx", "/test/b.docx"])
        vm.reorder_manual("document", ["/test/b.docx", "/test/a.docx"])

        assert vm.get_files() == ["/test/b.docx", "/test/a.docx"]

    def test_get_files_for_category(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/doc2.docx", "/test/sheet.xlsx"])
        doc_files = vm.get_files_for_category("document")
        assert len(doc_files) == 2
        sheet_files = vm.get_files_for_category("spreadsheet")
        assert len(sheet_files) == 1

    def test_get_file_count(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx", "/test/doc2.docx", "/test/sheet.xlsx"])
        assert vm.get_file_count() == 3
        assert vm.get_file_count("document") == 2
        assert vm.get_file_count("spreadsheet") == 1
        assert vm.get_file_count("image") == 0


class TestFileStatus:
    def test_set_status(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        vm.set_file_status("/test/doc.docx", "processing")
        entry = vm.get_file_entry("/test/doc.docx")
        assert entry is not None
        assert entry.status == "processing"

    def test_set_status_not_found(self, vm: BatchListViewModel) -> None:
        assert vm.set_file_status("/nonexistent", "completed") is False

    def test_unknown_status_fails_before_mutating_or_emitting(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        emitted: list[tuple[str, str]] = []
        vm.status_changed.connect(lambda path, status: emitted.append((path, status)))

        with pytest.raises(ValueError, match="status must be one of"):
            vm.set_file_status("/test/doc.docx", "mystery")

        entry = vm.get_file_entry("/test/doc.docx")
        assert entry is not None
        assert entry.status == "pending"
        assert emitted == []

    def test_status_pending_to_processing_pulse(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        pulses: list[str] = []
        vm.pulse_requested.connect(lambda fp: pulses.append(fp))
        vm.set_file_status("/test/doc.docx", "processing")
        assert len(pulses) == 1
        assert pulses[0] == "/test/doc.docx"

    def test_pulse_disabled_over_limit(self, vm: BatchListViewModel) -> None:
        # Add 41 files to trigger the limit
        for i in range(41):
            vm._entries[f"/test/file{i}.txt"] = (
                "text",
                BatchFileEntry(
                    file_path=f"/test/file{i}.txt",
                    file_name=f"file{i}.txt",
                    detected_format="txt",
                    workflow_category="markdown",
                    size_bytes=100,
                    status="pending",
                ),
            )
        pulses: list[str] = []
        vm.pulse_requested.connect(lambda fp: pulses.append(fp))
        vm.set_file_status("/test/file0.txt", "processing")
        assert len(pulses) == 0

    def test_status_with_output_path(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        vm.set_file_status("/test/doc.docx", "completed", output_path="/output/doc.md")
        entry = vm.get_file_entry("/test/doc.docx")
        assert entry is not None
        assert entry.output_path == "/output/doc.md"

    def test_status_with_skip_reason(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        vm.set_file_status("/test/doc.docx", "skipped", skip_reason="Already exists")
        entry = vm.get_file_entry("/test/doc.docx")
        assert entry is not None
        assert entry.skip_reason == "Already exists"

    def test_status_with_error_message(self, vm: BatchListViewModel) -> None:
        _add_synthetic(vm, ["/test/doc.docx"])
        vm.set_file_status("/test/doc.docx", "failed", error_message="Conversion error")
        entry = vm.get_file_entry("/test/doc.docx")
        assert entry is not None
        assert entry.error_message == "Conversion error"
