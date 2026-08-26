"""Focused tests split from test_main_window_projection_binding.py."""

from __future__ import annotations

import pytest

from ._main_window_projection_binding_support import (
    SimpleNamespace,
    optimization_capability_projection,
)
from ._main_window_projection_binding_support import (
    left_frame as left_frame,
)

pytestmark = pytest.mark.gui
from ._main_window_projection_binding_support import (
    right_frame as right_frame,
)
from ._main_window_projection_binding_support import (
    right_stack as right_stack,
)
from ._main_window_projection_binding_support import (
    window as window,
)


class TestRuntimeRequestBinding:
    def test_batch_thread_setup_failure_reserves_batch_scope_and_rolls_back(
        self,
        window,
        tmp_path,
        monkeypatch,
    ) -> None:
        import docwen_gui.main_window as main_window_module

        first = tmp_path / "first.md"
        second = tmp_path / "second.md"
        first.write_text("# First", encoding="utf-8")
        second.write_text("# Second", encoding="utf-8")
        paths = [str(first).replace("\\", "/"), str(second).replace("\\", "/")]
        window._batch_list_vm.add_files(paths)
        request, context = window._build_batch_request(
            file_paths=paths,
            target_format="docx",
            action_name="",
            options={},
        )
        reservation = object()
        prepared: list[tuple[object, bool]] = []
        released: list[tuple[str, object]] = []

        def prepare(candidate: object, *, batch: bool = False) -> object:
            prepared.append((candidate, batch))
            return reservation

        controller = SimpleNamespace(
            has_runtime=True,
            describe_runtime_capabilities=optimization_capability_projection,
            prepare_execution_cancellation=prepare,
            release_execution_cancellation=lambda task_id, handle: released.append((task_id, handle)),
            stop=lambda: None,
        )
        window._view_model._controller = controller
        monkeypatch.setattr(window, "_build_batch_request", lambda **_kwargs: (request, context))
        monkeypatch.setattr(window, "_confirm_request_admission", lambda _request: True)

        def fail_thread_setup(**_kwargs: object) -> object:
            raise RuntimeError("Batch QThread setup failed")

        monkeypatch.setattr(main_window_module, "_ExecutionThread", fail_thread_setup)

        window._start_batch_execution(
            file_paths=paths,
            target_format="docx",
            action_name="",
            options={},
        )

        assert prepared == [(request, True)]
        assert released == [(request.request_id, reservation)]
        assert window._active_threads == {}
        assert window._action_area_vm.cancel_visible is False
        entries = [window._batch_list_vm.get_file_entry(path) for path in paths]
        assert [entry.status for entry in entries if entry is not None] == ["failed", "failed"]

    def test_split_pdf_request_preserves_pdf_source_and_action(self, window, tmp_path) -> None:
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "sample.pdf"
        source.write_bytes(b"%PDF-1.4\n")
        window._file_contexts = {_normalize_path(str(source)): ("pdf", "layout")}

        request, context = window._build_request(
            file_path=str(source),
            target_format="pdf",
            action_name="split_pdf",
            options={"split_mode": "custom", "pages": [1]},
        )

        input_ref = request.input_refs[0]
        assert request.target_format == "pdf"
        assert request.action_name == "split_pdf"
        assert request.options == {"split_mode": "custom", "pages": [1]}
        assert input_ref.format == "pdf"
        assert input_ref.category == "layout"
        assert context["action_name"] == "split_pdf"
        assert context["options"]["split_mode"] == "custom"

    def test_merge_pdfs_aggregate_request_contains_all_pdf_refs(self, window, tmp_path) -> None:
        first = tmp_path / "a.pdf"
        second = tmp_path / "b.pdf"
        first.write_bytes(b"%PDF-1.4\n")
        second.write_bytes(b"%PDF-1.4\n")
        window._batch_list_vm.add_files([str(first), str(second)])

        request, context = window._build_aggregate_request(
            file_paths=[str(first), str(second)],
            target_format="pdf",
            action_name="merge_pdfs",
            options={},
        )

        assert request.target_format == "pdf"
        assert request.action_name == "merge_pdfs"
        assert [ref.path for ref in request.input_refs] == [str(first), str(second)]
        assert [ref.format for ref in request.input_refs] == ["pdf", "pdf"]
        assert [ref.category for ref in request.input_refs] == ["layout", "layout"]
        assert context["aggregate"] is True
        assert context["total_count"] == 2

    def test_txt_document_context_builds_markdown_runtime_request(self, window, tmp_path) -> None:
        """Core currently classifies txt as document; GUI request routing must
        still preserve the old TXT-as-Markdown source workflow."""
        from docwen_gui.main_window import _normalize_path

        source = tmp_path / "note.txt"
        source.write_text("# Title\n\ncontent", encoding="utf-8")
        window._file_contexts = {_normalize_path(str(source)): ("markdown", "markdown")}

        request, _context = window._build_request(
            file_path=str(source),
            target_format="docx",
            action_name="",
            options={},
        )

        input_ref = request.input_refs[0]
        assert input_ref.format == "markdown"
        assert input_ref.category == "markdown"
