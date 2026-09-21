"""Request snapshots and admission decisions without constructing a window."""

from __future__ import annotations

from copy import deepcopy
from types import SimpleNamespace
from typing import TYPE_CHECKING, cast

import pytest

from docwen_core.detection import inspect_file
from docwen_core.models import FILE_ADMISSION_ACCEPTANCE_METADATA_KEY, FILE_INSPECTION_METADATA_KEY
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import POSTPROCESS_PROOFREAD_OPTION
from docwen_gui.execution_admission import ExecutionAdmission, ExecutionAdmissionError, check_frozen_request
from docwen_gui.execution_requests import ExecutionRequestBuilder
from docwen_gui.path_identity import normalize_path

if TYPE_CHECKING:
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel

pytestmark = pytest.mark.unit


def _models(
    files: list[FileRef], entry: SimpleNamespace | None = None
) -> tuple[MainWindowViewModel, BatchListViewModel]:
    """Minimal model doubles; the builder must not need a QWidget or application."""
    return (
        cast("MainWindowViewModel", SimpleNamespace(files=files, controller=None)),
        cast("BatchListViewModel", SimpleNamespace(get_file_entry=lambda _path: entry)),
    )


@pytest.mark.parametrize("mode", ["single", "batch", "aggregate"])
def test_requests_own_nested_state_and_preserve_typed_input(mode, tmp_path):
    source = tmp_path / "input.md"
    source.write_text("# Original", encoding="utf-8")
    inspection = inspect_file(source)
    metadata = {FILE_INSPECTION_METADATA_KEY: inspection.to_dict(), "extra": {"values": ["original"]}}
    ref = FileRef(
        path=str(source),
        format="markdown",
        category="text",
        input_kind="document",
        input_role="source",
        logical_path="input.md",
        media_type="text/markdown",
        metadata=metadata,
    )
    options = {"spreadsheet_password": "not-history", "nested": {"values": ["original"]}, "unsupported": True}
    builder = ExecutionRequestBuilder(
        *_models([ref]),
        file_contexts=dict,
        selected_template=lambda: None,
    )
    request, context = getattr(builder, mode)(
        **({"file_path": str(source)} if mode == "single" else {"file_paths": [str(source)]}),
        target_format="md",
        action_name="test",
        options=options,
        route_options=("spreadsheet_password", "nested"),
    )
    options["nested"]["values"].append("caller-change")
    ref.metadata["extra"]["values"].append("list-change")
    assert request.options == {"spreadsheet_password": "not-history", "nested": {"values": ["original"]}}
    frozen = request.input_refs[0]
    assert frozen.metadata["extra"] == {"values": ["original"]}
    assert (frozen.input_kind, frozen.input_role, frozen.logical_path, frozen.media_type) == (
        "document",
        "source",
        "input.md",
        "text/markdown",
    )
    request.options["nested"]["values"].append("worker-change")
    assert context["options"] == {"spreadsheet_password": "<redacted>", "nested": {"values": ["original"]}}
    assert context["request_id"] == request.request_id
    assert context["file_path"] == source.as_posix()
    assert context.get(mode) is (None if mode == "single" else True)
    check_frozen_request(request)


def test_md_to_docx_proofread_options_survive_route_scoping_as_application_intent(tmp_path):
    source = tmp_path / "input.md"
    source.write_text("# Source", encoding="utf-8")
    inspection = inspect_file(source)
    ref = FileRef(
        path=str(source),
        format="markdown",
        category="markdown",
        metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
    )
    builder = ExecutionRequestBuilder(
        *_models([ref]),
        file_contexts=lambda: {normalize_path(str(source)): ("markdown", "markdown")},
        selected_template=lambda: ("docx", "standard"),
    )

    request, context = builder.single(
        file_path=str(source),
        target_format="docx",
        action_name="",
        options={
            "symbol_pairing": True,
            "symbol_correction": False,
            "typos_rule": True,
            "sensitive_word": False,
            "remove_numbering": True,
        },
        route_options=("remove_numbering", "template_name"),
    )

    assert request.options == {
        "template_name": "standard",
        "remove_numbering": True,
        POSTPROCESS_PROOFREAD_OPTION: {
            "enable_symbol_pairing": True,
            "enable_symbol_correction": False,
            "enable_typos_rule": True,
            "enable_sensitive_word": False,
        },
    }
    assert POSTPROCESS_PROOFREAD_OPTION not in context["options"]
    assert context["options"]["proofread"] == {
        "enabled": True,
        "options": {
            "enable_symbol_pairing": True,
            "enable_symbol_correction": False,
            "enable_typos_rule": True,
            "enable_sensitive_word": False,
        },
    }


def test_non_docx_conversion_does_not_carry_inert_proofread_options(tmp_path):
    source = tmp_path / "input.md"
    source.write_text("# Source", encoding="utf-8")
    inspection = inspect_file(source)
    ref = FileRef(
        path=str(source),
        format="markdown",
        category="markdown",
        metadata={FILE_INSPECTION_METADATA_KEY: inspection.to_dict()},
    )
    builder = ExecutionRequestBuilder(
        *_models([ref]),
        file_contexts=lambda: {normalize_path(str(source)): ("markdown", "markdown")},
        selected_template=lambda: ("docx", "standard"),
    )

    request, _ = builder.single(
        file_path=str(source),
        target_format="pdf",
        action_name="",
        options={"symbol_pairing": True, "typos_rule": True},
        route_options=(),
    )

    assert POSTPROCESS_PROOFREAD_OPTION not in request.options
    assert "symbol_pairing" not in request.options
    assert "typos_rule" not in request.options


def test_builder_reads_replaced_context_and_current_template(tmp_path):
    source = tmp_path / "input.md"
    source.write_text("# Source", encoding="utf-8")
    state = {"contexts": {}, "template": ("docx", "first-id")}
    builder = ExecutionRequestBuilder(
        *_models([]),
        file_contexts=lambda: state["contexts"],
        selected_template=lambda: state["template"],
    )
    assert builder.file_context(str(source)) is None
    state["contexts"] = {normalize_path(str(source)): ("markdown", "markdown")}
    first, _ = builder.single(file_path=str(source), target_format="docx", action_name="", options={})
    state["template"] = ("docx", "second-id")
    second, _ = builder.batch(file_paths=[str(source)], target_format="docx", action_name="", options={})
    assert first.options["template_name"] == "first-id"
    assert second.options["template_name"] == "second-id"
    assert first.request_id != second.request_id


@pytest.mark.parametrize("replace_live_facts", [False, True])
def test_confirmation_applies_only_to_the_inspection_shown(tmp_path, replace_live_facts):
    source = tmp_path / "renamed.docx"
    source.write_bytes(b"%PDF-1.4\n% first\n")
    first_facts = inspect_file(source).to_dict()
    ref = FileRef(
        path=str(source), format="pdf", category="layout", metadata={FILE_INSPECTION_METADATA_KEY: first_facts}
    )
    entry = SimpleNamespace(detected_format="pdf", workflow_category="layout", metadata=deepcopy(ref.metadata))
    view_model, batch = _models([ref], entry)
    builder = ExecutionRequestBuilder(view_model, batch, file_contexts=dict, selected_template=lambda: None)
    admission = ExecutionAdmission(view_model, batch)
    request, _ = builder.single(file_path=str(source), target_format="md", action_name="", options={})
    pending = admission.pending(request)
    assert len(pending) == 1
    if replace_live_facts:
        source.write_bytes(b"%PDF-1.4\n% replacement with new bytes\n")
        replacement = inspect_file(source).to_dict()
        ref.metadata[FILE_INSPECTION_METADATA_KEY].update(replacement)
        entry.metadata[FILE_INSPECTION_METADATA_KEY] = deepcopy(replacement)
    admission.accept(pending)
    assert not admission.pending(request)
    for metadata in [ref.metadata, entry.metadata]:
        assert (FILE_ADMISSION_ACCEPTANCE_METADATA_KEY in metadata) is not replace_live_facts
    if replace_live_facts:
        with pytest.raises(ExecutionAdmissionError):
            check_frozen_request(request)
    else:
        check_frozen_request(request)
