"""File admission from frozen inspection facts, independent of widgets and workers."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import TYPE_CHECKING

from docwen_core.models import (
    FILE_ADMISSION_ACCEPTANCE_METADATA_KEY,
    FILE_INSPECTION_METADATA_KEY,
    AdmissionDecision,
    FileInspection,
    admission_is_satisfied,
    make_admission_acceptance,
)
from docwen_gui.file_admission_i18n import render_file_inspection_message
from docwen_gui.i18n import t as _t
from docwen_gui.path_identity import normalize_path

if TYPE_CHECKING:
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest
    from docwen_gui.view_models.batch_list_vm import BatchListViewModel
    from docwen_gui.view_models.main_window_vm import MainWindowViewModel


class ExecutionAdmissionError(RuntimeError):
    """A localized reason why a requested execution cannot start."""


def _managed_resource_integrity(ref: "FileRef") -> None:
    from docwen_core.models.file_ref import (
        MANAGED_INPUT_SHA256_METADATA_KEY,
        MANAGED_INPUT_SIZE_BYTES_METADATA_KEY,
        MANAGED_RESOURCE_ID_METADATA_KEY,
    )

    resource_id = ref.metadata.get(MANAGED_RESOURCE_ID_METADATA_KEY)
    expected_size = ref.metadata.get(MANAGED_INPUT_SIZE_BYTES_METADATA_KEY)
    expected_sha = ref.metadata.get(MANAGED_INPUT_SHA256_METADATA_KEY)
    if (
        ref.input_role != "linked_resource"
        or ref.input_kind != "resource"
        or not isinstance(resource_id, str)
        or not resource_id
        or type(expected_size) is not int
        or expected_size < 0
        or not isinstance(expected_sha, str)
        or len(expected_sha) != 64
    ):
        raise ExecutionAdmissionError(_t("main_window.file_admission_invalid", "File inspection data is invalid."))

    path = Path(ref.path)
    try:
        if path.is_symlink() or not path.is_file():
            raise FileNotFoundError(ref.path)
        is_junction = getattr(path, "is_junction", None)
        if callable(is_junction) and is_junction():
            raise OSError("managed resource must not be a junction")
        stat = path.stat()
        if stat.st_size != expected_size:
            raise ValueError("managed resource size changed")
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            while chunk := stream.read(1024 * 1024):
                digest.update(chunk)
        if digest.hexdigest() != expected_sha:
            raise ValueError("managed resource hash changed")
    except FileNotFoundError as exc:
        raise ExecutionAdmissionError(
            _t("main_window.file_admission_missing", "The input file no longer exists: {path}", path=ref.path)
        ) from exc
    except OSError as exc:
        raise ExecutionAdmissionError(
            _t("main_window.file_admission_unreadable", "The input file cannot be read: {path}", path=ref.path)
        ) from exc
    except ValueError as exc:
        raise ExecutionAdmissionError(
            _t(
                "main_window.file_admission_changed",
                "The file changed after it was added. Remove it from the list and add it again to re-check the file, then retry.",
            )
        ) from exc


def _validate_structured_resource_group(request: "ConversionRequest") -> None:
    from docwen_core.models.clipboard_document import (
        ClipboardDocumentError,
        load_clipboard_document_bytes,
        validate_clipboard_resource_refs,
    )

    resources = tuple(ref for ref in request.input_refs if ref.input_role == "linked_resource")
    structured_sources = tuple(
        ref for ref in request.input_refs if ref.input_role == "source" and ref.format == "clipboard_document"
    )
    if not resources and not structured_sources:
        return
    if len(structured_sources) != 1 or any(
        ref.input_role not in {"source", "linked_resource"} for ref in request.input_refs
    ):
        raise ExecutionAdmissionError(_t("main_window.file_admission_invalid", "File inspection data is invalid."))
    source = structured_sources[0]
    try:
        document = load_clipboard_document_bytes(Path(source.path).read_bytes())
        validate_clipboard_resource_refs(document, resources)
    except FileNotFoundError as exc:
        raise ExecutionAdmissionError(
            _t("main_window.file_admission_missing", "The input file no longer exists: {path}", path=source.path)
        ) from exc
    except OSError as exc:
        raise ExecutionAdmissionError(
            _t("main_window.file_admission_unreadable", "The input file cannot be read: {path}", path=source.path)
        ) from exc
    except (ClipboardDocumentError, TypeError, ValueError) as exc:
        raise ExecutionAdmissionError(
            _t("main_window.file_admission_invalid", "File inspection data is invalid.")
        ) from exc


def check_frozen_request(request: ConversionRequest) -> None:
    """Revalidate exact ingress bytes on the execution worker before conversion."""
    from docwen_core.detection import reinspect_frozen_file

    for ref in request.input_refs:
        if ref.input_role == "linked_resource":
            _managed_resource_integrity(ref)
            continue
        raw_inspection = ref.metadata.get(FILE_INSPECTION_METADATA_KEY)
        if not isinstance(raw_inspection, dict):
            raise ExecutionAdmissionError(
                _t(
                    "main_window.file_admission_invalid",
                    "File inspection data is invalid.",
                )
            )
        try:
            frozen = FileInspection.from_dict(raw_inspection)
            inspection = reinspect_frozen_file(ref.path, frozen)
        except FileNotFoundError as exc:
            raise ExecutionAdmissionError(
                _t(
                    "main_window.file_admission_missing",
                    "The input file no longer exists: {path}",
                    path=ref.path,
                )
            ) from exc
        except OSError as exc:
            raise ExecutionAdmissionError(
                _t(
                    "main_window.file_admission_unreadable",
                    "The input file cannot be read: {path}",
                    path=ref.path,
                )
            ) from exc
        except (TypeError, ValueError) as exc:
            raise ExecutionAdmissionError(
                _t("main_window.file_admission_invalid", "File inspection data is invalid.")
            ) from exc
        if raw_inspection != inspection.to_dict():
            raise ExecutionAdmissionError(
                _t(
                    "main_window.file_admission_changed",
                    "The file changed after it was added. Remove it from the list and add it again to re-check the file, then retry.",
                )
            )
    _validate_structured_resource_group(request)


class ExecutionAdmission:
    """Inspect pending confirmations and record acceptance only for matching facts."""

    def __init__(self, view_model: MainWindowViewModel, batch_list_vm: BatchListViewModel) -> None:
        self._view_model = view_model
        self._batch_list_vm = batch_list_vm

    @staticmethod
    def pending(request: ConversionRequest) -> list[tuple[FileRef, FileInspection]]:
        pending: list[tuple[FileRef, FileInspection]] = []
        for ref in request.input_refs:
            if ref.input_role == "linked_resource":
                _managed_resource_integrity(ref)
                continue
            raw_inspection = ref.metadata.get(FILE_INSPECTION_METADATA_KEY)
            if not isinstance(raw_inspection, dict):
                raise ExecutionAdmissionError(
                    _t("main_window.file_admission_invalid", "File inspection data is invalid.")
                )
            try:
                inspection = FileInspection.from_dict(raw_inspection)
            except (TypeError, ValueError) as exc:
                raise ExecutionAdmissionError(
                    _t("main_window.file_admission_invalid", "File inspection data is invalid.")
                ) from exc
            if inspection.decision is AdmissionDecision.BLOCK:
                raise ExecutionAdmissionError(
                    render_file_inspection_message(inspection, prefer_reason=True)
                    or _t("main_window.file_admission_blocked", "The selected file cannot be processed.")
                )
            if not admission_is_satisfied(inspection, ref.metadata):
                pending.append((ref, inspection))
        _validate_structured_resource_group(request)
        return pending

    def accept(self, pending: list[tuple[FileRef, FileInspection]]) -> None:
        for ref, inspection in pending:
            acceptance = make_admission_acceptance(inspection)
            ref.metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY] = acceptance
            frozen_facts = inspection.to_dict()
            normalized = normalize_path(ref.path)
            for source_ref in self._view_model.files:
                if (
                    normalize_path(source_ref.path) == normalized
                    and source_ref.metadata.get(FILE_INSPECTION_METADATA_KEY) == frozen_facts
                ):
                    source_ref.metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY] = dict(acceptance)
            entry = self._batch_list_vm.get_file_entry(ref.path)
            if entry is not None and entry.metadata.get(FILE_INSPECTION_METADATA_KEY) == frozen_facts:
                entry.metadata[FILE_ADMISSION_ACCEPTANCE_METADATA_KEY] = dict(acceptance)
