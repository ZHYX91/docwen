"""File admission from frozen inspection facts, independent of widgets and workers."""

from __future__ import annotations

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


def check_frozen_request(request: ConversionRequest) -> None:
    """Revalidate exact ingress bytes on the execution worker before conversion."""
    from docwen_core.detection import inspect_file

    for ref in request.input_refs:
        raw_inspection = ref.metadata.get(FILE_INSPECTION_METADATA_KEY)
        try:
            inspection = inspect_file(ref.path)
        except FileNotFoundError as exc:
            raise ExecutionAdmissionError(
                _t("main_window.file_admission_missing", "The input file no longer exists: {path}", path=ref.path)
            ) from exc
        except OSError as exc:
            raise ExecutionAdmissionError(
                _t("main_window.file_admission_unreadable", "The input file cannot be read: {path}", path=ref.path)
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


class ExecutionAdmission:
    """Inspect pending confirmations and record acceptance only for matching facts."""

    def __init__(self, view_model: MainWindowViewModel, batch_list_vm: BatchListViewModel) -> None:
        self._view_model = view_model
        self._batch_list_vm = batch_list_vm

    @staticmethod
    def pending(request: ConversionRequest) -> list[tuple[FileRef, FileInspection]]:
        pending: list[tuple[FileRef, FileInspection]] = []
        for ref in request.input_refs:
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
