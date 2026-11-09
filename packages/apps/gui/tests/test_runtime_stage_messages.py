"""Error and progress copy is stable, actionable and does not leak paths."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from docwen_gui.execution_presenter import _localized_failure_message
from docwen_gui.view_models.info_area_vm import InfoAreaViewModel, _localized_progress_stage

pytestmark = pytest.mark.gui


@pytest.mark.parametrize(
    "message",
    [
        "Preparing output folder",
        "Preparing temporary workspace",
        "Starting conversion via markdown",
        "Reading C:/Users/private/document.md",
    ],
)
def test_stage_copy_does_not_leak_backend_or_private_filename(message: str) -> None:
    label = _localized_progress_stage(message)
    assert label and "private" not in label and "markdown" not in label.casefold()
    assert "Starting conversion via" not in label


@pytest.mark.parametrize("code", ["WORKSPACE_ROOT_UNAVAILABLE", "OUTPUT_PREFLIGHT_FAILED"])
def test_file_preparation_failure_has_actionable_public_copy(code: str) -> None:
    result = _localized_failure_message(
        SimpleNamespace(diagnostic_code=code, error_type="conversion_failed", message="C:/Users/private/secret")
    )
    assert result and "private" not in result and "C:/Users" not in result


def test_live_stage_and_cancelling_hint_appear_in_task_summary(qapp) -> None:
    info = InfoAreaViewModel()
    info.begin_task(operation_id="one", current_file="sample.md", total_count=1)
    info.update_task_progress("one", message="Preparing temporary workspace", percent=0.0)
    assert _localized_progress_stage("Preparing temporary workspace") in info.status_summary_text
    info.mark_cancelling("one")
    assert info.task_summary.state == "cancelling"
    assert info.status_summary_text
    assert _localized_progress_stage("Preparing temporary workspace") not in info.status_summary_text
    info.clear_task_summary()


def test_finalizing_output_is_presented_as_saving_not_conversion(qapp):
    from docwen_gui.i18n import t

    message = _localized_progress_stage("Finalizing output")
    assert message == t("info_area.phase_saving", "Saving the result…")
