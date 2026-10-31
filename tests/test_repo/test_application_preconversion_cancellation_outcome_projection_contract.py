from __future__ import annotations

from dataclasses import FrozenInstanceError, fields
from inspect import Parameter, signature

import pytest

from docwen_application.controller import ApplicationController
from docwen_application.ports.runtime import CancellationReservationPort, RuntimePort
from docwen_application.preconversion.pre_converter import PreConversionFailure
from docwen_core.models.request import ConversionRequest, FileRef
from docwen_runtime.engine.task_manager import TaskManager

pytestmark = pytest.mark.contract


def _parameter_names(owner: type[object], method: str) -> list[str]:
    return list(signature(getattr(owner, method)).parameters)


def test_application_and_runtime_expose_the_cancellation_ownership_contract() -> None:
    assert _parameter_names(ApplicationController, "prepare_execution_cancellation") == [
        "self",
        "request",
        "batch",
        "on_reserved",
    ]
    prepare_parameters = signature(ApplicationController.prepare_execution_cancellation).parameters
    assert prepare_parameters["batch"].kind is Parameter.KEYWORD_ONLY
    assert prepare_parameters["batch"].default is False
    assert prepare_parameters["on_reserved"].kind is Parameter.KEYWORD_ONLY
    assert prepare_parameters["on_reserved"].default is None
    assert _parameter_names(ApplicationController, "release_execution_cancellation") == [
        "self",
        "task_id",
        "reservation",
    ]
    assert _parameter_names(ApplicationController, "cancel") == ["self", "task_id", "reservation"]
    cancel_parameters = signature(ApplicationController.cancel).parameters
    assert cancel_parameters["reservation"].kind is Parameter.KEYWORD_ONLY
    assert cancel_parameters["reservation"].default is None

    assert _parameter_names(RuntimePort, "cancel") == ["self", "task_id"]
    assert _parameter_names(CancellationReservationPort, "reserve_cancellation") == ["self", "task_id"]
    assert _parameter_names(CancellationReservationPort, "release_cancellation") == ["self", "task_id"]
    assert _parameter_names(TaskManager, "cancel") == ["self", "task_id"]
    assert _parameter_names(TaskManager, "reserve_cancellation") == ["self", "task_id"]
    assert _parameter_names(TaskManager, "release_cancellation") == ["self", "task_id"]


def test_optional_capture_and_precise_cancel_preserve_legacy_and_exact_identity() -> None:
    """Exercise the real in-memory owner; this is not Runtime publication evidence."""
    request = ConversionRequest(
        request_id="projection-cancellation-owner",
        input_refs=[FileRef(path="unopened-source.md", format="markdown", category="markdown")],
        target_format="docx",
    )
    controller = ApplicationController()
    previous = controller.prepare_execution_cancellation(request)
    assert controller.prepare_execution_cancellation(request) is previous
    controller.cancel(request.request_id)
    assert previous.token.is_cancelled
    controller.release_execution_cancellation(request.request_id, previous)
    assert controller._cancellation_scopes == {}

    captured: list[object] = []

    def capture(handle: object) -> None:
        assert controller._cancellation_scopes == {}
        captured.append(handle)

    current = controller.prepare_execution_cancellation(request, on_reserved=capture)
    assert captured == [current]
    assert current is not previous
    controller.cancel(request.request_id, reservation=previous)
    controller.release_execution_cancellation(request.request_id, previous)
    assert controller._cancellation_scopes[request.request_id] is current
    assert not current.token.is_cancelled
    controller.cancel(request.request_id, reservation=current)
    assert current.token.is_cancelled
    controller.release_execution_cancellation(request.request_id, current)
    assert controller._cancellation_scopes == {}


def test_preconversion_failure_is_a_frozen_structured_outcome() -> None:
    assert [field.name for field in fields(PreConversionFailure)] == [
        "message",
        "cancelled",
        "error_type",
        "diagnostic_code",
        "cleanup_message",
        "cleanup_failed",
    ]
    outcome = PreConversionFailure(
        message="bridge stopped cooperatively",
        cancelled=True,
        error_type="cancelled",
        diagnostic_code="OFFICE_CANCELLED",
        cleanup_message="private workspace retained",
        cleanup_failed=True,
    )

    assert outcome.cancelled is True
    assert outcome.error_type == "cancelled"
    assert outcome.diagnostic_code == "OFFICE_CANCELLED"
    assert outcome.cleanup_failed is True
    with pytest.raises(FrozenInstanceError):
        outcome.cancelled = False  # type: ignore[misc]
