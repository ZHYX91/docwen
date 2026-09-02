from __future__ import annotations

from dataclasses import FrozenInstanceError, fields
from inspect import signature

import pytest

from docwen_application.controller import ApplicationController
from docwen_application.ports.runtime import CancellationReservationPort, RuntimePort
from docwen_application.preconversion.pre_converter import PreConversionFailure
from docwen_runtime.engine.task_manager import TaskManager

pytestmark = pytest.mark.contract


def _parameter_names(owner: type[object], method: str) -> list[str]:
    return list(signature(getattr(owner, method)).parameters)


def test_application_and_runtime_expose_the_cancellation_ownership_contract() -> None:
    assert _parameter_names(ApplicationController, "prepare_execution_cancellation") == [
        "self",
        "request",
        "batch",
    ]
    assert _parameter_names(ApplicationController, "release_execution_cancellation") == [
        "self",
        "task_id",
        "reservation",
    ]
    assert _parameter_names(ApplicationController, "cancel") == ["self", "task_id"]

    assert _parameter_names(RuntimePort, "cancel") == ["self", "task_id"]
    assert _parameter_names(CancellationReservationPort, "reserve_cancellation") == ["self", "task_id"]
    assert _parameter_names(CancellationReservationPort, "release_cancellation") == ["self", "task_id"]
    assert _parameter_names(TaskManager, "cancel") == ["self", "task_id"]
    assert _parameter_names(TaskManager, "reserve_cancellation") == ["self", "task_id"]
    assert _parameter_names(TaskManager, "release_cancellation") == ["self", "task_id"]


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
