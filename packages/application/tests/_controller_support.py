"""Unit tests for ApplicationController."""

import shutil
import tempfile
import threading
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, call, patch

import pytest

from docwen_application.controller import ApplicationController, ControllerError
from docwen_application.ports.runtime import ConfigPort, PresenterPort, RuntimePort
from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import PRECONVERSION_INTERMEDIATES_OPTION, ConversionRequest, OutputPolicy

pytestmark = pytest.mark.unit


def _file_ref(path: str, fmt: str = "txt") -> FileRef:
    return FileRef(path=path, format=fmt, category="document")


def _request(
    request_id: str,
    *paths: str,
    target_format: str = "md",
    source_format: str = "txt",
    source_formats: tuple[str, ...] | None = None,
) -> ConversionRequest:
    formats = source_formats or (source_format,) * len(paths)
    if len(formats) != len(paths):
        raise ValueError("source_formats must match paths")
    return ConversionRequest(
        request_id=request_id,
        input_refs=[_file_ref(path, fmt) for path, fmt in zip(paths, formats, strict=True)],
        target_format=target_format,
    )


@pytest.fixture
def mock_runtime() -> MagicMock:
    return MagicMock(spec=RuntimePort)


@pytest.fixture
def mock_config() -> MagicMock:
    return MagicMock(spec=ConfigPort)


@pytest.fixture
def mock_presenter() -> MagicMock:
    return MagicMock(spec=PresenterPort)


@pytest.fixture(autouse=True)
def _isolate_controller_tests_from_core_admission(monkeypatch: pytest.MonkeyPatch) -> None:
    """Admission has dedicated tests; controller cases start with admitted refs."""

    monkeypatch.setattr("docwen_core.detection.enforce_file_admission", lambda request: request)


__all__ = (
    "PRECONVERSION_INTERMEDIATES_OPTION",
    "Any",
    "ApplicationController",
    "ControllerError",
    "ConversionRequest",
    "FileRef",
    "MagicMock",
    "OutputPolicy",
    "Path",
    "_file_ref",
    "_isolate_controller_tests_from_core_admission",
    "_request",
    "call",
    "mock_config",
    "mock_presenter",
    "mock_runtime",
    "patch",
    "pytest",
    "pytestmark",
    "shutil",
    "tempfile",
    "threading",
)
