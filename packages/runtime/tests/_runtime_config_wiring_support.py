"""Contract tests for runtime config wiring (F-B2-002, F-I2a-036).

Verifies:
- ConfigLoader snapshots project into request-owned export semantics
- Link config stays in the request snapshot instead of a process-wide global
- Config flows from ConfigLoader → Adapter → Request → ExecutionContext
- Logging pre_init / init / reconfigure entry points exist and work
- Config is consumed by subsystems, not just loaded into memory
"""

from __future__ import annotations

import logging
import os
import tempfile
import tomllib
from dataclasses import dataclass
from logging import handlers as logging_handlers
from pathlib import Path
from typing import Any

import pytest

from docwen_runtime.config import build_document_style_catalog

pytestmark = pytest.mark.unit

PROJECT_CONFIGS = Path(__file__).resolve().parent.parent.parent.parent / "configs"


def _dummy_result(task_id: str) -> Any:
    from docwen_core.models.result import ConversionResult

    return ConversionResult(task_id=task_id, success=True)


__all__ = (
    "PROJECT_CONFIGS",
    "Any",
    "Path",
    "_dummy_result",
    "build_document_style_catalog",
    "dataclass",
    "logging",
    "logging_handlers",
    "os",
    "pytest",
    "pytestmark",
    "tempfile",
    "tomllib",
)
