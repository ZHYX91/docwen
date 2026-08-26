"""Tests for validate/execute user semantics (F-C4-041, F-C4-075).

Verify:
- validate_files detects empty files
- validate_files performs content-based format detection
- validate_files generates extension/content mismatch warnings
- _execute_single handles FileNotFoundError and DocWenError explicitly
"""

from __future__ import annotations

import argparse
import os
from unittest.mock import MagicMock, patch

import pytest

from docwen_application.runtime_capability_catalog import RuntimeRoute, parse_runtime_capability_catalog
from docwen_core.errors import DocWenError
from docwen_core.models import (
    AdmissionDecision,
    DetectionConfidence,
    DetectionMethod,
    FileInspection,
    FormatRelation,
    StructureStatus,
)

pytestmark = pytest.mark.unit


def _runtime_route(action: str, target: str) -> RuntimeRoute:
    return RuntimeRoute(
        id=f"test:markdown:{target}:{action or 'convert'}",
        operation="action" if action else "conversion",
        source="markdown",
        source_category="markdown",
        target=target,
        action_name=action,
        available=True,
        state="available",
        options=(),
    )


def _validation_capability_projection() -> dict[str, object]:
    """Return the canonical Runtime routes used by validation integration tests."""

    proofread_options = [
        "enable_sensitive_word",
        "enable_symbol_correction",
        "enable_symbol_pairing",
        "enable_typos_rule",
    ]
    sources: list[dict[str, object]] = [
        {
            "id": "markdown",
            "category": "markdown",
            "available": True,
            "routes": [
                {
                    "id": "proofread:markdown:markdown:validate",
                    "operation": "action",
                    "source": "markdown",
                    "target": "markdown",
                    "action": "validate",
                    "available": True,
                    "state": "available",
                    "options": proofread_options,
                }
            ],
        },
        {
            "id": "docx",
            "category": "document",
            "available": True,
            "routes": [
                {
                    "id": "proofread:docx:docx:validate",
                    "operation": "action",
                    "source": "docx",
                    "target": "docx",
                    "action": "validate",
                    "available": True,
                    "state": "available",
                    "options": proofread_options,
                }
            ],
        },
        {
            "id": "document",
            "category": "document",
            "available": True,
            "routes": [
                {
                    "id": "proofread:document:docx:validate",
                    "operation": "action",
                    "source": "document",
                    "target": "docx",
                    "action": "validate",
                    "available": True,
                    "state": "available",
                    "options": proofread_options,
                }
            ],
        },
    ]
    return {
        "resource": "formats",
        "contract": {"id": "docwen.runtime-capabilities", "version": 1},
        "runtime": {"state": "available", "platform": "windows"},
        "security": {"dependency_egress_guard": {}},
        "gates": [],
        "sources": sources,
        "counts": {
            "sources": 3,
            "routes": 3,
            "available_routes": 3,
            "unavailable_routes": 0,
            "actions": 3,
        },
    }


def _validation_controller() -> MagicMock:
    controller = MagicMock()
    controller.has_runtime = True
    controller.describe_runtime_capabilities.return_value = _validation_capability_projection()
    controller.execute_single.return_value = MagicMock(success=True)
    return controller


def _markdown_inspection(path: str) -> FileInspection:
    """Build a frozen valid inspection for controller-error unit tests."""

    return FileInspection(
        file_path=path,
        size_bytes=16,
        mtime_ns=0,
        extension=".md",
        declared_format="markdown",
        declared_category="markdown",
        detected_format="markdown",
        detected_category="markdown",
        workflow_category="markdown",
        detection_method=DetectionMethod.TEXT_SNIFF,
        confidence=DetectionConfidence.PROBABLE,
        structure_status=StructureStatus.NOT_APPLICABLE,
        relation=FormatRelation.EQUIVALENT_ALIAS,
        decision=AdmissionDecision.ALLOW,
        declared_supported=True,
        detected_supported=True,
    )


__all__ = (
    "AdmissionDecision",
    "DetectionConfidence",
    "DetectionMethod",
    "DocWenError",
    "FileInspection",
    "FormatRelation",
    "MagicMock",
    "StructureStatus",
    "_markdown_inspection",
    "_runtime_route",
    "_validation_capability_projection",
    "_validation_controller",
    "argparse",
    "os",
    "parse_runtime_capability_catalog",
    "patch",
    "pytest",
    "pytestmark",
)
