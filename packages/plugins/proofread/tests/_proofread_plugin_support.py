"""Semantic tests for the Proofread plugin routes.

Covered routes:
- ACT-VALIDATE-DOCX: docx→docx with action_name="validate"
- ACT-VALIDATE-MD:   markdown→markdown with action_name="validate"

Coverage targets:
- Basic DOCX validation path
- Basic MD validation path
- Symbol pairing checks (unmatched brackets, quotes)
- Symbol correction checks (fullwidth digits)
- Typo checks (common Chinese misspellings)
- Empty document / no-error scenarios
- All-checks-disabled skip scenario
- Corrupted/invalid input handling
- Plugin dispatch routing
- Cancellation before execution
- TextValidator unit tests (four checks)
"""

from __future__ import annotations

import base64
import json
import os
import shutil
import tempfile
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pytest

from docwen_runtime.config import build_document_style_catalog

pytestmark = pytest.mark.golden

_PROJECT_ROOT = Path(__file__).resolve().parents[4]

_PROOFREAD_OLD_SYSTEM_FIXTURE = _PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_proofread_semantics.json"

_REAL_DICTIONARY_FIXTURE = (
    _PROJECT_ROOT / "tests" / "fixtures" / "files" / "proofread_numbering_real" / "typos-current.toml"
)


def _build_fake_context(
    input_path: str,
    staging_dir: str,
    target_format: str = "",
    options: dict[str, Any] | None = None,
    action_name: str = "",
    source_format: str = "",
    *,
    pre_cancelled: bool = False,
    proofread_rules: Any = None,
) -> Any:
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    detected_format = source_format or Path(input_path).suffix.lstrip(".")
    file_refs = [
        FileRef(
            path=input_path,
            format=detected_format,
            category="document",
        )
    ]
    request = ConversionRequest(
        request_id="test-proofread-001",
        input_refs=file_refs,
        target_format=target_format,
        action_name=action_name,
        options=options or {},
        output_policy=OutputPolicy(),
    )
    config = FakeConfigView()
    token = CancellationToken()
    if pre_cancelled:
        token.cancel("test cancellation")
    return FakeExecutionContext(
        request,
        FakeWorkspaceHandle(input_path, staging_dir),
        config,
        FakeProgressSink(),
        token.view(),
        FakePluginLogger(),
        proofread_rules=proofread_rules,
    )


def _create_test_docx(path: str, paragraphs: list[str]) -> None:
    """Create a minimal DOCX file with the given paragraph texts."""
    from docx import Document

    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)


def _load_proofread_old_system_fixture() -> dict[str, Any]:
    return json.loads(_PROOFREAD_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _proofread_rules_from_fixture(fixture: dict[str, Any]) -> Any:
    from docwen_core.models.proofread import ProofreadRules

    rules = fixture["rules"]
    return ProofreadRules(
        symbol_pairs=tuple(tuple(pair) for pair in rules["symbol_pairs"]),
        symbol_map={key: tuple(values) for key, values in rules["symbol_map"].items()},
        typos_map={key: tuple(values) for key, values in rules["typos_map"].items()},
        sensitive_words={key: tuple(values) for key, values in rules["sensitive_words"].items()},
    )


def _proofread_options_from_fixture(fixture: dict[str, Any]) -> dict[str, bool]:
    enabled = fixture["rules"]["checks_enabled"]
    return {
        "enable_symbol_pairing": bool(enabled["symbol_pairing"]),
        "enable_symbol_correction": bool(enabled["symbol_correction"]),
        "enable_typos_rule": bool(enabled["typos_rule"]),
        "enable_sensitive_word": bool(enabled["sensitive_word"]),
    }


def _expected_text_issues_from_fixture(fixture: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        {
            "rule_key": issue["rule_key"],
            "source": issue["source"],
            "start_pos": issue["start_pos"],
            "end_pos": issue["end_pos"],
            "error_text": issue["error_text"],
            "suggestion": issue["suggestion"],
        }
        for issue in fixture["expected_issues"]
    ]


def _normalize_text_error(error: Any) -> dict[str, Any]:
    from docwen_plugin_proofread.text_validator import rule_key

    suggestion = str(error.suggestion)
    if error.source == "sensitive":
        suggestion = "sensitive"
    elif error.source == "pairing":
        suggestion = "unclosed"
    return {
        "rule_key": rule_key(str(error.source)),
        "source": str(error.source),
        "start_pos": int(error.start_pos),
        "end_pos": int(error.end_pos),
        "error_text": str(error.error_text),
        "suggestion": suggestion,
    }


def _expected_markdown_issues_from_fixture(fixture: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        {
            "rule_key": issue["rule_key"],
            "range": {
                "start": {
                    "offset": int(issue["start_pos"]),
                    "line": 0,
                    "column": int(issue["start_pos"]),
                },
                "end": {
                    "offset": int(issue["end_pos"]),
                    "line": 0,
                    "column": int(issue["end_pos"]),
                },
            },
            "error_text": issue["error_text"],
        }
        for issue in fixture["expected_issues"]
    ]


def _markdown_issue_projection(issue: dict[str, Any]) -> dict[str, Any]:
    return {
        "rule_key": issue["rule_key"],
        "range": issue["range"],
        "error_text": issue["error_text"],
    }


def _count_by_key(values: Iterable[str]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    return counts


def _load_real_dictionary_rules(user_root: Path) -> Any:
    from docwen_runtime.config import build_proofread_rules
    from docwen_runtime.config.loader import ConfigLoader

    target = user_root / "proofread" / "typos.toml"
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(_REAL_DICTIONARY_FIXTURE, target)
    loader = ConfigLoader(base_dir=_PROJECT_ROOT / "configs", user_dir=user_root)
    return build_proofread_rules(loader.config.as_dict())


__all__ = (
    "_PROJECT_ROOT",
    "Any",
    "Path",
    "_build_fake_context",
    "_count_by_key",
    "_create_test_docx",
    "_expected_markdown_issues_from_fixture",
    "_expected_text_issues_from_fixture",
    "_load_proofread_old_system_fixture",
    "_load_real_dictionary_rules",
    "_markdown_issue_projection",
    "_normalize_text_error",
    "_proofread_options_from_fixture",
    "_proofread_rules_from_fixture",
    "base64",
    "build_document_style_catalog",
    "dataclass",
    "field",
    "json",
    "os",
    "pytest",
    "pytestmark",
    "tempfile",
    "zipfile",
)
