from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from .dependencies import (
    _DOCX_OK,
    _FITZ_OK,
    _LXML_OK,
    _NOT_COLLECTED_BY_REASON,
    _PANDAS_OK,
    _PIL_OK,
    _ROOT,
)

_PIL_REQUIRED_BASENAMES = {
    "test_executor.py",
    "test_merge_images_to_tiff.py",
    "test_markdown_utils_base64_compress.py",
    "test_image_compression_atomic.py",
    "test_image_core.py",
    "test_image_limit_size_compression.py",
    "test_image_format_conversion_output_save.py",
    "test_image_to_pdf.py",
    "test_xlsx2md_core.py",
    "test_md2xlsx_data_uri_image.py",
}

_PANDAS_REQUIRED_BASENAMES = {
    "test_xlsx2md_blocks.py",
    "test_spreadsheet_pipeline_smoke.py",
}

_DOCX_REQUIRED_BASENAMES = {
    "test_md_frontmatter_parsing.py",
    "test_xml_processor_output_naming.py",
    "test_txt_input_only.py",
    "test_strategies_registry_specs.py",
}

_FITZ_REQUIRED_BASENAMES = {
    "test_invoice_cn_layout2md.py",
    "test_invoice_cn_pdf_parsing.py",
}


_LEGACY_DIR_NAMES = {
    "test_application",
    "test_config",
    "test_converter",
    "test_docx_spell",
    "test_gui",
    "test_security",
    "test_services",
    "test_table_merger",
    "test_template",
    "test_utils",
}

_LEGACY_CLI_TEST_BASENAMES = {
    "_executor_testkit.py",
    "test_executor_batch_and_contracts.py",
    "test_executor_catalog_and_listing.py",
    "test_executor_format_groups_diagnostics.py",
    "test_executor_routing_and_convert.py",
    "test_json_golden.py",
    "test_main_headless_batch.py",
    "test_main_validation.py",
    "test_utils_input.py",
}

_LEGACY_BASENAMES = {"test_coexistence.py"}


def _relative_collection_path(collection_path: Path) -> str:
    try:
        return collection_path.relative_to(_ROOT).as_posix()
    except ValueError:
        return collection_path.as_posix()


def _record_not_collected(collection_path: Path, reason: str) -> None:
    _NOT_COLLECTED_BY_REASON[reason].add(_relative_collection_path(collection_path))


def _legacy_collection_violation(collection_path: Path) -> str | None:
    name = collection_path.name.lower()
    if name in _LEGACY_DIR_NAMES or name in _LEGACY_CLI_TEST_BASENAMES or name in _LEGACY_BASENAMES:
        return f"legacy monolith test path returned: {_relative_collection_path(collection_path)}"
    return None


def _get_collection_gate_reason(collection_path: Path) -> str | None:
    name = collection_path.name.lower()
    full = str(collection_path).lower()

    if not _DOCX_OK and "docx" in full:
        return "missing python-docx dependency"
    if not _DOCX_OK and name in _DOCX_REQUIRED_BASENAMES:
        return "missing python-docx dependency"
    if not _LXML_OK and name == "test_note_handler_replace_atomic.py":
        return "missing lxml dependency"
    if not _PIL_OK and name in _PIL_REQUIRED_BASENAMES:
        return "missing Pillow dependency"
    if not _PANDAS_OK and name in _PANDAS_REQUIRED_BASENAMES:
        return "missing pandas dependency"
    if not _FITZ_OK and name in _FITZ_REQUIRED_BASENAMES:
        return "missing PyMuPDF dependency"
    return None


def pytest_ignore_collect(collection_path: Path, config: Any) -> bool:
    legacy_violation = _legacy_collection_violation(collection_path)
    if legacy_violation is not None:
        raise pytest.UsageError(legacy_violation)
    reason = _get_collection_gate_reason(collection_path)
    if reason is None:
        return False
    _record_not_collected(collection_path, reason)
    return True


__all__ = ["pytest_ignore_collect"]
