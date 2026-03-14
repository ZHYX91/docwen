from __future__ import annotations

import importlib
import sys
from pathlib import Path
from typing import Any

_ROOT = Path(__file__).resolve().parents[1]
_SRC = _ROOT / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))


def _can_import(module: str) -> bool:
    try:
        importlib.import_module(module)
        return True
    except Exception:
        return False


_PIL_OK = _can_import("PIL.Image")
_DOCX_OK = _can_import("docx")
_LXML_OK = _can_import("lxml.etree")
_PANDAS_OK = _can_import("pandas")
_WATCHDOG_OK = _can_import("watchdog.events")
_FITZ_OK = _can_import("fitz")

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

_WATCHDOG_REQUIRED_DIR_NAMES = {"test_ipc"}


def pytest_ignore_collect(collection_path: Path, config: Any) -> bool:
    name = collection_path.name.lower()
    full = str(collection_path).lower()

    if not _DOCX_OK and "docx" in full:
        return True
    if not _DOCX_OK and name in _DOCX_REQUIRED_BASENAMES:
        return True
    if not _LXML_OK and name == "test_note_handler_replace_atomic.py":
        return True

    if not _PIL_OK and name in _PIL_REQUIRED_BASENAMES:
        return True
    if not _PANDAS_OK and name in _PANDAS_REQUIRED_BASENAMES:
        return True
    if not _WATCHDOG_OK and any(part in _WATCHDOG_REQUIRED_DIR_NAMES for part in collection_path.parts):
        return True
    return (not _FITZ_OK) and (name in _FITZ_REQUIRED_BASENAMES)
