from __future__ import annotations

import importlib
import sys
from collections import defaultdict
from pathlib import Path

# ---------------------------------------------------------------------------
# sys.path 副作用：必须在任何可能 import docwen 的语句之前执行
# ---------------------------------------------------------------------------
_ROOT = Path(__file__).resolve().parents[2]
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
_FITZ_OK = _can_import("fitz")
_LATEX2MATHML_OK = _can_import("latex2mathml")
_QFLUENTWIDGETS_OK = _can_import("qfluentwidgets")

_COLLECTION_DEPENDENCY_STATUS = {
    "python-docx": _DOCX_OK,
    "lxml": _LXML_OK,
    "Pillow": _PIL_OK,
    "pandas": _PANDAS_OK,
    "PyMuPDF": _FITZ_OK,
}
_RUNTIME_SKIP_DEPENDENCY_STATUS = {
    "latex2mathml": _LATEX2MATHML_OK,
    "qfluentwidgets": _QFLUENTWIDGETS_OK,
}
_NOT_COLLECTED_BY_REASON: dict[str, set[str]] = defaultdict(set)
_REPORT_DIR_ENV = "DOCWEN_PYTEST_REPORT_DIR"

__all__ = [
    "_COLLECTION_DEPENDENCY_STATUS",
    "_DOCX_OK",
    "_FITZ_OK",
    "_LATEX2MATHML_OK",
    "_LXML_OK",
    "_NOT_COLLECTED_BY_REASON",
    "_PANDAS_OK",
    "_PIL_OK",
    "_QFLUENTWIDGETS_OK",
    "_REPORT_DIR_ENV",
    "_ROOT",
    "_RUNTIME_SKIP_DEPENDENCY_STATUS",
]
