"""Golden and semantic parity tests for MD → DOCX conversion."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest
from tests.support.config import FakeConfigView

pytestmark = pytest.mark.contract

from docwen_plugin_markdown.to_docx.converter import MdToDocxConverter

from .conftest import (
    SAMPLE_MD_CONTENT,
    make_context,
    write_temp_md,
)

__all__ = (
    "SAMPLE_MD_CONTENT",
    "FakeConfigView",
    "MdToDocxConverter",
    "Path",
    "make_context",
    "pytest",
    "pytestmark",
    "re",
    "write_temp_md",
    "zipfile",
)
