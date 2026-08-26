"""Shared fixtures for markdown plugin tests."""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
from contextvars import ContextVar
from dataclasses import dataclass
from functools import cache
from pathlib import Path
from typing import Any

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[4]

LOCAL_SRC_PATHS = [
    PROJECT_ROOT,
    PROJECT_ROOT / "packages" / "core" / "src",
    PROJECT_ROOT / "packages" / "runtime" / "src",
    PROJECT_ROOT / "packages" / "plugins" / "markdown" / "src",
]

for path in reversed(LOCAL_SRC_PATHS):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_runtime.config.document_styles import build_document_style_catalog


@dataclass(frozen=True, slots=True)
class _OwnedTempPath:
    path: Path
    kind: str


_CURRENT_TEST_TEMP_PATHS: ContextVar[list[_OwnedTempPath] | None] = ContextVar(
    "markdown_test_owned_temp_paths",
    default=None,
)


def _register_test_temp_path(path: str | Path, *, kind: str) -> Path:
    owned_paths = _CURRENT_TEST_TEMP_PATHS.get()
    if owned_paths is None:
        raise RuntimeError("markdown temp helper used outside an active test")
    resolved = Path(path).resolve()
    owned_paths.append(_OwnedTempPath(path=resolved, kind=kind))
    return resolved


@pytest.fixture(autouse=True)
def _cleanup_owned_temp_paths() -> Any:
    """Remove only temp paths allocated by this test's shared helpers."""

    owned_paths: list[_OwnedTempPath] = []
    token = _CURRENT_TEST_TEMP_PATHS.set(owned_paths)
    try:
        yield
    finally:
        try:
            for owned in reversed(owned_paths):
                if owned.kind == "directory":
                    if owned.path.exists():
                        shutil.rmtree(owned.path)
                elif owned.kind == "file":
                    if owned.path.exists():
                        owned.path.unlink()
                else:
                    raise AssertionError(f"unknown markdown test temp ownership kind: {owned.kind}")
            remaining = [str(owned.path) for owned in owned_paths if owned.path.exists()]
            assert not remaining, f"markdown test temp ownership leaked: {remaining}"
        finally:
            _CURRENT_TEST_TEMP_PATHS.reset(token)


@cache
def _default_document_style_catalog() -> Any:
    """Load the immutable shipped style catalog once per pytest worker."""

    return build_document_style_catalog(
        {"gui": {"language": {"locale": "zh_CN"}}},
        locales_dir=PROJECT_ROOT / "i18n" / "locales",
    )


# ── Sample Markdown content ───────────────────────────────────────────

SAMPLE_MD_CONTENT = """# Heading Level 1

## Heading Level 2

Some paragraph with **bold** and *italic* text.

### Heading Level 3

Here is a `code span` and a [link](https://example.com).

#### Heading Level 4

##### Heading Level 5

###### Heading Level 6

---

## List Test

- Item 1
- Item 2
- Item 3

## Numbered List

1. First item
2. Second item
3. Third item

---

## Table Test

| Name  | Age | City     |
|-------|-----|----------|
| Alice | 30  | Beijing  |
| Bob   | 25  | Shanghai |
| Carol | 35  | Chengdu  |

## Blockquote

> This is a blockquote.
> It can span multiple lines.

## Code Block

```python
def hello():
    print("Hello, world!")
```
"""

SAMPLE_MD_NUMBERED = """一、 Heading Level 1

（一） Heading Level 2

Some paragraph with **bold** and *italic* text.

1. Heading Level 3

Here is a `code span` and a [link](https://example.com).

（1）Heading Level 4

① Heading Level 5

"""

SAMPLE_MD_TABLES = """# Table Test Document

| Name  | Age | City     |
|-------|-----|----------|
| Alice | 30  | Beijing  |
| Bob   | 25  | Shanghai |

Some text between tables.

| Product | Price | Qty |
|---------|-------|-----|
| Widget  | 10.50 | 100 |
| Gadget  | 25.00 | 50  |
"""


def make_context(
    input_path: str,
    target_format: str = "docx",
    action_name: str = "",
    options: dict[str, Any] | None = None,
    *,
    config_values: dict[str, Any] | None = None,
    numbering_registry: Any = None,
    heading_cleanup_rules: Any = (),
    document_style_catalog: Any = None,
) -> tuple[FakeExecutionContext, FakeWorkspaceHandle]:
    """Build a fake execution context for a single-file conversion."""
    staging = _register_test_temp_path(tempfile.mkdtemp(prefix="docwen_test_md_"), kind="directory")
    file_ref = FileRef(
        path=input_path,
        format="markdown",
        category="document",
    )
    request = ConversionRequest(
        request_id="test-md-req-001",
        input_refs=[file_ref],
        target_format=target_format,
        action_name=action_name,
        options=options or {},
        output_policy=OutputPolicy(),
    )
    workspace = FakeWorkspaceHandle(input_path, str(staging))
    config = FakeConfigView(config_values)
    progress = FakeProgressSink()
    cancellation = FakeCancellationTokenView()
    logger = FakePluginLogger()
    ctx = FakeExecutionContext(
        request,
        workspace,
        config,
        progress,
        cancellation,
        logger,
        numbering_registry=numbering_registry,
        heading_cleanup_rules=heading_cleanup_rules,
        document_style_catalog=(
            document_style_catalog if document_style_catalog is not None else _default_document_style_catalog()
        ),
    )
    return ctx, workspace


def write_temp_md(content: str, suffix: str = ".md") -> str:
    """Write content to a temp .md file and return the path."""
    fd, raw_path = tempfile.mkstemp(suffix=suffix, prefix="docwen_test_")
    os.close(fd)
    path = _register_test_temp_path(raw_path, kind="file")
    path.write_text(content, encoding="utf-8")
    return str(path)
