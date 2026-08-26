"""Best-effort EPUB paths must report any degraded content or metadata."""

from __future__ import annotations

from pathlib import Path
from typing import ClassVar

import pytest

pytestmark = pytest.mark.contract


def _build_context(input_path: Path, staging_dir: Path):
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    request = ConversionRequest(
        request_id="epub-diagnostic-test",
        input_refs=[FileRef(path=str(input_path), format="epub", category="markup")],
        target_format="markdown",
        options={"to_md_keep_images": False},
        output_policy=OutputPolicy(),
    )
    return FakeExecutionContext(
        request,
        FakeWorkspaceHandle(str(input_path), str(staging_dir)),
        FakeConfigView(),
        FakeProgressSink(),
        CancellationToken().view(),
        FakePluginLogger(),
    )


def test_epub_degradation_is_typed_and_readable_chapters_survive(tmp_path, monkeypatch) -> None:
    import ebooklib
    from ebooklib import epub

    from docwen_plugin_markup.publication.converter import EpubToMarkdownConverter

    class Chapter:
        properties: ClassVar[list[str]] = []

        def __init__(self, item_id: str, *, broken: bool = False):
            self.id = item_id
            self._broken = broken

        def get_name(self) -> str:
            return f"{self.id}.xhtml"

        def get_content(self) -> bytes:
            if self._broken:
                raise ValueError("unreadable chapter")
            return b"<html><body><p>Readable chapter body.</p></body></html>"

    good = Chapter("good")
    bad = Chapter("bad", broken=True)

    class Book:
        spine: ClassVar[list[tuple[str, str]]] = [
            ("missing", "yes"),
            ("good", "yes"),
            ("bad", "yes"),
        ]
        toc: ClassVar[list[object]] = []

        def get_metadata(self, _namespace: str, _key: str):
            raise RuntimeError("metadata unavailable")

        def get_item_with_id(self, item_id: str):
            if item_id == "missing":
                raise KeyError(item_id)
            return {"good": good, "bad": bad}[item_id]

        def get_items_of_type(self, item_type: int):
            if item_type == ebooklib.ITEM_IMAGE:
                return []
            if item_type == ebooklib.ITEM_DOCUMENT:
                return [good, bad]
            return []

    source = tmp_path / "degraded.epub"
    source.write_bytes(b"fake epub bytes")
    staging = tmp_path / "staging"
    staging.mkdir()
    monkeypatch.setattr(epub, "read_epub", lambda _path: Book())

    result = EpubToMarkdownConverter().convert(_build_context(source, staging))

    assert result.success is True
    markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    assert "Readable chapter body." in markdown
    codes = [diagnostic.code for diagnostic in result.diagnostics]
    assert codes.count("EPUB2MD-METADATA-FALLBACK") == 2
    assert codes.count("EPUB2MD-SPINE-ITEM-FALLBACK") == 1
    assert codes.count("EPUB2MD-CHAPTER-SKIPPED") == 1
    assert result.metrics.extra["degradation_count"] == 4
    assert result.metrics.extra["spine_lookup_failure_count"] == 1
    assert result.metrics.extra["chapter_skip_count"] == 1
    assert result.artifacts[0].metadata["degradation_count"] == 4
