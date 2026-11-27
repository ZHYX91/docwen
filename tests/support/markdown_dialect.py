"""Converter context for Markdown dialect and native note integration tests."""

from pathlib import Path

from docwen_core.models.file_ref import FileRef
from docwen_core.models.request import ConversionRequest, OutputPolicy
from docwen_runtime.config.document_styles import build_document_style_catalog
from tests.support.cancellation import FakeCancellationTokenView
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

ROOT = Path(__file__).resolve().parents[2]


def dialect_context(tmp_path: Path, source: Path, target: str, extensions: dict) -> FakeExecutionContext:
    staging = tmp_path / f"staging-{target}"
    staging.mkdir()
    ref = FileRef(path=str(source), format="markdown" if target == "docx" else "docx", category="document")
    request = ConversionRequest(
        request_id=f"dialect-{target}",
        input_refs=[ref],
        target_format=target,
        options={"markdown_extensions": extensions, "heading_merge_mode": "never", "remove_numbering": False},
        output_policy=OutputPolicy(),
    )
    return FakeExecutionContext(
        request,
        FakeWorkspaceHandle(str(source), str(staging), (ref,)),
        FakeConfigView(),
        FakeProgressSink(),
        FakeCancellationTokenView(),
        FakePluginLogger(),
        document_style_catalog=build_document_style_catalog(
            {"gui": {"language": {"locale": "zh_CN"}}},
            locales_dir=ROOT / "i18n" / "locales",
        ),
    )
