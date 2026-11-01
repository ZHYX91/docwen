"""ENEX multiple-note title uses the admitted public source identity."""

from pathlib import Path

import pytest
import yaml
from tests.support.config import FakeConfigView
from tests.support.execution import FakeExecutionContext
from tests.support.logging import FakePluginLogger
from tests.support.progress import FakeProgressSink
from tests.support.workspace import FakeWorkspaceHandle

from docwen_core.cancellation import CancellationToken
from docwen_core.models.document_node import ConversionIdentity
from docwen_core.models.file_ref import SOURCE_PRESENTATION_NAME_METADATA_KEY, FileRef
from docwen_core.models.request import ConversionRequest
from docwen_plugin_markup.note_export.converter import EnexToMarkdownConverter

pytestmark = pytest.mark.contract


@pytest.mark.parametrize("count", [1, 2])
def test_release_review_enex_title_source_fallback(tmp_path: Path, count: int) -> None:
    staged = tmp_path / "input-0000.enex"
    staged.write_text(
        "<en-export>"
        + "".join(
            f"<note><title>Note {index}</title><content><![CDATA[<en-note>body</en-note>]]></content></note>"
            for index in range(count)
        )
        + "</en-export>",
        encoding="utf-8",
    )
    ref = FileRef(
        path=str(staged),
        format="enex",
        category="markup",
        metadata={SOURCE_PRESENTATION_NAME_METADATA_KEY: "[公开] #笔记.enex"},
    )
    request = ConversionRequest(request_id="review-enex", input_refs=[ref], target_format="md")
    context = FakeExecutionContext(
        request=request,
        workspace=FakeWorkspaceHandle(str(staged), str(tmp_path)),
        config=FakeConfigView(),
        progress=FakeProgressSink(),
        cancellation=CancellationToken().view(),
        logger=FakePluginLogger(),
    )
    request.conversion_identity = ConversionIdentity.create(
        task_id="review", source_stem="truncated-output", source_format="enex", source_name="wrong-name.enex"
    )
    result = EnexToMarkdownConverter().convert(context)
    assert result.success, result.error
    text = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
    metadata = yaml.safe_load(text.split("---", 2)[1])
    expected = "Note 0" if count == 1 else "[公开] #笔记"
    assert metadata["title"] == expected
    assert metadata["aliases"] == [expected]
    assert "input-0000" not in text
