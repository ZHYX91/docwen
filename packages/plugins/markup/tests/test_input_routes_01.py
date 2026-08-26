"""Focused tests split from test_input_routes.py."""

from __future__ import annotations

from ._input_routes_support import (
    OcrOutcome,
    OcrStatus,
    Path,
    _run_request,
    pytest,
)
from ._input_routes_support import (
    pipeline as pipeline,
)

pytestmark = pytest.mark.golden
from ._input_routes_support import (
    sample_enex_file as sample_enex_file,
)
from ._input_routes_support import (
    sample_enex_with_markdown_resource as sample_enex_with_markdown_resource,
)
from ._input_routes_support import (
    sample_enex_with_resources as sample_enex_with_resources,
)
from ._input_routes_support import (
    sample_epub_with_image as sample_epub_with_image,
)
from ._input_routes_support import (
    sample_html_file as sample_html_file,
)
from ._input_routes_support import (
    sample_html_file_with_companion_image as sample_html_file_with_companion_image,
)
from ._input_routes_support import (
    sample_html_file_with_data_uri_image as sample_html_file_with_data_uri_image,
)
from ._input_routes_support import (
    sample_html_file_with_remote_image as sample_html_file_with_remote_image,
)
from ._input_routes_support import (
    sample_mhtml_file as sample_mhtml_file,
)


@pytest.mark.parametrize(
    ("source_format", "fixture_name", "expected_location"),
    [
        ("enex", "sample_enex_with_resources", "test_image.png"),
        ("html", "sample_html_file_with_companion_image", "picture.png"),
        ("mhtml", "sample_mhtml_file", "embedded.png"),
        ("epub", "sample_epub_with_image", "pic.png"),
    ],
)
@pytest.mark.integration
def test_markup_routes_keep_base_markdown_when_typed_ocr_fails(
    pipeline,
    request: pytest.FixtureRequest,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    source_format: str,
    fixture_name: str,
    expected_location: str,
) -> None:
    """A per-image OCR failure must not destroy any markup base conversion."""
    monkeypatch.setattr(
        "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
        lambda _path, **_kwargs: OcrOutcome(
            OcrStatus.RECOGNITION_FAILED,
            message="forced route-level OCR failure",
        ),
    )
    _plugin, task_mgr, _ws_mgr = pipeline
    input_path = request.getfixturevalue(fixture_name)
    output_dir = tmp_path / f"output_{source_format}_ocr_failure"
    output_dir.mkdir()
    events = []

    result = _run_request(
        task_mgr,
        input_path,
        source_format,
        output_dir,
        to_md_enable_ocr=True,
        ocr_placement="main_md",
        _on_event=events.append,
    )

    assert result.success
    merged_warnings = [diagnostic for diagnostic in result.diagnostics if diagnostic.code == "OCR-BEST-EFFORT"]
    assert len(merged_warnings) == 1
    assert merged_warnings[0].location == expected_location
    primary = next(artifact for artifact in result.artifacts if artifact.is_primary)
    assert Path(primary.staging_path).read_text(encoding="utf-8").strip()
    best_effort = [
        event for event in events if event.event_type == "diagnostic" and event.payload["code"] == "OCR-BEST-EFFORT"
    ]
    assert len(best_effort) == 1
    payload = best_effort[0].payload
    assert payload["level"] == "warning"
    assert payload["location"] == expected_location
    assert "status=recognition_failed" in payload["message"]
    assert f"{source_format} image {expected_location}" in payload["message"]
    assert "forced route-level OCR failure" not in payload["message"]
