"""Focused tests split from test_input_routes.py."""

from __future__ import annotations

from ._input_routes_support import (
    ConversionRequest,
    FileRef,
    OutputPolicy,
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


class TestEnexToMd:
    """Golden parity tests for ROUTE-ENEX-001: enex → md."""

    @pytest.mark.integration
    def test_enex_artifact_metadata(self, pipeline, sample_enex_file, tmp_path) -> None:
        """Artifact must carry correct metadata."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_meta"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_enex_file, "enex", output_dir)

        assert result.success
        artifact = result.artifacts[0]
        assert artifact.artifact_id
        assert artifact.kind == "primary"
        assert artifact.suggested_name.endswith(".md")
        assert artifact.media_type == "text/markdown"
        assert artifact.is_primary is True
        assert artifact.metadata.get("note_count", 0) >= 1

    @pytest.mark.integration
    def test_enex_cancellation(self, pipeline, sample_enex_file, tmp_path) -> None:
        """ENEX→MD must support cancellation."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_cancel"
        output_dir.mkdir()

        request = ConversionRequest(
            request_id="enex-cancel-test",
            input_refs=[
                FileRef(
                    path=str(sample_enex_file),
                    format="enex",
                    category="markup",
                    size_bytes=sample_enex_file.stat().st_size,
                )
            ],
            target_format="md",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        task_mgr.cancel("enex-cancel-test")
        result = task_mgr.execute_single(request)

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "cancelled"
