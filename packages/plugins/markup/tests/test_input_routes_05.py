"""Focused tests split from test_input_routes.py."""

from __future__ import annotations

from ._input_routes_support import (
    Path,
    _document_node_root,
    _run_request,
    _successful_ocr,
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


class TestHtmlToMd:
    """Golden parity tests for ROUTE-HTML-001 and ROUTE-HTM-001."""

    @pytest.mark.integration
    def test_html_ocr_main_md_can_run_without_keeping_image_artifact(
        self,
        pipeline,
        sample_html_file_with_companion_image,
        tmp_path,
        monkeypatch,
    ) -> None:
        """HTML OCR should still work when images are not kept."""
        monkeypatch.setattr(
            "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
            lambda _path, **_kwargs: _successful_ocr("OCR text from HTML image"),
        )
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_html_ocr_no_image"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_html_file_with_companion_image,
            "html",
            output_dir,
            to_md_keep_images=False,
            to_md_enable_ocr=True,
            ocr_placement="main_md",
        )

        assert result.success
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)

        primary_artifacts = [a for a in result.artifacts if a.kind == "primary"]
        assert len(primary_artifacts) == 1
        markdown_path = Path(primary_artifacts[0].staging_path)
        _document_node_root(markdown_path, output_dir)
        assert markdown_path.exists()

        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        assert image_artifacts == []
        assert [a for a in result.artifacts if a.kind == "auxiliary"] == []

        content = markdown_path.read_text(encoding="utf-8")
        assert "> OCR text from HTML image" in content
        assert "picture.png" not in content
        assert str(Path(ws_mgr.root_dir)) not in content
