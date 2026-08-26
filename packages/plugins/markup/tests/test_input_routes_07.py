"""Focused tests split from test_input_routes.py."""

from __future__ import annotations

from ._input_routes_support import (
    Path,
    _assert_finalized_markdown_content,
    _assert_yaml_title,
    _document_node_root,
    _load_epub_old_system_fixture,
    _run_request,
    _successful_ocr,
    _write_epub_multi_resource_probe,
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


class TestMhtmlToMd:
    """Semantic parity tests for ROUTE-MHTML-001 and ROUTE-MHT-001."""

    @pytest.mark.integration
    def test_mhtml_ocr_placement_falls_back_to_export_semantics(
        self,
        pipeline,
        sample_mhtml_file,
        tmp_path,
        monkeypatch,
    ) -> None:
        """MHTML→MD must inherit Export OCR placement when the request omits ocr_placement."""
        monkeypatch.setattr(
            "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
            lambda _path, **_kwargs: _successful_ocr("OCR text from Export fallback"),
        )
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_mhtml_ocr_export_fallback"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_mhtml_file,
            "mhtml",
            output_dir,
            config_snapshot={"export": {"to_md_ocr_placement_mode": "main_md"}},
            to_md_enable_ocr=True,
        )

        assert result.success
        sidecar_artifacts = [a for a in result.artifacts if a.kind == "auxiliary" and a.media_type == "text/markdown"]
        assert sidecar_artifacts == []
        main_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "> OCR text from Export fallback" in main_content
        assert "embedded_ocr.md" not in main_content


class TestEpubToMd:
    """Semantic parity tests for ROUTE-EPUB-001 image resources."""

    @pytest.mark.parametrize(
        "image_html",
        [
            '<img src="images/pic.png" alt="chapter image" />',
            '<p><img src="images/pic.png" alt="chapter image" /></p>',
        ],
    )
    @pytest.mark.contract
    def test_epub_html_image_alt_is_text_for_block_and_inline_rendering(self, image_html: str) -> None:
        """BeautifulSoup attributes should become readable Markdown alt text."""
        from bs4 import BeautifulSoup

        from docwen_plugin_markup.publication.converter import EpubToMarkdownConverter

        soup = BeautifulSoup(f"<html><body>{image_html}</body></html>", "html.parser")

        markdown = EpubToMarkdownConverter._html_to_markdown(
            soup,
            {"images/pic.png": "![pic.png](pic.png)"},
        )

        assert markdown.strip() == "![chapter image](pic.png)"

    @pytest.mark.integration
    def test_epub_to_md_matches_old_system_semantic_fixture(self, pipeline, tmp_path) -> None:
        """Current EPUB→MD should preserve old-system core book semantics."""
        from ebooklib import epub

        fixture = _load_epub_old_system_fixture()
        input_epub = fixture["input_epub"]
        book = epub.EpubBook()
        book.set_identifier(input_epub["identifier"])
        book.set_title(input_epub["title"])
        book.set_language(input_epub["language"])
        book.add_author(input_epub["author"])
        chapter = epub.EpubHtml(
            title=input_epub["chapter_title"],
            file_name=input_epub["chapter_filename"],
            lang=input_epub["language"],
        )
        chapter.content = input_epub["chapter_html"]
        book.add_item(chapter)
        book.toc = [chapter]
        book.spine = ["nav", chapter]
        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())

        epub_path = tmp_path / input_epub["filename"]
        epub.write_epub(str(epub_path), book)
        output_dir = tmp_path / "output_epub_old_system_fixture"
        output_dir.mkdir()
        _plugin, task_mgr, ws_mgr = pipeline

        result = _run_request(task_mgr, epub_path, "epub", output_dir)

        assert result.success, f"EPUB conversion failed: {result.error.message if result.error else 'unknown'}"
        artifact = result.artifacts[0]
        expected_metadata = fixture["current_expected_semantics"]["artifact_metadata"]
        assert artifact.media_type == expected_metadata["media_type"]
        assert artifact.metadata["title"] == expected_metadata["title"]
        assert artifact.metadata["author"] == expected_metadata["author"]
        assert artifact.metadata["image_count"] == expected_metadata["image_count"]
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)

        artifact_path = Path(artifact.staging_path)
        _document_node_root(artifact_path, output_dir)
        assert artifact_path.exists()

        content = artifact_path.read_text(encoding="utf-8")
        assert str(Path(ws_mgr.root_dir)) not in content
        _assert_yaml_title(content, "title", input_epub["title"])
        for token in fixture["current_expected_semantics"]["required_markdown_tokens"]:
            assert token in content
        for fragment in fixture["current_expected_semantics"]["forbidden_markdown_fragments"]:
            assert fragment not in content

        non_empty_lines = [line.strip() for line in content.splitlines() if line.strip()]
        assert non_empty_lines[0] == fixture["probe_outputs"]["current_first_non_empty_line"]

    @pytest.mark.integration
    def test_epub_multi_resource_images_match_old_system_projection(self, pipeline, tmp_path) -> None:
        """Multiple EPUB resources should finalize in order with readable links."""
        fixture = _load_epub_old_system_fixture()
        probe = fixture["image_artifact_scope"]["multi_resource_probe"]
        _plugin, task_mgr, ws_mgr = pipeline

        epub_path = tmp_path / probe["filename"]
        _write_epub_multi_resource_probe(epub_path, probe)
        output_dir = tmp_path / "output_epub_multi_resources"
        output_dir.mkdir()

        result = _run_request(task_mgr, epub_path, "epub", output_dir)

        assert result.success
        expected = probe["current_expected_semantics"]
        assert [d.code for d in result.diagnostics if d.code in expected["diagnostics"]] == expected["diagnostics"]
        assert len(result.artifacts) == expected["artifact_count"] + 1

        primary_path = Path(result.artifacts[0].staging_path)
        node_root = _document_node_root(primary_path, output_dir)
        content = primary_path.read_text(encoding="utf-8")
        assert str(Path(ws_mgr.root_dir)) not in content
        for token in expected["required_markdown_tokens"]:
            assert token in content
        for fragment in expected["forbidden_markdown_fragments"]:
            assert fragment not in content

        first_token, second_token = expected["ordered_markdown_tokens"]
        assert content.index(first_token) < content.index(second_token)

        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        assert len(image_artifacts) == expected["image_count"]
        for artifact, expected_artifact in zip(image_artifacts, expected["expected_image_artifacts"], strict=True):
            assert artifact.suggested_name == expected_artifact["suggested_name"]
            assert artifact.media_type == expected_artifact["media_type"]
            artifact_path = Path(artifact.staging_path)
            assert _document_node_root(artifact_path, output_dir) == node_root
            assert artifact_path.exists()

    @pytest.mark.integration
    def test_epub_resource_links_honor_request_image_link_style(
        self,
        pipeline,
        sample_epub_with_image,
        tmp_path,
    ) -> None:
        """EPUB resource links must honor per-request Link style overrides."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_epub_request_link_style"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_epub_with_image,
            "epub",
            output_dir,
            image_link_style="markdown_embed",
        )

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "![pic.png](pic.png)" in content
        assert "![[pic.png]]" not in content

    @pytest.mark.integration
    def test_epub_resource_links_honor_request_image_mode_base64(
        self,
        pipeline,
        sample_epub_with_image,
        tmp_path,
    ) -> None:
        """EPUB resources must honor per-request base64 image mode."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_epub_request_base64_mode"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_epub_with_image,
            "epub",
            output_dir,
            image_mode="base64",
            image_link_style="markdown_embed",
        )

        assert result.success
        assert [a.kind for a in result.artifacts].count("image") == 0
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "![pic.png](data:image/png;base64," in content
        assert "](pic.png)" not in content

    @pytest.mark.integration
    def test_epub_yaml_frontmatter_consumes_locale_title_label(
        self,
        pipeline,
        sample_epub_with_image,
        tmp_path,
    ) -> None:
        """EPUB YAML frontmatter should consume app-resolved locale labels."""
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_epub_locale_yaml"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_epub_with_image,
            "epub",
            output_dir,
            yaml_key_labels={"title": "Titel"},
        )

        assert result.success
        content = _assert_finalized_markdown_content(result, output_dir, ws_mgr.root_dir)
        _assert_yaml_title(content, "Titel", "EPUB Image Test")

    @pytest.mark.integration
    def test_epub_ocr_image_md_creates_sidecar(
        self,
        pipeline,
        sample_epub_with_image,
        tmp_path,
        monkeypatch,
    ) -> None:
        """EPUB image OCR sidecars must be returned as final artifacts."""
        monkeypatch.setattr(
            "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
            lambda _path, **_kwargs: _successful_ocr("OCR text from EPUB image"),
        )
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_epub_ocr_sidecar"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_epub_with_image,
            "epub",
            output_dir,
            to_md_enable_ocr=True,
            ocr_placement="image_md",
        )

        assert result.success
        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        sidecar_artifacts = [a for a in result.artifacts if a.kind == "auxiliary" and a.media_type == "text/markdown"]
        assert len(image_artifacts) == 1
        assert len(sidecar_artifacts) == 1

        md_path = Path(result.artifacts[0].staging_path)
        image_path = Path(image_artifacts[0].staging_path)
        sidecar_path = Path(sidecar_artifacts[0].staging_path)
        node_root = _document_node_root(md_path, output_dir)
        assert _document_node_root(image_path, output_dir) == node_root
        assert _document_node_root(sidecar_path, output_dir) == node_root
        assert md_path.exists()
        assert image_path.exists()
        assert sidecar_path.exists()

        main_content = md_path.read_text(encoding="utf-8")
        assert sidecar_artifacts[0].logical_path is not None
        assert sidecar_artifacts[0].logical_path.split("/", 1)[1] in main_content
        assert str(Path(ws_mgr.root_dir)) not in main_content
        sidecar_content = sidecar_path.read_text(encoding="utf-8")
        assert "pic.png" in sidecar_content
        assert "> OCR text from EPUB image" in sidecar_content
        assert str(Path(ws_mgr.root_dir)) not in sidecar_content

    @pytest.mark.integration
    def test_epub_keep_images_false_omits_image_reference(
        self,
        pipeline,
        sample_epub_with_image,
        tmp_path,
    ) -> None:
        """EPUB to_md_keep_images=False must not leak original image src refs."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_epub_without_images"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_epub_with_image,
            "epub",
            output_dir,
            to_md_keep_images=False,
        )

        assert result.success
        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        assert image_artifacts == []

        main_content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "images/pic.png" not in main_content
        assert "pic.png" not in main_content


class TestPluginDispatch:
    """Verify MarkupPlugin correctly dispatches to the right converters."""

    @pytest.mark.contract
    def test_can_handle_all_markup_routes(self) -> None:
        """can_handle must return True for all markup input routes."""
        from docwen_plugin_markup import MarkupPlugin

        plugin = MarkupPlugin()
        for fmt in ("enex", "html", "mhtml", "htm", "mht", "epub"):
            assert plugin.can_handle(fmt, "md") is True, f"can_handle({fmt}, md) should be True"

    @pytest.mark.contract
    def test_can_handle_rejects_non_markup_routes(self) -> None:
        """can_handle must reject routes belonging to other plugins."""
        from docwen_plugin_markup import MarkupPlugin

        plugin = MarkupPlugin()
        assert plugin.can_handle("docx", "md") is False
        assert plugin.can_handle("pptx", "md") is False
        assert plugin.can_handle("pdf", "md") is False

    @pytest.mark.integration
    def test_convert_dispatch_enex(self, pipeline, sample_enex_file, tmp_path) -> None:
        """plugin.convert() must dispatch enex→md to EnexToMarkdownConverter."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_dispatch_enex"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_enex_file, "enex", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "Test Note Title" in content

    @pytest.mark.integration
    def test_convert_dispatch_html(self, pipeline, sample_html_file, tmp_path) -> None:
        """plugin.convert() must dispatch html→md to HtmlToMarkdownConverter."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_dispatch_html"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file, "html", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "Test HTML Document" in content


@pytest.mark.contract
def test_epub_missing_parser_is_a_dependency_failure() -> None:
    from docwen_plugin_markup.publication.converter import EpubToMarkdownConverter

    result = EpubToMarkdownConverter._dependency_missing("task", "ebooklib is required")

    assert result.success is False
    assert result.error is not None
    assert result.error.error_type == "dependency_missing"
    assert result.error.diagnostic_code == "EPUB2MD-DEPENDENCY-MISSING"
