"""Focused tests split from test_input_routes.py."""

from __future__ import annotations

from ._input_routes_support import (
    ConversionRequest,
    FileRef,
    OutputPolicy,
    Path,
    _assert_finalized_markdown_content,
    _assert_yaml_title,
    _document_node_root,
    _load_html_old_system_fixture,
    _run_request,
    _write_html_multi_resource_probe,
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

    @pytest.mark.contract
    def test_image_node_replacement_uses_plain_token_paragraph(self) -> None:
        """F-G2-005: image substitution leaves one plain token node in the HTML tree."""
        import lxml.etree as etree
        import lxml.html as lxml_html

        from docwen_plugin_markup.web_archive.converter import _replace_node_with_text

        document = lxml_html.fragment_fromstring(
            "<div><span>before</span><img src='image.png'>tail text<span>after</span></div>"
        )
        image = document.xpath(".//img")[0]
        _replace_node_with_text(etree, image, "DOCWEN_IMAGE_TOKEN")

        assert [child.tag for child in document] == ["span", "p", "span"]
        assert document[1].text == "DOCWEN_IMAGE_TOKEN"
        assert document[1].tail == "tail text"
        assert document.xpath(".//img") == []

        detached = etree.Element("img", src="image.png", alt="preview")
        _replace_node_with_text(etree, detached, "DETACHED_TOKEN")
        assert detached.tag == "p"
        assert detached.text == "DETACHED_TOKEN"
        assert detached.attrib == {}

    @pytest.mark.integration
    def test_html_conversion_succeeds(self, pipeline, sample_html_file, tmp_path) -> None:
        """HTML file must convert to Markdown successfully."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file, "html", output_dir)

        assert result.success, f"HTML conversion failed: {result.error.message if result.error else 'unknown'}"
        assert len(result.artifacts) >= 1
        assert result.artifacts[0].is_primary is True

        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert len(content.strip()) > 0

    @pytest.mark.integration
    def test_html_to_md_matches_old_system_semantic_fixture(
        self,
        pipeline,
        tmp_path,
    ) -> None:
        """Current HTML→MD should preserve old-system core Markdown semantics."""
        fixture = _load_html_old_system_fixture()
        _plugin, task_mgr, ws_mgr = pipeline

        html_path = tmp_path / fixture["input_html"]["filename"]
        html_path.write_text(fixture["input_html"]["body"], encoding="utf-8")
        output_dir = tmp_path / "output_html_old_system_fixture"
        output_dir.mkdir()

        result = _run_request(task_mgr, html_path, "html", output_dir)

        assert result.success, f"HTML conversion failed: {result.error.message if result.error else 'unknown'}"
        artifact = result.artifacts[0]
        assert artifact.media_type == "text/markdown"
        assert artifact.metadata["image_count"] == fixture["current_expected_semantics"]["image_count"]
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)

        artifact_path = Path(artifact.staging_path)
        _document_node_root(artifact_path, output_dir)
        assert artifact_path.exists()

        content = artifact_path.read_text(encoding="utf-8")
        assert str(Path(ws_mgr.root_dir)) not in content
        _assert_yaml_title(content, "title", fixture["input_html"]["title"])
        for token in fixture["current_expected_semantics"]["required_markdown_tokens"]:
            assert token in content
        for fragment in fixture["current_expected_semantics"]["forbidden_markdown_fragments"]:
            assert fragment not in content

        non_empty_lines = [line.strip() for line in content.splitlines() if line.strip()]
        assert non_empty_lines[:2] == fixture["current_expected_semantics"]["first_non_empty_lines"]
        assert (
            content.count(fixture["input_html"]["title"]) == fixture["current_expected_semantics"]["title_occurrences"]
        )

    @pytest.mark.integration
    def test_html_title_extraction(self, pipeline, sample_html_file, tmp_path) -> None:
        """HTML <title> must be extracted into YAML metadata."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html_title"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file, "html", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        _assert_yaml_title(content, "title", "Test HTML Document")

    @pytest.mark.integration
    def test_html_yaml_frontmatter_consumes_locale_title_label(self, pipeline, sample_html_file, tmp_path) -> None:
        """HTML YAML frontmatter should consume app-resolved locale labels."""
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_html_locale_yaml"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_html_file,
            "html",
            output_dir,
            yaml_key_labels={"title": "Titel"},
        )

        assert result.success
        content = _assert_finalized_markdown_content(result, output_dir, ws_mgr.root_dir)
        _assert_yaml_title(content, "Titel", "Test HTML Document")

    @pytest.mark.integration
    def test_html_head_title_does_not_leak_into_body(
        self,
        pipeline,
        sample_html_file,
        tmp_path,
    ) -> None:
        """HTML <title> should stay in YAML metadata, not leak into body text."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html_title_body"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file, "html", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        _assert_yaml_title(content, "title", "Test HTML Document")
        body = content.split("---", 2)[2]
        non_empty_body_lines = [line.strip() for line in body.splitlines() if line.strip()]
        assert non_empty_body_lines[0] == "Main Heading"
        assert "# Test HTML Document" not in body

    @pytest.mark.integration
    def test_html_heading_structure(self, pipeline, sample_html_file, tmp_path) -> None:
        """HTML headings must be preserved in Markdown output."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html_headings"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file, "html", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "Main Heading" in content
        assert "Subheading" in content
        assert "# " in content

    @pytest.mark.integration
    def test_html_list_preservation(self, pipeline, sample_html_file, tmp_path) -> None:
        """Ordered and unordered lists must be preserved."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html_lists"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file, "html", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "Item A" in content
        assert "Item B" in content
        assert "First" in content
        assert "Second" in content

    @pytest.mark.integration
    def test_html_link_preservation(self, pipeline, sample_html_file, tmp_path) -> None:
        """Hyperlinks must be preserved."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html_links"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file, "html", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "https://example.com" in content or "link" in content.lower()

    @pytest.mark.integration
    def test_html_table_preservation(self, pipeline, sample_html_file, tmp_path) -> None:
        """Tables must be preserved."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html_tables"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file, "html", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "Name" in content
        assert "Foo" in content
        assert "Bar" in content

    @pytest.mark.integration
    def test_html_artifact_metadata(self, pipeline, sample_html_file, tmp_path) -> None:
        """Artifact must carry correct metadata."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html_meta"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file, "html", output_dir)

        assert result.success
        artifact = result.artifacts[0]
        assert artifact.kind == "primary"
        assert artifact.suggested_name.endswith(".md")
        assert artifact.media_type == "text/markdown"
        assert artifact.is_primary is True
        assert "title" in artifact.metadata

    @pytest.mark.integration
    def test_htm_alias_works(self, pipeline, sample_html_file, tmp_path) -> None:
        """ROUTE-HTM-001: .htm files must convert same as .html."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_htm"
        output_dir.mkdir()

        htm_path = tmp_path / "test_alias.htm"
        htm_path.write_text(sample_html_file.read_text(encoding="utf-8"), encoding="utf-8")

        result = _run_request(task_mgr, htm_path, "htm", output_dir)
        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "Main Heading" in content

    @pytest.mark.integration
    def test_htm_alias_matches_html_old_system_semantic_fixture(self, pipeline, tmp_path) -> None:
        """ROUTE-HTM-001 should preserve the same old-system semantics as HTML."""
        fixture = _load_html_old_system_fixture()
        _plugin, task_mgr, _ws_mgr = pipeline

        htm_path = tmp_path / "sample_alias.htm"
        htm_path.write_text(fixture["input_html"]["body"], encoding="utf-8")
        output_dir = tmp_path / "output_htm_old_system_fixture"
        output_dir.mkdir()

        result = _run_request(task_mgr, htm_path, "htm", output_dir)

        assert result.success, f"HTM conversion failed: {result.error.message if result.error else 'unknown'}"
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        _assert_yaml_title(content, "title", fixture["input_html"]["title"])
        for token in fixture["current_expected_semantics"]["required_markdown_tokens"]:
            assert token in content
        for fragment in fixture["current_expected_semantics"]["forbidden_markdown_fragments"]:
            assert fragment not in content

    @pytest.mark.integration
    def test_html_cancellation(self, pipeline, sample_html_file, tmp_path) -> None:
        """HTML→MD must support cancellation."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html_cancel"
        output_dir.mkdir()

        request = ConversionRequest(
            request_id="html-cancel-test",
            input_refs=[
                FileRef(
                    path=str(sample_html_file),
                    format="html",
                    category="markup",
                    size_bytes=sample_html_file.stat().st_size,
                )
            ],
            target_format="md",
            output_policy=OutputPolicy(output_dir=str(output_dir)),
        )

        task_mgr.cancel("html-cancel-test")
        result = task_mgr.execute_single(request)

        assert result.success is False
        assert result.error is not None
        assert result.error.error_type == "cancelled"

    @pytest.mark.integration
    def test_html_companion_image_is_finalized_and_uses_link_style(
        self,
        pipeline,
        sample_html_file_with_companion_image,
        tmp_path,
    ) -> None:
        """HTML companion-folder images must be finalized and honor Link settings."""
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_html_resources"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_html_file_with_companion_image,
            "html",
            output_dir,
            config_snapshot={"link": {"format": {"image_link_style": "wiki_link"}}},
        )

        assert result.success
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)

        primary_artifacts = [a for a in result.artifacts if a.kind == "primary"]
        assert len(primary_artifacts) == 1
        markdown_path = Path(primary_artifacts[0].staging_path)
        node_root = _document_node_root(markdown_path, output_dir)
        assert markdown_path.exists()

        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        assert len(image_artifacts) == 1
        assert _document_node_root(Path(image_artifacts[0].staging_path), output_dir) == node_root
        assert Path(image_artifacts[0].staging_path).exists()

        content = markdown_path.read_text(encoding="utf-8")
        assert "[[picture.png]]" in content
        assert "![picture.png](picture.png)" not in content
        assert str(Path(ws_mgr.root_dir)) not in content

    @pytest.mark.integration
    def test_html_multi_resource_images_match_old_system_projection(self, pipeline, tmp_path) -> None:
        """Multiple HTML companion images should finalize in order with readable links."""
        fixture = _load_html_old_system_fixture()
        probe = fixture["image_artifact_scope"]["multi_resource_probe"]
        _plugin, task_mgr, ws_mgr = pipeline

        html_path = tmp_path / probe["filename"]
        _write_html_multi_resource_probe(html_path, probe)
        output_dir = tmp_path / "output_html_multi_resources"
        output_dir.mkdir()

        result = _run_request(task_mgr, html_path, "html", output_dir, image_link_style="markdown_embed")

        assert result.success
        expected = probe["current_expected_semantics"]
        assert [d.code for d in result.diagnostics if d.code in expected["diagnostics"]] == expected["diagnostics"]
        assert len(result.artifacts) == expected["artifact_count"] + 1

        markdown_path = Path(result.artifacts[0].staging_path)
        node_root = _document_node_root(markdown_path, output_dir)
        content = markdown_path.read_text(encoding="utf-8")
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
    def test_html_data_uri_image_is_finalized_artifact(
        self,
        pipeline,
        sample_html_file_with_data_uri_image,
        tmp_path,
    ) -> None:
        """HTML data URI images must be materialized through the resource writer."""
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_html_data_uri"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_html_file_with_data_uri_image, "html", output_dir)

        assert result.success
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)

        primary_artifacts = [a for a in result.artifacts if a.kind == "primary"]
        assert len(primary_artifacts) == 1
        markdown_path = Path(primary_artifacts[0].staging_path)
        node_root = _document_node_root(markdown_path, output_dir)
        assert markdown_path.exists()

        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        assert len(image_artifacts) == 1
        image_path = Path(image_artifacts[0].staging_path)
        assert _document_node_root(image_path, output_dir) == node_root
        assert image_path.exists()
        assert image_artifacts[0].suggested_name.endswith(".png")
        assert image_path.read_bytes().startswith(b"\x89PNG\r\n\x1a\n")

        content = markdown_path.read_text(encoding="utf-8")
        assert image_artifacts[0].suggested_name in content
        assert "data:image/png;base64" not in content
        assert str(Path(ws_mgr.root_dir)) not in content

    @pytest.mark.integration
    def test_html_resource_links_honor_request_image_mode_base64(
        self,
        pipeline,
        sample_html_file_with_companion_image,
        tmp_path,
    ) -> None:
        """HTML resources must honor per-request base64 image mode."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_html_base64_mode"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_html_file_with_companion_image,
            "html",
            output_dir,
            image_mode="base64",
            image_link_style="markdown_embed",
        )

        assert result.success
        assert [a.kind for a in result.artifacts].count("image") == 0
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "![picture.png](data:image/png;base64," in content
        assert "](picture.png)" not in content

    @pytest.mark.integration
    def test_html_remote_image_uses_link_style_without_artifact(
        self,
        pipeline,
        sample_html_file_with_remote_image,
        tmp_path,
    ) -> None:
        """HTML remote images must honor Link settings without creating local artifacts."""
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_html_remote"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_html_file_with_remote_image,
            "html",
            output_dir,
            config_snapshot={"link": {"format": {"image_link_style": "wiki_link"}}},
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
        assert "[[https://example.com/image.png]]" in content
        assert "![https://example.com/image.png](https://example.com/image.png)" not in content
        assert str(Path(ws_mgr.root_dir)) not in content
