"""Focused tests split from test_input_routes.py."""

from __future__ import annotations

from ._input_routes_support import (
    MIMEBase,
    MIMEImage,
    MIMEMultipart,
    MIMEText,
    Path,
    _assert_finalized_markdown_content,
    _assert_yaml_title,
    _document_node_root,
    _load_mhtml_old_system_fixture,
    _run_request,
    _successful_ocr,
    _test_png_bytes,
    _write_mhtml_from_fixture,
    encoders,
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
    def test_mhtml_to_md_matches_old_system_semantic_fixture(self, pipeline, tmp_path) -> None:
        """Current MHTML→MD should preserve old-system MIME archive semantics."""
        fixture = _load_mhtml_old_system_fixture()
        input_mhtml = fixture["input_mhtml"]
        _plugin, task_mgr, ws_mgr = pipeline

        mhtml_path = tmp_path / input_mhtml["filename"]
        _write_mhtml_from_fixture(mhtml_path, input_mhtml)
        output_dir = tmp_path / "output_mhtml_old_system_fixture"
        output_dir.mkdir()

        result = _run_request(task_mgr, mhtml_path, "mhtml", output_dir)

        assert result.success, f"MHTML conversion failed: {result.error.message if result.error else 'unknown'}"
        expected = fixture["current_expected_semantics"]
        markdown_artifact = result.artifacts[0]
        assert markdown_artifact.media_type == "text/markdown"
        assert markdown_artifact.metadata["image_count"] == expected["image_count"]
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)

        markdown_path = Path(markdown_artifact.staging_path)
        _document_node_root(markdown_path, output_dir)
        assert markdown_path.exists()

        content = markdown_path.read_text(encoding="utf-8")
        assert str(Path(ws_mgr.root_dir)) not in content
        _assert_yaml_title(content, "title", input_mhtml["html_title"])
        for token in expected["required_markdown_tokens"]:
            assert token in content
        for fragment in expected["forbidden_markdown_fragments"]:
            assert fragment not in content

        non_empty_lines = [line.strip() for line in content.splitlines() if line.strip()]
        assert non_empty_lines[:2] == expected["first_non_empty_lines"]
        assert content.count(input_mhtml["html_title"]) == expected["title_occurrences"]

        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        assert len(image_artifacts) == len(expected["expected_image_artifacts"])
        for artifact, expected_artifact in zip(image_artifacts, expected["expected_image_artifacts"], strict=True):
            assert artifact.suggested_name == expected_artifact["suggested_name"]
            assert artifact.media_type == expected_artifact["media_type"]
            assert Path(artifact.staging_path).exists()
            _document_node_root(Path(artifact.staging_path), output_dir)

    @pytest.mark.integration
    def test_mhtml_conversion_succeeds(self, pipeline, sample_mhtml_file, tmp_path) -> None:
        """MHTML file must convert to Markdown successfully."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_mhtml"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_mhtml_file, "mhtml", output_dir)

        assert result.success, f"MHTML conversion failed: {result.error.message if result.error else 'unknown'}"
        assert len(result.artifacts) >= 1

        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert len(content.strip()) > 0
        assert "MHTML Test" in content or "MHTML" in content
        assert "from an MHTML file" in content or "MHTML file" in content.lower()

    @pytest.mark.parametrize(
        "invalid_archive",
        ["forged-bytes", "multipart-without-html", "blank-html"],
    )
    @pytest.mark.integration
    def test_mhtml_without_usable_html_is_a_typed_conversion_failure(
        self,
        pipeline,
        tmp_path,
        invalid_archive: str,
    ) -> None:
        """Malformed archives must not succeed with filename-only YAML."""
        mhtml_path = tmp_path / f"{invalid_archive}.mhtml"
        if invalid_archive == "forged-bytes":
            mhtml_path.write_bytes(b"This is not an MHTML archive.\n")
        elif invalid_archive == "multipart-without-html":
            message = MIMEMultipart("related")
            message.attach(MIMEText("Only a plain-text fallback exists.", "plain", "utf-8"))
            mhtml_path.write_bytes(message.as_bytes())
        else:
            message = MIMEMultipart("related")
            message.attach(MIMEText(" \r\n\t", "html", "utf-8"))
            mhtml_path.write_bytes(message.as_bytes())

        output_dir = tmp_path / f"output_{invalid_archive}"
        output_dir.mkdir()
        _plugin, task_mgr, _ws_mgr = pipeline

        result = _run_request(task_mgr, mhtml_path, "mhtml", output_dir)

        assert result.success is False
        assert result.artifacts == []
        assert result.error is not None
        assert result.error.error_type == "conversion_failed"
        assert result.error.diagnostic_code == "HTML2MD-PARSE-ERROR"
        assert "usable text/html body" in result.error.message
        assert any(d.code == "HTML2MD-PARSE-ERROR" for d in result.diagnostics)
        assert list(output_dir.iterdir()) == []

    @pytest.mark.integration
    def test_mhtml_skips_blank_html_part_before_later_usable_part(self, pipeline, tmp_path) -> None:
        """A blank alternative must not mask a later usable HTML body."""
        message = MIMEMultipart("alternative")
        message.attach(MIMEText(" \r\n\t", "html", "utf-8"))
        message.attach(
            MIMEText(
                "<html><head><title>Fallback title</title></head><body><p>Usable fallback body.</p></body></html>",
                "html",
                "utf-8",
            )
        )
        mhtml_path = tmp_path / "blank-then-usable.mhtml"
        mhtml_path.write_bytes(message.as_bytes())
        output_dir = tmp_path / "output_blank_then_usable"
        output_dir.mkdir()
        _plugin, task_mgr, _ws_mgr = pipeline

        result = _run_request(task_mgr, mhtml_path, "mhtml", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        _assert_yaml_title(content, "title", "Fallback title")
        assert "Usable fallback body." in content

    @pytest.mark.integration
    def test_mhtml_yaml_frontmatter_consumes_locale_title_label(self, pipeline, tmp_path) -> None:
        """MHTML YAML frontmatter should consume app-resolved locale labels."""
        fixture = _load_mhtml_old_system_fixture()
        input_mhtml = fixture["input_mhtml"]
        _plugin, task_mgr, ws_mgr = pipeline

        mhtml_path = tmp_path / input_mhtml["filename"]
        _write_mhtml_from_fixture(mhtml_path, input_mhtml)
        output_dir = tmp_path / "output_mhtml_locale_yaml"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            mhtml_path,
            "mhtml",
            output_dir,
            yaml_key_labels={"title": "Titel"},
        )

        assert result.success
        content = _assert_finalized_markdown_content(result, output_dir, ws_mgr.root_dir)
        _assert_yaml_title(content, "Titel", input_mhtml["html_title"])

    @pytest.mark.integration
    def test_mht_alias_works(self, pipeline, sample_mhtml_file, tmp_path) -> None:
        """ROUTE-MHT-001: .mht files must convert same as .mhtml."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_mht"
        output_dir.mkdir()

        mht_path = tmp_path / "test_alias.mht"
        mht_path.write_bytes(sample_mhtml_file.read_bytes())

        result = _run_request(task_mgr, mht_path, "mht", output_dir)
        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert len(content.strip()) > 0

    @pytest.mark.integration
    def test_mht_alias_matches_mhtml_old_system_semantic_fixture(self, pipeline, tmp_path) -> None:
        """ROUTE-MHT-001 should preserve the same old-system semantics as MHTML."""
        fixture = _load_mhtml_old_system_fixture()
        input_mhtml = fixture["input_mhtml"]
        _plugin, task_mgr, _ws_mgr = pipeline

        mht_path = tmp_path / "sample_alias.mht"
        _write_mhtml_from_fixture(mht_path, input_mhtml)
        output_dir = tmp_path / "output_mht_old_system_fixture"
        output_dir.mkdir()

        result = _run_request(task_mgr, mht_path, "mht", output_dir)

        assert result.success, f"MHT conversion failed: {result.error.message if result.error else 'unknown'}"
        expected = fixture["current_expected_semantics"]
        assert result.artifacts[0].metadata["image_count"] == expected["image_count"]
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        _assert_yaml_title(content, "title", input_mhtml["html_title"])
        for token in expected["required_markdown_tokens"]:
            assert token in content
        for fragment in expected["forbidden_markdown_fragments"]:
            assert fragment not in content
        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        assert len(image_artifacts) == len(expected["expected_image_artifacts"])

    @pytest.mark.integration
    def test_mhtml_artifact_metadata(self, pipeline, sample_mhtml_file, tmp_path) -> None:
        """MHTML artifact must carry correct metadata."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_mhtml_meta"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_mhtml_file, "mhtml", output_dir)

        assert result.success
        artifact = result.artifacts[0]
        assert artifact.kind == "primary"
        assert artifact.media_type == "text/markdown"
        assert artifact.is_primary is True
        assert "source_format" in artifact.metadata

    @pytest.mark.integration
    def test_mhtml_embedded_image_is_finalized_with_markdown(self, pipeline, sample_mhtml_file, tmp_path) -> None:
        """MHTML embedded resources must be placed beside the final Markdown."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_mhtml_resources"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_mhtml_file, "mhtml", output_dir)

        assert result.success
        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        assert len(image_artifacts) == 1

        md_path = Path(result.artifacts[0].staging_path)
        image_path = Path(image_artifacts[0].staging_path)
        node_root = _document_node_root(md_path, output_dir)
        assert _document_node_root(image_path, output_dir) == node_root
        assert image_path.exists()

        content = md_path.read_text(encoding="utf-8")
        assert image_artifacts[0].suggested_name in content

    @pytest.mark.integration
    def test_mhtml_decodes_title_and_finalizes_only_body_images(self, pipeline, tmp_path) -> None:
        """Real archives must not expose HTML entities or orphan MIME resources."""
        image_bytes = _test_png_bytes()

        message = MIMEMultipart("related")
        html_part = MIMEText(
            """<html><head><title>Entity&nbsp;Title &amp; More</title>
            <link rel="stylesheet" href="https://example.test/site.css"></head>
            <body><h1>Archive body</h1><img src="cid:used-image"></body></html>""",
            "html",
            "utf-8",
        )
        html_part["Content-Location"] = "https://example.test/page"
        message.attach(html_part)

        for content_id, location in (
            ("used-image", "https://example.test/used.png"),
            ("unused-image", "https://example.test/unused.png"),
        ):
            image_part = MIMEImage(image_bytes, _subtype="png")
            image_part["Content-ID"] = f"<{content_id}>"
            image_part["Content-Location"] = location
            message.attach(image_part)

        css_part = MIMEText("body { color: black; }", "css", "utf-8")
        css_part["Content-Location"] = "https://example.test/site.css"
        message.attach(css_part)

        mhtml_path = tmp_path / "entity-and-orphan-resources.mhtml"
        mhtml_path.write_bytes(message.as_bytes())
        output_dir = tmp_path / "output_entity_and_orphan_resources"
        output_dir.mkdir()
        _plugin, task_mgr, _ws_mgr = pipeline

        result = _run_request(
            task_mgr,
            mhtml_path,
            "mhtml",
            output_dir,
            image_link_style="markdown_embed",
            yaml_key_labels={"title": "title"},
        )

        assert result.success
        assert result.artifacts[0].metadata["title"] == "Entity\N{NO-BREAK SPACE}Title & More"
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        _assert_yaml_title(content, "title", "Entity\N{NO-BREAK SPACE}Title & More")
        assert [artifact.kind for artifact in result.artifacts] == ["primary", "image", "manifest"]
        assert result.artifacts[1].suggested_name == "used.png"
        assert "![used.png](used.png)" in content
        node_root = _document_node_root(Path(result.artifacts[0].staging_path), output_dir)
        assert sorted(path.name for path in output_dir.iterdir()) == [node_root.name]
        assert sorted(path.name for path in node_root.iterdir()) == [
            "docwen-node.json",
            f"{node_root.name}.md",
            "used.png",
        ]

    @pytest.mark.integration
    def test_mhtml_uses_html_meta_charset_when_mime_part_omits_it(self, pipeline, tmp_path) -> None:
        """Chromium snapshots may declare the HTML charset only inside the payload."""
        document = (
            '<html><head><meta http-equiv="Content-Type" '
            'content="text/html; charset=windows-1252">'
            "<title>Named characters</title></head>"
            "<body><p>Tantek Çelik and Håkon Wium Lie</p></body></html>"
        ).encode("windows-1252")
        html_part = MIMEBase("text", "html")
        html_part.set_payload(document)
        encoders.encode_base64(html_part)
        html_part["Content-Location"] = "https://example.test/charset"
        message = MIMEMultipart("related")
        message.attach(html_part)

        mhtml_path = tmp_path / "meta-charset.mhtml"
        mhtml_path.write_bytes(message.as_bytes())
        output_dir = tmp_path / "output_meta_charset"
        output_dir.mkdir()
        _plugin, task_mgr, _ws_mgr = pipeline

        result = _run_request(task_mgr, mhtml_path, "mhtml", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "Tantek Çelik and Håkon Wium Lie" in content
        assert "\N{REPLACEMENT CHARACTER}" not in content

    @pytest.mark.integration
    def test_mhtml_multi_resource_images_match_old_system_projection(self, pipeline, tmp_path) -> None:
        """Multiple MHTML cid resources should finalize in order with readable links."""
        fixture = _load_mhtml_old_system_fixture()
        probe = fixture["image_artifact_scope"]["multi_resource_probe"]
        _plugin, task_mgr, ws_mgr = pipeline

        mhtml_path = tmp_path / probe["filename"]
        _write_mhtml_from_fixture(mhtml_path, probe)
        output_dir = tmp_path / "output_mhtml_multi_resources"
        output_dir.mkdir()

        result = _run_request(task_mgr, mhtml_path, "mhtml", output_dir, image_link_style="markdown_embed")

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
    def test_mhtml_embedded_image_uses_request_snapshot_default(
        self,
        pipeline,
        sample_mhtml_file,
        tmp_path,
    ) -> None:
        """MHTML embedded image links must honor request-owned Link settings."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_mhtml_link_style"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_mhtml_file,
            "mhtml",
            output_dir,
            config_snapshot={"link": {"format": {"image_link_style": "wiki_link"}}},
        )

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "[[embedded.png]]" in content
        assert "![embedded.png](embedded.png)" not in content

    @pytest.mark.integration
    def test_mhtml_resource_links_honor_request_image_mode_base64(
        self,
        pipeline,
        sample_mhtml_file,
        tmp_path,
    ) -> None:
        """MHTML resources must honor per-request base64 image mode."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_mhtml_base64_mode"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_mhtml_file,
            "mhtml",
            output_dir,
            image_mode="base64",
            image_link_style="markdown_embed",
        )

        assert result.success
        assert [a.kind for a in result.artifacts].count("image") == 0
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "![embedded.png](data:image/png;base64," in content
        assert "](embedded.png)" not in content

    @pytest.mark.integration
    def test_mhtml_keep_images_false_omits_embedded_image(self, pipeline, sample_mhtml_file, tmp_path) -> None:
        """MHTML to_md_keep_images=False must not emit image artifacts or image links."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_mhtml_without_images"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_mhtml_file, "mhtml", output_dir, to_md_keep_images=False)

        assert result.success
        image_artifacts = [a for a in result.artifacts if a.kind == "image"]
        assert image_artifacts == []

        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "embedded.png" not in content
        assert "cid:embedded-img" not in content
        assert "DOCWENHTMLIMAGE" not in content

    @pytest.mark.integration
    def test_mhtml_ocr_image_md_creates_sidecar(
        self,
        pipeline,
        sample_mhtml_file,
        tmp_path,
        monkeypatch,
    ) -> None:
        """MHTML image_md OCR must emit a Markdown sidecar artifact."""
        monkeypatch.setattr(
            "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
            lambda _path, **_kwargs: _successful_ocr("OCR text from MHTML image"),
        )
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_mhtml_ocr_sidecar"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_mhtml_file,
            "mhtml",
            output_dir,
            to_md_enable_ocr=True,
            ocr_placement="image_md",
        )

        assert result.success
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
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
        assert "embedded.png" in sidecar_content
        assert "> OCR text from MHTML image" in sidecar_content
        assert str(Path(ws_mgr.root_dir)) not in sidecar_content
