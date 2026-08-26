"""Focused tests split from test_input_routes.py."""

from __future__ import annotations

from ._input_routes_support import (
    Path,
    _assert_finalized_markdown_content,
    _assert_yaml_title,
    _document_node_root,
    _load_enex_old_system_fixture,
    _run_request,
    _successful_ocr,
    _test_png_base64,
    _test_png_bytes,
    _write_enex_multi_resource_probe,
    base64,
    hashlib,
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

    @pytest.mark.contract
    def test_normalize_enml_strips_xml_envelope_and_keeps_resource_token(self) -> None:
        """Only the ENML body fragment should be sent to markdownify."""
        from docwen_plugin_markup.note_export.converter import EnexToMarkdownConverter

        html = EnexToMarkdownConverter._normalize_enml_to_html(
            """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-note SYSTEM "http://xml.evernote.com/pub/enml2.dtd">
<en-note><p>Body</p><en-media hash="ABCDEFABCDEFABCDEFABCDEFABCDEFAB" type="image/png"/></en-note>"""
        )

        assert "<?xml" not in html
        assert "<!DOCTYPE" not in html
        assert "<en-note" not in html
        assert "</en-note>" not in html
        assert "<p>Body</p>" in html
        assert '<img src="__DOCWEN_RES_abcdefabcdefabcdefabcdefabcdefab__" />' in html

    @pytest.mark.integration
    def test_enex_conversion_succeeds(self, pipeline, sample_enex_file, tmp_path) -> None:
        """ENEX file must convert to Markdown successfully."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_enex_file, "enex", output_dir)

        assert result.success, f"ENEX conversion failed: {result.error.message if result.error else 'unknown'}"
        assert len(result.artifacts) >= 1
        assert result.artifacts[0].is_primary is True

        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert len(content.strip()) > 0
        assert "Test Note Title" in content

    @pytest.mark.integration
    def test_enex_to_md_matches_old_system_semantic_fixture(self, pipeline, tmp_path) -> None:
        """Current ENEX→MD should preserve old-system core note semantics."""
        fixture = _load_enex_old_system_fixture()
        _plugin, task_mgr, ws_mgr = pipeline

        enex_path = tmp_path / fixture["input_enex"]["filename"]
        enex_path.write_text(fixture["input_enex"]["body"], encoding="utf-8")
        output_dir = tmp_path / "output_enex_old_system_fixture"
        output_dir.mkdir()

        result = _run_request(task_mgr, enex_path, "enex", output_dir)

        assert result.success, f"ENEX conversion failed: {result.error.message if result.error else 'unknown'}"
        artifact = result.artifacts[0]
        expected_metadata = fixture["current_expected_semantics"]["artifact_metadata"]
        assert artifact.media_type == expected_metadata["media_type"]
        assert artifact.metadata["note_count"] == expected_metadata["note_count"]
        assert artifact.metadata["resource_count"] == expected_metadata["resource_count"]
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)

        artifact_path = Path(artifact.staging_path)
        _document_node_root(artifact_path, output_dir)
        assert artifact_path.exists()

        content = artifact_path.read_text(encoding="utf-8")
        assert str(Path(ws_mgr.root_dir)) not in content
        _assert_yaml_title(content, "title", fixture["input_enex"]["note_title"])
        for token in fixture["current_expected_semantics"]["required_markdown_tokens"]:
            assert token in content
        for fragment in fixture["current_expected_semantics"]["forbidden_markdown_fragments"]:
            assert fragment not in content

        non_empty_lines = [line.strip() for line in content.splitlines() if line.strip()]
        assert non_empty_lines[0] == fixture["probe_outputs"]["current_first_non_empty_line"]

    @pytest.mark.integration
    def test_enex_yaml_frontmatter_consumes_locale_title_label(self, pipeline, sample_enex_file, tmp_path) -> None:
        """ENEX YAML frontmatter should consume app-resolved locale labels."""
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_locale_yaml"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_enex_file,
            "enex",
            output_dir,
            yaml_key_labels={"title": "Titel"},
        )

        assert result.success
        content = _assert_finalized_markdown_content(result, output_dir, ws_mgr.root_dir)
        _assert_yaml_title(content, "Titel", "Test Note Title")

    @pytest.mark.integration
    def test_enex_output_is_valid_markdown(self, pipeline, sample_enex_file, tmp_path) -> None:
        """Output must be well-formed Markdown with structural markers."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_md"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_enex_file, "enex", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "# " in content, f"No heading found in:\n{content[:300]}"
        lines = [line for line in content.splitlines() if line.strip()]
        assert len(lines) >= 3, f"Expected >=3 lines, got {len(lines)}"

    @pytest.mark.integration
    def test_enex_converts_html_elements(self, pipeline, sample_enex_file, tmp_path) -> None:
        """Bold, italic, and lists must be preserved."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_elements"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_enex_file, "enex", output_dir)

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "**" in content, f"No bold markers found. Content:\n{content[:300]}"
        assert "Item One" in content or "* Item" in content or "- Item" in content

    @pytest.mark.integration
    def test_enex_with_resources(self, pipeline, sample_enex_with_resources, tmp_path) -> None:
        """ENEX embedded resources must be finalized beside the Markdown."""
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_res"
        output_dir.mkdir()
        result = _run_request(task_mgr, sample_enex_with_resources, "enex", output_dir)

        assert result.success
        assert any(d.code == "FINALIZER_DONE" for d in result.diagnostics)
        primary = result.artifacts[0]
        primary_path = Path(primary.staging_path)
        node_root = _document_node_root(primary_path, output_dir)
        assert primary_path.exists()
        content = primary_path.read_text(encoding="utf-8")
        assert "Note with Image" in content
        assert "After image" in content
        assert "test_image.png" in content
        assert str(Path(ws_mgr.root_dir)) not in content

        image_artifacts = [a for a in result.artifacts if a.kind in ("auxiliary", "image")]
        assert len(image_artifacts) >= 1, f"Expected at least 1 auxiliary/image artifact, got {len(image_artifacts)}"
        image_artifact = image_artifacts[0]
        image_path = Path(image_artifact.staging_path)
        assert _document_node_root(image_path, output_dir) == node_root
        assert image_path.name == "test_image.png"
        assert image_path.exists()
        assert image_path.read_bytes().startswith(b"\x89PNG")

    @pytest.mark.integration
    def test_enex_multi_resource_images_match_old_system_projection(self, pipeline, tmp_path) -> None:
        """Multiple ENEX resources should finalize in order with readable links."""
        fixture = _load_enex_old_system_fixture()
        probe = fixture["image_artifact_scope"]["multi_resource_probe"]
        _plugin, task_mgr, ws_mgr = pipeline

        enex_path = tmp_path / probe["filename"]
        _write_enex_multi_resource_probe(enex_path, probe)
        output_dir = tmp_path / "output_enex_multi_resources"
        output_dir.mkdir()

        result = _run_request(task_mgr, enex_path, "enex", output_dir)

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
    def test_enex_resource_adapter_keeps_bytes_and_uniquifies_duplicate_names(self, pipeline, tmp_path) -> None:
        """F-G1-005/F-G1-007: ENEX resources retain bytes and collision-safe links."""
        first_payload = _test_png_bytes((20, 40, 60))
        second_payload = _test_png_bytes((80, 100, 120))
        first_hash = hashlib.md5(first_payload).hexdigest()
        second_hash = hashlib.md5(second_payload).hexdigest()
        source = tmp_path / "duplicate-resource-names.enex"
        source.write_text(
            f"""<?xml version="1.0" encoding="UTF-8"?>
<en-export>
  <note>
    <title>Duplicate resources</title>
    <content><![CDATA[<?xml version="1.0" encoding="UTF-8"?>
      <en-note>
        <en-media hash="{first_hash}" type="image/png" />
        <en-media hash="{second_hash}" type="image/png" />
      </en-note>]]></content>
    <resource>
      <data encoding="base64">{base64.b64encode(first_payload).decode("ascii")}</data>
      <mime>image/png</mime>
      <resource-attributes><file-name>duplicate.png</file-name></resource-attributes>
    </resource>
    <resource>
      <data encoding="base64">{base64.b64encode(second_payload).decode("ascii")}</data>
      <mime>image/png</mime>
      <resource-attributes><file-name>duplicate.png</file-name></resource-attributes>
    </resource>
  </note>
</en-export>
""",
            encoding="utf-8",
        )
        output_dir = tmp_path / "output_duplicate_resources"
        output_dir.mkdir()
        _plugin, task_mgr, ws_mgr = pipeline

        result = _run_request(task_mgr, source, "enex", output_dir)

        assert result.success, result.error
        image_artifacts = [artifact for artifact in result.artifacts if artifact.kind == "image"]
        assert [artifact.suggested_name for artifact in image_artifacts] == [
            "duplicate.png",
            "duplicate-2.png",
        ]
        assert [Path(artifact.staging_path).read_bytes() for artifact in image_artifacts] == [
            first_payload,
            second_payload,
        ]
        markdown = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert markdown.count("![[duplicate.png]]") == 1
        assert markdown.count("![[duplicate-2.png]]") == 1
        assert markdown.index("duplicate.png") < markdown.index("duplicate-2.png")
        assert str(Path(ws_mgr.root_dir)) not in markdown

    @pytest.mark.integration
    def test_enex_resource_links_use_request_snapshot_default(
        self,
        pipeline,
        sample_enex_with_resources,
        tmp_path,
    ) -> None:
        """ENEX resource links must honor request-owned Link settings."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_res_link_style"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_enex_with_resources,
            "enex",
            output_dir,
            config_snapshot={"link": {"format": {"image_link_style": "wiki_link"}}},
        )

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "[[test_image.png]]" in content
        assert "![test_image.png](test_image.png)" not in content

    @pytest.mark.parametrize(
        ("resource_data", "reference_hash", "expected_message"),
        [
            ("   ", "0" * 32, "missing base64 data"),
            ("%%%not-base64%%%", "0" * 32, "invalid base64 data"),
            (_test_png_base64(), "f" * 32, "references missing or corrupt resources"),
        ],
    )
    @pytest.mark.integration
    def test_enex_malformed_or_mismatched_resources_fail_without_output(
        self,
        pipeline,
        tmp_path: Path,
        resource_data: str,
        reference_hash: str,
        expected_message: str,
    ) -> None:
        """Corrupt attachments must not leak internal tokens through a success result."""
        source = tmp_path / "malformed-resource.enex"
        source.write_text(
            f"""<?xml version="1.0" encoding="UTF-8"?>
<en-export><note><title>Malformed resource</title>
<content><![CDATA[<en-note><en-media hash="{reference_hash}" type="image/png"/></en-note>]]></content>
<resource><data encoding="base64">{resource_data}</data><mime>image/png</mime></resource>
</note></en-export>
""",
            encoding="utf-8",
        )
        output_dir = tmp_path / "malformed-resource-output"
        output_dir.mkdir()
        _plugin, task_mgr, _ws_mgr = pipeline

        result = _run_request(task_mgr, source, "enex", output_dir)

        assert result.success is False
        assert result.artifacts == []
        assert result.error is not None
        assert result.error.error_type == "conversion_failed"
        assert result.error.diagnostic_code == "ENEX2MD-PARSE-ERROR"
        assert expected_message in result.error.message
        assert all("__DOCWEN_RES_" not in diagnostic.message for diagnostic in result.diagnostics)
        assert list(output_dir.iterdir()) == []

    @pytest.mark.integration
    def test_enex_resource_links_honor_request_image_link_style(
        self,
        pipeline,
        sample_enex_with_resources,
        tmp_path,
    ) -> None:
        """ENEX resource links must honor per-request Link style overrides."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_res_request_link_style"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_enex_with_resources,
            "enex",
            output_dir,
            image_link_style="markdown_embed",
        )

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "![test_image.png](test_image.png)" in content
        assert "![[test_image.png]]" not in content

    @pytest.mark.integration
    def test_enex_resource_links_honor_request_image_mode_base64(
        self,
        pipeline,
        sample_enex_with_resources,
        tmp_path,
    ) -> None:
        """ENEX resources must honor per-request base64 image mode."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_res_request_base64_mode"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_enex_with_resources,
            "enex",
            output_dir,
            image_mode="base64",
            image_link_style="markdown_embed",
        )

        assert result.success
        assert [a.kind for a in result.artifacts].count("image") == 0
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        assert "![test_image.png](data:image/png;base64," in content
        assert "](test_image.png)" not in content

    @pytest.mark.integration
    def test_enex_markdown_resource_links_use_md_file_style(
        self,
        pipeline,
        sample_enex_with_markdown_resource,
        tmp_path,
    ) -> None:
        """ENEX .md resource links must honor the md_file_link_style setting."""
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_md_res_link_style"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_enex_with_markdown_resource,
            "enex",
            output_dir,
            config_snapshot={"link": {"format": {"md_file_link_style": "wiki_link"}}},
        )

        assert result.success
        content = Path(result.artifacts[0].staging_path).read_text(encoding="utf-8")
        linked_note = next(
            artifact
            for artifact in result.artifacts
            if artifact.media_type == "text/markdown" and not artifact.is_primary
        )
        assert linked_note.logical_path is not None
        relative_link = linked_note.logical_path.split("/", 1)[1]
        assert f"[[{relative_link}]]" in content
        assert "[linked_note.md](linked_note.md)" not in content

    @pytest.mark.integration
    def test_enex_ocr_main_md_replaces_resource_token_cleanly(
        self,
        pipeline,
        sample_enex_with_resources,
        tmp_path,
        monkeypatch,
    ) -> None:
        """ENEX OCR inline output must replace the whole generated image token."""
        monkeypatch.setattr(
            "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
            lambda _path, **_kwargs: _successful_ocr("OCR text from ENEX image"),
        )
        _plugin, task_mgr, _ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_ocr_main"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_enex_with_resources,
            "enex",
            output_dir,
            to_md_enable_ocr=True,
            ocr_placement="main_md",
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
        image_path = Path(image_artifacts[0].staging_path)
        assert _document_node_root(image_path, output_dir) == node_root
        assert image_path.exists()
        assert [a for a in result.artifacts if a.kind == "auxiliary"] == []

        content = markdown_path.read_text(encoding="utf-8")
        assert "> OCR text from ENEX image" in content
        assert "__DOCWEN_RES_" not in content
        assert "![](>" not in content
        assert str(Path(_ws_mgr.root_dir)) not in content

    @pytest.mark.integration
    def test_enex_ocr_image_md_creates_sidecar(
        self,
        pipeline,
        sample_enex_with_resources,
        tmp_path,
        monkeypatch,
    ) -> None:
        """ENEX image_md OCR must finalize the image and Markdown sidecar."""
        monkeypatch.setattr(
            "docwen_plugin_markup.markdown_resources.run_ocr_outcome",
            lambda _path, **_kwargs: _successful_ocr("OCR text from ENEX image"),
        )
        _plugin, task_mgr, ws_mgr = pipeline

        output_dir = tmp_path / "output_enex_ocr_sidecar"
        output_dir.mkdir()
        result = _run_request(
            task_mgr,
            sample_enex_with_resources,
            "enex",
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
        assert "test_image.png" in sidecar_content
        assert "> OCR text from ENEX image" in sidecar_content
        assert str(Path(ws_mgr.root_dir)) not in sidecar_content
