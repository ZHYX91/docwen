"""services 单元测试。"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from docwen.config.config_manager import config_manager
from docwen.services.strategies.layout.to_markdown import LayoutToMarkdownStrategy
from docwen.services.strategies.layout.utils import LayoutPreprocessResult
from docwen.utils.workspace_manager import IntermediateItem


pytestmark = pytest.mark.unit

@pytest.mark.unit
def test_layout_to_markdown_keeps_intermediate_pdf_and_writes_manifest(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    src = tmp_path / "a.ofd"
    src.write_bytes(b"fake")

    work = tmp_path / "work"
    work.mkdir()
    converted = work / "converted.pdf"
    converted.write_bytes(b"%PDF-1.4\n%fake\n")

    def fake_preprocess_layout_file(
        _file_path: str,
        _temp_dir: str,
        _cancel_event=None,
        actual_format: str | None = None,
    ) -> LayoutPreprocessResult:
        assert actual_format == "ofd"
        return LayoutPreprocessResult(
            processed_file=str(converted),
            actual_format="ofd",
            intermediates=[IntermediateItem(kind="layout_pdf", path=str(converted))],
            preprocess_chain=["ofd", "pdf"],
        )

    monkeypatch.setattr(
        "docwen.services.strategies.layout.to_markdown.preprocess_layout_file",
        fake_preprocess_layout_file,
        raising=True,
    )

    def fake_extract_pdf_with_pymupdf4llm(
        _pdf_path: str,
        _extract_images: bool,
        _extract_ocr: bool,
        output_dir: str,
        basename_for_output: str,
        _cancel_event=None,
        _progress_callback=None,
        **_kwargs,
    ) -> dict:
        folder = Path(output_dir) / basename_for_output
        folder.mkdir(parents=True, exist_ok=True)
        md_path = folder / f"{basename_for_output}.md"
        md_path.write_text("ok", encoding="utf-8")
        return {
            "md_path": str(md_path),
            "folder_path": str(folder),
            "image_count": 0,
            "ocr_count": 0,
        }

    monkeypatch.setattr(
        "docwen.converter.pdf2md.extract_pdf_with_pymupdf4llm",
        fake_extract_pdf_with_pymupdf4llm,
        raising=True,
    )

    monkeypatch.setattr(config_manager, "get_save_intermediate_files", lambda: True, raising=True)
    monkeypatch.setattr(config_manager, "get_save_manifest", lambda: True, raising=True)
    monkeypatch.setattr(config_manager, "get_mask_manifest_input_path", lambda: True, raising=True)

    strategy = LayoutToMarkdownStrategy()
    result = strategy.execute(
        file_path=str(src),
        options={"actual_format": "ofd", "extract_image": False, "extract_ocr": False},
    )

    assert result.success is True
    assert result.output_path is not None

    output_md = Path(result.output_path)
    assert output_md.exists()
    output_folder = output_md.parent

    manifest_path = output_folder / "manifest.json"
    assert manifest_path.exists()

    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["schema_version"] == 1
    assert isinstance(manifest["generated_at"], str) and manifest["generated_at"]
    assert manifest["input"]["path"].startswith("<redacted>")
    assert manifest["actual_format"] == "ofd"
    assert manifest["preprocess_chain"] == ["ofd", "pdf"]
    assert manifest["intermediates"][0]["kind"] == "layout_pdf"
    assert Path(manifest["intermediates"][0]["path"]).name == "converted.pdf"
    assert manifest["result"]["success"] is True
    assert Path(manifest["result"]["output_path"]).name.endswith(".md")

    saved_intermediate = tmp_path / "converted.pdf"
    assert saved_intermediate.exists()


@pytest.mark.unit
def test_layout_to_markdown_writes_failure_manifest_on_import_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    src = tmp_path / "a.ofd"
    src.write_bytes(b"fake")

    work = tmp_path / "work"
    work.mkdir()
    converted = work / "converted.pdf"
    converted.write_bytes(b"%PDF-1.4\n%fake\n")

    def fake_preprocess_layout_file(
        _file_path: str,
        _temp_dir: str,
        _cancel_event=None,
        actual_format: str | None = None,
    ) -> LayoutPreprocessResult:
        assert actual_format == "ofd"
        return LayoutPreprocessResult(
            processed_file=str(converted),
            actual_format="ofd",
            intermediates=[IntermediateItem(kind="layout_pdf", path=str(converted))],
            preprocess_chain=["ofd", "pdf"],
        )

    monkeypatch.setattr(
        "docwen.services.strategies.layout.to_markdown.preprocess_layout_file",
        fake_preprocess_layout_file,
        raising=True,
    )

    def fake_extract_pdf_with_pymupdf4llm(*_args, **_kwargs):
        raise ImportError("missing dep")

    monkeypatch.setattr(
        "docwen.converter.pdf2md.extract_pdf_with_pymupdf4llm",
        fake_extract_pdf_with_pymupdf4llm,
        raising=True,
    )

    monkeypatch.setattr(config_manager, "get_save_manifest", lambda: True, raising=True)
    monkeypatch.setattr(config_manager, "get_mask_manifest_input_path", lambda: True, raising=True)
    monkeypatch.setattr(config_manager, "get_save_intermediate_files", lambda: True, raising=True)

    strategy = LayoutToMarkdownStrategy()
    result = strategy.execute(
        file_path=str(src),
        options={"actual_format": "ofd", "extract_image": False, "extract_ocr": False},
    )

    assert result.success is False
    assert "missing dep" in (result.message or "")

    failed_manifests = sorted(tmp_path.glob("manifest_failed_*.json"))
    assert failed_manifests

    manifest = json.loads(failed_manifests[-1].read_text(encoding="utf-8"))
    assert manifest["schema_version"] == 1
    assert isinstance(manifest["generated_at"], str) and manifest["generated_at"]
    assert manifest["input"]["path"].startswith("<redacted>")
    assert manifest["result"]["success"] is False
