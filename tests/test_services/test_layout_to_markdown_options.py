"""services 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.config.config_manager import config_manager
from docwen.services.strategies.layout.to_markdown import LayoutToMarkdownStrategy
from docwen.services.strategies.layout.utils import LayoutPreprocessResult
from docwen.utils.workspace_manager import IntermediateItem

pytestmark = pytest.mark.unit


def test_layout_to_markdown_sets_default_md_modes_and_passes_through(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    src = tmp_path / "a.pdf"
    src.write_bytes(b"%PDF-1.4\n%fake\n")

    work = tmp_path / "work"
    work.mkdir()
    converted = work / "converted.pdf"
    converted.write_bytes(b"%PDF-1.4\n%fake\n")

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    def fake_preprocess_layout_file(
        _file_path: str,
        _temp_dir: str,
        _cancel_event=None,
        actual_format: str | None = None,
    ) -> LayoutPreprocessResult:
        assert actual_format == "pdf"
        return LayoutPreprocessResult(
            processed_file=str(converted),
            actual_format="pdf",
            intermediates=[IntermediateItem(kind="layout_pdf", path=str(converted))],
            preprocess_chain=["pdf"],
        )

    monkeypatch.setattr(
        "docwen.services.strategies.layout.to_markdown.preprocess_layout_file",
        fake_preprocess_layout_file,
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_a, **_k: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.path_utils.generate_output_path",
        lambda *_a, **_k: str(out_dir / "fixed.md"),
        raising=True,
    )
    monkeypatch.setattr(config_manager, "get_layout_to_md_image_extraction_mode", lambda: "base64", raising=True)
    monkeypatch.setattr(config_manager, "get_layout_to_md_ocr_placement_mode", lambda: "main_md", raising=True)

    def fake_extract_pdf_with_pymupdf4llm(
        _pdf_path: str,
        _extract_images: bool,
        _extract_ocr: bool,
        output_dir: str,
        basename_for_output: str,
        _cancel_event=None,
        _progress_callback=None,
        *,
        ocr_placement_mode: str,
        extraction_mode: str,
    ) -> dict:
        assert ocr_placement_mode == "main_md"
        assert extraction_mode == "base64"
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

    strategy = LayoutToMarkdownStrategy()
    result = strategy.execute(
        file_path=str(src),
        options={"actual_format": "pdf", "extract_image": True, "extract_ocr": True},
    )

    assert result.success is True
    assert result.output_path is not None
    assert Path(result.output_path).exists()


def test_layout_to_markdown_respects_explicit_md_modes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    src = tmp_path / "a.pdf"
    src.write_bytes(b"%PDF-1.4\n%fake\n")

    work = tmp_path / "work"
    work.mkdir()
    converted = work / "converted.pdf"
    converted.write_bytes(b"%PDF-1.4\n%fake\n")

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    def fake_preprocess_layout_file(
        _file_path: str,
        _temp_dir: str,
        _cancel_event=None,
        actual_format: str | None = None,
    ) -> LayoutPreprocessResult:
        assert actual_format == "pdf"
        return LayoutPreprocessResult(
            processed_file=str(converted),
            actual_format="pdf",
            intermediates=[],
            preprocess_chain=["pdf"],
        )

    monkeypatch.setattr(
        "docwen.services.strategies.layout.to_markdown.preprocess_layout_file",
        fake_preprocess_layout_file,
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_a, **_k: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.path_utils.generate_output_path",
        lambda *_a, **_k: str(out_dir / "fixed.md"),
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
        *,
        ocr_placement_mode: str,
        extraction_mode: str,
    ) -> dict:
        assert ocr_placement_mode == "image_md"
        assert extraction_mode == "file"
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

    strategy = LayoutToMarkdownStrategy()
    result = strategy.execute(
        file_path=str(src),
        options={
            "actual_format": "pdf",
            "extract_image": True,
            "extract_ocr": True,
            "to_md_image_extraction_mode": "file",
            "to_md_ocr_placement_mode": "image_md",
        },
    )

    assert result.success is True
    assert result.output_path is not None
    assert Path(result.output_path).exists()


def test_layout_to_markdown_invoice_cn_uses_invoice_parser_after_preprocess(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    src = tmp_path / "a.xps"
    src.write_bytes(b"PK\x03\x04\n%fake\n")

    work = tmp_path / "work"
    work.mkdir()
    converted = work / "converted.pdf"
    converted.write_bytes(b"%PDF-1.4\n%fake\n")

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    def fake_preprocess_layout_file(
        _file_path: str,
        _temp_dir: str,
        _cancel_event=None,
        actual_format: str | None = None,
    ) -> LayoutPreprocessResult:
        assert actual_format == "xps"
        return LayoutPreprocessResult(
            processed_file=str(converted),
            actual_format="xps",
            intermediates=[],
            preprocess_chain=["xps", "pdf"],
        )

    monkeypatch.setattr(
        "docwen.services.strategies.layout.to_markdown.preprocess_layout_file",
        fake_preprocess_layout_file,
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_a, **_k: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.path_utils.generate_output_path",
        lambda *_a, **_k: str(out_dir / "fixed.md"),
        raising=True,
    )

    def fake_convert_invoice_layout_to_md(
        *,
        file_path: str,
        actual_format: str,
        output_dir: str,
        basename_for_output: str,
        original_file_stem: str | None = None,
        cancel_event=None,
        progress_callback=None,
    ) -> dict:
        assert file_path == str(converted)
        assert actual_format == "pdf"
        assert original_file_stem == "a"

        folder = Path(output_dir) / basename_for_output
        folder.mkdir(parents=True, exist_ok=True)
        md_path = folder / f"{basename_for_output}.md"
        md_path.write_text("ok", encoding="utf-8")
        return {"md_path": str(md_path), "folder_path": str(folder)}

    monkeypatch.setattr(
        "docwen.converter.layout2md.invoice_parser.convert_invoice_layout_to_md",
        fake_convert_invoice_layout_to_md,
        raising=True,
    )

    def fail_pdf2md(*_a, **_k):
        raise AssertionError("PDF→MD 不应被调用（应走发票解析）")

    monkeypatch.setattr("docwen.converter.pdf2md.extract_pdf_with_pymupdf4llm", fail_pdf2md, raising=True)

    strategy = LayoutToMarkdownStrategy()
    result = strategy.execute(
        file_path=str(src),
        options={"actual_format": "xps", "optimize_for_type": "invoice_cn"},
    )

    assert result.success is True
    assert result.output_path is not None
    assert Path(result.output_path).exists()
