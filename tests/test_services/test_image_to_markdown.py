"""services 单元测试。"""

from __future__ import annotations

import base64
from pathlib import Path

import pytest

from docwen.config.config_manager import config_manager
from docwen.services.strategies.image.to_markdown import ImageToMarkdownStrategy
from docwen.utils import ocr_utils

pytestmark = pytest.mark.unit


def _write_png(path: Path) -> None:
    png = base64.b64decode(
        b"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO6n7WQAAAAASUVORK5CYII="
    )
    path.write_bytes(png)


def test_image_to_markdown_missing_actual_format_is_ok(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    in_file = tmp_path / "a.png"
    _write_png(in_file)

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_args, **_kwargs: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.utils.file_type_utils.detect_actual_file_format",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("x")),
        raising=True,
    )

    strategy = ImageToMarkdownStrategy()
    result = strategy.execute(str(in_file), options={"extract_ocr": False})

    assert result.success is True
    assert result.output_path is not None
    assert Path(result.output_path).exists()
    assert Path(result.output_path).suffix == ".md"


def test_image_to_markdown_output_dir_conflict_does_not_delete_existing(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    in_file = tmp_path / "a.png"
    _write_png(in_file)

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    existing = out_dir / "fixed"
    existing.mkdir()
    sentinel = existing / "keep.txt"
    sentinel.write_text("x", encoding="utf-8")

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_args, **_kwargs: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.services.strategies.image.to_markdown.generate_output_path",
        lambda *_args, **_kwargs: str(out_dir / "fixed.md"),
        raising=True,
    )

    strategy = ImageToMarkdownStrategy()
    result = strategy.execute(str(in_file), options={"extract_ocr": False})

    assert result.success is True
    assert sentinel.exists()
    assert result.output_path is not None
    assert Path(result.output_path).exists()
    assert Path(result.output_path).parent.name != "fixed"


def test_image_to_markdown_base64_mode_inlines_image_and_deletes_files(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    in_file = tmp_path / "a.png"
    _write_png(in_file)

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_args, **_kwargs: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.services.strategies.image.to_markdown.generate_output_path",
        lambda *_args, **_kwargs: str(out_dir / "fixed.md"),
        raising=True,
    )
    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"image_link_style": "markdown_embed", "md_file_link_style": "markdown_link"},
        raising=True,
    )

    strategy = ImageToMarkdownStrategy()
    result = strategy.execute(
        str(in_file),
        options={
            "extract_image": True,
            "extract_ocr": False,
            "to_md_image_extraction_mode": "base64",
            "to_md_ocr_placement_mode": "main_md",
        },
    )

    assert result.success is True
    assert result.output_path is not None

    md_path = Path(result.output_path)
    content = md_path.read_text(encoding="utf-8")
    assert "data:image/" in content

    assert not list(md_path.parent.glob("*.png"))


def test_image_to_markdown_image_md_mode_creates_per_image_md(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    in_file = tmp_path / "a.png"
    _write_png(in_file)

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_args, **_kwargs: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.services.strategies.image.to_markdown.generate_output_path",
        lambda *_args, **_kwargs: str(out_dir / "fixed.md"),
        raising=True,
    )
    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"image_link_style": "markdown_embed", "md_file_link_style": "markdown_link"},
        raising=True,
    )
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda *_a, **_k: "OCR_LINE", raising=True)
    monkeypatch.setattr(config_manager, "get_ocr_blockquote_title_enabled", lambda: False, raising=True)

    strategy = ImageToMarkdownStrategy()
    result = strategy.execute(
        str(in_file),
        options={
            "extract_image": True,
            "extract_ocr": True,
            "to_md_image_extraction_mode": "file",
            "to_md_ocr_placement_mode": "image_md",
        },
    )

    assert result.success is True
    assert result.output_path is not None

    md_path = Path(result.output_path)
    md_files = sorted(md_path.parent.glob("*.md"))
    assert len(md_files) == 2
    assert (md_path.parent / "a.md").exists()

    per_image = (md_path.parent / "a.md").read_text(encoding="utf-8")
    assert "OCR_LINE" in per_image

    assert list(md_path.parent.glob("*.png"))


def test_image_to_markdown_base64_main_md_with_ocr_writes_blockquote(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    in_file = tmp_path / "a.png"
    _write_png(in_file)

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_args, **_kwargs: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.services.strategies.image.to_markdown.generate_output_path",
        lambda *_args, **_kwargs: str(out_dir / "fixed.md"),
        raising=True,
    )
    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"image_link_style": "markdown_embed", "md_file_link_style": "markdown_link"},
        raising=True,
    )
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda *_a, **_k: "OCR_LINE", raising=True)
    monkeypatch.setattr(config_manager, "get_ocr_blockquote_title_enabled", lambda: False, raising=True)

    strategy = ImageToMarkdownStrategy()
    result = strategy.execute(
        str(in_file),
        options={
            "extract_image": True,
            "extract_ocr": True,
            "to_md_image_extraction_mode": "base64",
            "to_md_ocr_placement_mode": "main_md",
        },
    )

    assert result.success is True
    assert result.output_path is not None

    md_path = Path(result.output_path)
    content = md_path.read_text(encoding="utf-8")
    assert "data:image/" in content
    assert "> OCR_LINE" in content

    assert not list(md_path.parent.glob("*.png"))


def test_image_to_markdown_invoice_cn_uses_invoice_ocr_template(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    in_file = tmp_path / "a.png"
    _write_png(in_file)

    out_dir = tmp_path / "out"
    out_dir.mkdir()

    monkeypatch.setattr(
        "docwen.utils.workspace_manager.get_output_directory",
        lambda *_args, **_kwargs: str(out_dir),
        raising=True,
    )
    monkeypatch.setattr(
        "docwen.services.strategies.image.to_markdown.generate_output_path",
        lambda *_args, **_kwargs: str(out_dir / "fixed.md"),
        raising=True,
    )
    monkeypatch.setattr(
        ocr_utils,
        "extract_text_simple",
        lambda *_a, **_k: "发票号码：12345678\n开票日期：2026年1月2日\n价税合计：3.00",
        raising=True,
    )

    strategy = ImageToMarkdownStrategy()
    result = strategy.execute(
        str(in_file),
        options={
            "extract_image": True,
            "extract_ocr": True,
            "to_md_image_extraction_mode": "file",
            "to_md_ocr_placement_mode": "main_md",
            "optimize_for_type": "invoice_cn",
        },
    )

    assert result.success is True
    assert result.output_path is not None

    md_path = Path(result.output_path)
    content = md_path.read_text(encoding="utf-8")
    assert "12345678" in content
