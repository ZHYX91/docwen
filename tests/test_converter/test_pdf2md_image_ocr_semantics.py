"""converter 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.config.config_manager import config_manager
from docwen.converter.pdf2md import core as pdf2md_core
from docwen.utils import ocr_utils

pytestmark = pytest.mark.unit


def test_convert_to_simple_paths_keeps_images(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"image_link_style": "markdown_embed"},
        raising=True,
    )

    md_text = "before ![alt](C:/temp/image_1.png) after"
    out = pdf2md_core._convert_to_simple_paths(md_text)

    assert "![image_1.png](image_1.png)" in out
    assert "C:/temp" not in out


@pytest.mark.parametrize(
    "md_text",
    [
        'before ![alt](C:/temp/image_1.png "title") after',
        'before ![alt](<C:/temp/image_1.png> "title") after',
    ],
)
def test_convert_to_simple_paths_handles_title_variants(monkeypatch: pytest.MonkeyPatch, md_text: str) -> None:
    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"image_link_style": "markdown_embed"},
        raising=True,
    )

    out = pdf2md_core._convert_to_simple_paths(md_text)

    assert "![image_1.png](image_1.png)" in out
    assert "C:/temp" not in out


def test_add_ocr_after_images_replaces_image_with_md_link_and_creates_image_md(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"md_file_link_style": "markdown_link", "image_link_style": "markdown_embed"},
        raising=True,
    )
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda _p, _c=None: "OCR_RESULT", raising=True)

    (tmp_path / "image_1.png").write_bytes(b"fake")

    md_text = "before ![alt](image_1.png) after"
    out, ocr_count = pdf2md_core._add_ocr_after_images(md_text, str(tmp_path))

    assert ocr_count == 1
    assert "[image_1.md](image_1.md)" in out
    assert "image_1.png" not in out

    image_md = tmp_path / "image_1.md"
    assert image_md.exists()

    content = image_md.read_text(encoding="utf-8")
    assert "![image_1.png](image_1.png)" in content
    assert "OCR_RESULT" in content


def test_add_ocr_after_images_handles_title_variant(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"md_file_link_style": "markdown_link", "image_link_style": "markdown_embed"},
        raising=True,
    )
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda _p, _c=None: "OCR_RESULT", raising=True)

    (tmp_path / "image_1.png").write_bytes(b"fake")

    md_text = 'before ![alt](image_1.png "title") after'
    out, ocr_count = pdf2md_core._add_ocr_after_images(md_text, str(tmp_path))

    assert ocr_count == 1
    assert "[image_1.md](image_1.md)" in out
    assert "image_1.png" not in out


def test_count_images_in_folder_counts_more_extensions(tmp_path: Path) -> None:
    (tmp_path / "a.png").write_bytes(b"fake")
    (tmp_path / "b.gif").write_bytes(b"fake")
    (tmp_path / "c.tif").write_bytes(b"fake")
    (tmp_path / "not_image.txt").write_text("x", encoding="utf-8")

    assert pdf2md_core._count_images_in_folder(str(tmp_path)) == 3


def test_replace_images_with_ocr_blockquote_replaces_image_with_blockquote(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(config_manager, "get_ocr_blockquote_title_enabled", lambda: True, raising=True)
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda _p, _c=None: "OCR_LINE", raising=True)

    (tmp_path / "image_1.png").write_bytes(b"fake")

    md_text = "before ![alt](image_1.png) after"
    out, ocr_count = pdf2md_core._replace_images_with_ocr_blockquote(md_text, str(tmp_path))

    assert ocr_count == 1
    assert "image_1.png" not in out
    assert "> OCR_LINE" in out


def test_replace_images_with_ocr_blockquote_can_hide_title_line(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(config_manager, "get_ocr_blockquote_title_enabled", lambda: False, raising=True)
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda _p, _c=None: "OCR_LINE", raising=True)

    (tmp_path / "image_1.png").write_bytes(b"fake")

    md_text = "before ![alt](image_1.png) after"
    out, ocr_count = pdf2md_core._replace_images_with_ocr_blockquote(md_text, str(tmp_path))

    assert ocr_count == 1
    assert "> OCR_LINE" in out


def test_add_ocr_blockquote_after_images_keeps_image_and_appends_blockquote(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(config_manager, "get_ocr_blockquote_title_enabled", lambda: True, raising=True)
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda _p, _c=None: "OCR_LINE", raising=True)
    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"image_link_style": "markdown_embed"},
        raising=True,
    )

    (tmp_path / "image_1.png").write_bytes(b"fake")

    md_text = "before ![alt](image_1.png) after"
    out, ocr_count = pdf2md_core._add_ocr_blockquote_after_images(md_text, str(tmp_path))

    assert ocr_count == 1
    assert "![image_1.png](image_1.png)" in out
    assert "> OCR_LINE" in out


def test_convert_images_to_base64_converts_multiple_link_styles(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    (tmp_path / "image_1.png").write_bytes(b"fake")

    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"image_link_style": "markdown_embed"},
        raising=True,
    )

    md_text = "\n".join(["![alt](image_1.png)", "[doc](doc.md)"])
    out, _ = pdf2md_core._apply_image_rules(
        md_text,
        str(tmp_path),
        keep_images=True,
        enable_ocr=False,
        extraction_mode="base64",
    )

    assert "data:image/" in out
    assert "[doc](doc.md)" in out


def test_add_ocr_after_images_supports_base64_image_md(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        config_manager,
        "get_markdown_link_style_settings",
        lambda: {"md_file_link_style": "markdown_link", "image_link_style": "markdown_embed"},
        raising=True,
    )
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda _p, _c=None: "OCR_RESULT", raising=True)

    (tmp_path / "image_1.png").write_bytes(b"fake")

    md_text = "before ![alt](image_1.png) after"
    out, ocr_count = pdf2md_core._add_ocr_after_images(
        md_text,
        str(tmp_path),
        extraction_mode="base64",
    )

    assert ocr_count == 1
    assert "[image_1.md](image_1.md)" in out

    image_md = tmp_path / "image_1.md"
    assert image_md.exists()

    content = image_md.read_text(encoding="utf-8")
    assert "data:image/" in content
    assert "OCR_RESULT" in content
