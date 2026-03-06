"""docwen.converter.shared.image_md 的单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.shared.image_md import create_image_md_file, process_image_with_ocr
from docwen.utils import ocr_utils

pytestmark = pytest.mark.unit


def test_create_image_md_file_with_image_link_only(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    (tmp_path / "image_1.png").write_bytes(b"fake")

    md_filename = create_image_md_file(
        image_path=str(tmp_path / "image_1.png"),
        image_filename="image_1.png",
        output_folder=str(tmp_path),
        include_image=True,
        include_ocr=False,
        image_link_style="markdown_embed",
    )

    assert md_filename == "image_1.md"
    content = (tmp_path / md_filename).read_text(encoding="utf-8")
    assert "![image_1.png](image_1.png)" in content


def test_process_image_with_ocr_ocr_only_creates_md_and_returns_md_link(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda _p, _c=None: "OCR_RESULT", raising=True)

    (tmp_path / "image_1.png").write_bytes(b"fake")
    img_info = {"filename": "image_1.png", "image_path": str(tmp_path / "image_1.png")}

    out = process_image_with_ocr(
        img=img_info,
        keep_images=False,
        enable_ocr=True,
        output_folder=str(tmp_path),
        current_index=1,
        total_images=1,
        image_link_style="markdown_embed",
        md_file_link_style="markdown_link",
    )

    assert out == "[image_1.md](image_1.md)"
    content = (tmp_path / "image_1.md").read_text(encoding="utf-8")
    assert "OCR_RESULT" in content


def test_process_image_with_ocr_main_md_uses_i18n_prefix_when_config_missing(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda _p, _c=None: "A\nB", raising=True)

    (tmp_path / "image_1.png").write_bytes(b"fake")
    img_info = {"filename": "image_1.png", "image_path": str(tmp_path / "image_1.png")}

    out = process_image_with_ocr(
        img=img_info,
        keep_images=False,
        enable_ocr=True,
        output_folder=str(tmp_path),
        ocr_placement_mode="main_md",
        ocr_blockquote_title="PREFIX",
    )

    assert out == "> PREFIX\n> A\n> B"


def test_process_image_with_ocr_main_md_allows_disabling_prefix(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(ocr_utils, "extract_text_simple", lambda _p, _c=None: "A", raising=True)

    (tmp_path / "image_1.png").write_bytes(b"fake")
    img_info = {"filename": "image_1.png", "image_path": str(tmp_path / "image_1.png")}

    out = process_image_with_ocr(
        img=img_info,
        keep_images=False,
        enable_ocr=True,
        output_folder=str(tmp_path),
        ocr_placement_mode="main_md",
    )

    assert out == "> A"


def test_create_image_md_file_uniquifies_when_exists(tmp_path: Path) -> None:
    output_folder = tmp_path

    (output_folder / "image_1.md").write_text("old", encoding="utf-8")

    md_filename = create_image_md_file(
        image_path=str(output_folder / "image_1.png"),
        image_filename="image_1.png",
        output_folder=str(output_folder),
        include_image=False,
        include_ocr=False,
    )

    assert md_filename == "image_1_001.md"
    assert (output_folder / "image_1.md").read_text(encoding="utf-8") == "old"
    assert (output_folder / "image_1_001.md").exists() is True
