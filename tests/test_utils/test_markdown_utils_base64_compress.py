"""utils 单元测试。"""

from __future__ import annotations

import base64
import re
from pathlib import Path

import pytest
from PIL import Image

from docwen.config.config_manager import config_manager
from docwen.converter.shared.data_uri_image import build_base64_image_link

pytestmark = pytest.mark.unit


def _parse_data_uri(md: str) -> tuple[str, bytes]:
    m = re.search(r"\(data:([^;]+);base64,([A-Za-z0-9+/=]+)\)", md)
    assert m is not None
    mime = m.group(1)
    data = base64.b64decode(m.group(2))
    return mime, data


def _make_noise_image(path: Path, *, fmt: str) -> None:
    img = Image.effect_noise((512, 512), 100).convert("RGB")
    img.save(path, format=fmt)


def test_base64_compress_over_threshold_outputs_jpeg_and_respects_threshold(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    src = tmp_path / "a.png"
    _make_noise_image(src, fmt="PNG")
    assert src.stat().st_size > 0

    threshold_kb = 200
    monkeypatch.setattr(config_manager, "get_export_base64_compress_enabled", lambda: True, raising=True)
    monkeypatch.setattr(config_manager, "get_export_base64_compress_threshold_kb", lambda: threshold_kb, raising=True)

    md = build_base64_image_link(str(src), style="markdown_embed")
    mime, data = _parse_data_uri(md)
    assert mime == "image/jpeg"
    assert len(data) <= threshold_kb * 1024


def test_base64_not_over_threshold_keeps_original_bytes_for_jpeg(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    src = tmp_path / "a.jpg"
    _make_noise_image(src, fmt="JPEG")

    threshold_kb = 10_000
    monkeypatch.setattr(config_manager, "get_export_base64_compress_enabled", lambda: True, raising=True)
    monkeypatch.setattr(config_manager, "get_export_base64_compress_threshold_kb", lambda: threshold_kb, raising=True)

    md = build_base64_image_link(str(src), style="markdown_embed")
    mime, data = _parse_data_uri(md)
    assert mime == "image/jpeg"
    assert data == src.read_bytes()


def test_base64_compress_warns_when_still_over_threshold(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    src = tmp_path / "a.png"
    _make_noise_image(src, fmt="PNG")

    threshold_kb = 1
    monkeypatch.setattr(config_manager, "get_export_base64_compress_enabled", lambda: True, raising=True)
    monkeypatch.setattr(config_manager, "get_export_base64_compress_threshold_kb", lambda: threshold_kb, raising=True)

    caplog.set_level("WARNING")
    md = build_base64_image_link(str(src), style="markdown_embed")
    mime, _data = _parse_data_uri(md)
    assert mime == "image/jpeg"
    assert "Base64 图片压缩后仍超出阈值" in caplog.text
