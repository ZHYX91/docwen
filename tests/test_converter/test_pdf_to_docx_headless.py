"""converter 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.formats.layout import external as layout_external


pytestmark = pytest.mark.unit

@pytest.mark.unit
def test_pdf_to_docx_headless_skips_dialog(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"%PDF-1.4\n%fake\n")

    out = tmp_path / "out.docx"

    monkeypatch.setattr(layout_external, "_convert_pdf_with_word", lambda *a, **k: None)
    monkeypatch.setattr(layout_external, "_convert_pdf_with_libreoffice", lambda *a, **k: None)
    monkeypatch.setattr(layout_external, "_convert_pdf_with_pdf2docx", lambda *a, **k: None)

    called = {"dialog": False}

    def fake_dialog(*a, **k) -> None:
        called["dialog"] = True

    monkeypatch.setattr(layout_external, "_show_conversion_failed_dialog", fake_dialog)

    result = layout_external.pdf_to_docx(str(pdf), str(out), headless=True)
    assert result is None
    assert called["dialog"] is False
