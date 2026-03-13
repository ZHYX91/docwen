"""services 单元测试。"""

from __future__ import annotations

from pathlib import Path

import pytest

from docwen.converter.formats import layout as layout_mod
from docwen.services.strategies.layout.utils import preprocess_layout_file
from docwen.utils.workspace_manager import IntermediateItem


pytestmark = pytest.mark.unit


@pytest.mark.unit
def test_preprocess_layout_file_pdf_creates_copy(tmp_path: Path) -> None:
    src = tmp_path / "a.pdf"
    src.write_bytes(b"%PDF-1.4\n%fake\n")

    work = tmp_path / "work"
    work.mkdir()

    result = preprocess_layout_file(str(src), str(work), actual_format="pdf")
    out_pdf = Path(result.processed_file)
    assert result.actual_format == "pdf"
    assert result.intermediates == []
    assert result.preprocess_chain == ["pdf"]
    assert out_pdf.parent == work
    assert out_pdf.name == "input.pdf"
    assert out_pdf.read_bytes() == src.read_bytes()


@pytest.mark.unit
def test_preprocess_layout_file_xps_converts(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    src = tmp_path / "a.xps"
    src.write_bytes(b"fake")

    work = tmp_path / "work"
    work.mkdir()

    def fake_xps_to_pdf(xps_path: str, _cancel_event=None, output_dir: str | None = None) -> str:
        assert Path(xps_path).parent == work
        assert Path(xps_path).name == "input.xps"
        assert output_dir == str(work)

        out = Path(output_dir) / "converted_xps.pdf"
        out.write_bytes(b"%PDF-1.4\n%fake\n")
        return str(out)

    monkeypatch.setattr(layout_mod, "xps_to_pdf", fake_xps_to_pdf, raising=True)

    result = preprocess_layout_file(str(src), str(work), actual_format="xps")
    assert result.actual_format == "xps"
    assert Path(result.processed_file).name == "converted_xps.pdf"
    assert result.intermediates == [IntermediateItem(kind="layout_pdf", path=result.processed_file)]
    assert result.preprocess_chain == ["xps", "pdf"]


@pytest.mark.unit
def test_preprocess_layout_file_ofd_converts(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    src = tmp_path / "a.ofd"
    src.write_bytes(b"fake")

    work = tmp_path / "work"
    work.mkdir()

    def fake_ofd_to_pdf(ofd_path: str, _cancel_event=None, output_dir: str | None = None) -> str:
        assert Path(ofd_path).parent == work
        assert Path(ofd_path).name == "input.ofd"
        assert output_dir == str(work)

        out = Path(output_dir) / "converted_ofd.pdf"
        out.write_bytes(b"%PDF-1.4\n%fake\n")
        return str(out)

    monkeypatch.setattr(layout_mod, "ofd_to_pdf", fake_ofd_to_pdf, raising=True)

    result = preprocess_layout_file(str(src), str(work), actual_format="ofd")
    assert result.actual_format == "ofd"
    assert Path(result.processed_file).name == "converted_ofd.pdf"
    assert result.intermediates == [IntermediateItem(kind="layout_pdf", path=result.processed_file)]
    assert result.preprocess_chain == ["ofd", "pdf"]


@pytest.mark.unit
def test_preprocess_layout_file_caj_not_implemented_raises_runtime_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    src = tmp_path / "a.caj"
    src.write_bytes(b"fake")

    work = tmp_path / "work"
    work.mkdir()

    def fake_caj_to_pdf(_caj_path: str, _cancel_event=None, output_dir: str | None = None) -> str:
        raise NotImplementedError("not implemented")

    monkeypatch.setattr(layout_mod, "caj_to_pdf", fake_caj_to_pdf, raising=True)

    with pytest.raises(RuntimeError) as exc:
        preprocess_layout_file(str(src), str(work), actual_format="caj")

    assert "CAJ" in str(exc.value)


@pytest.mark.unit
def test_preprocess_layout_file_unsupported_format_raises_runtime_error(tmp_path: Path) -> None:
    src = tmp_path / "a.bin"
    src.write_bytes(b"fake")

    work = tmp_path / "work"
    work.mkdir()

    with pytest.raises(RuntimeError) as exc:
        preprocess_layout_file(str(src), str(work), actual_format="unknownfmt")

    assert "unknownfmt" in str(exc.value)
