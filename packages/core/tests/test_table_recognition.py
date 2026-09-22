"""Tests for offline OCR table-structure recognition."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.unit


def _ocr_outcome():
    from docwen_core.text.ocr import OcrOutcome, OcrStatus, OcrTextRegion

    regions = tuple(
        OcrTextRegion(
            points=((x, y), (x + 40, y), (x + 40, y + 20), (x, y + 20)),
            text=text,
            confidence=0.99,
        )
        for x, y, text in (
            (10, 10, "姓名"),
            (80, 10, "金额"),
            (10, 50, "甲"),
            (80, 50, "100"),
        )
    )
    return OcrOutcome(OcrStatus.SUCCESS, text="姓名\n金额\n甲\n100", regions=regions)


def test_missing_table_model_is_explicitly_unavailable(tmp_path: Path) -> None:
    from docwen_core.text.table_recognition import (
        TableRecognitionStatus,
        recognize_table_markdown,
    )

    image = tmp_path / "table.png"
    image.write_bytes(b"image")

    outcome = recognize_table_markdown(
        image,
        _ocr_outcome(),
        model_path=tmp_path / "missing.onnx",
    )

    assert outcome.status is TableRecognitionStatus.UNAVAILABLE
    assert "model" in outcome.message.lower()


def test_table_recognition_reuses_existing_ocr_geometry(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text._slanet_table as slanet
    from docwen_core.text.table_recognition import (
        TableRecognitionStatus,
        recognize_table_markdown,
    )

    image = tmp_path / "table.png"
    image.write_bytes(b"image")
    model = tmp_path / "slanet-plus.onnx"
    model.write_bytes(b"model")
    observed: dict[str, object] = {}

    def _infer(
        image_path: str | Path,
        model_path: str | Path,
        *,
        boxes: object,
        texts: tuple[str, ...],
        scores: tuple[float, ...],
    ) -> str:
        observed["image_path"] = str(image_path)
        observed["model_path"] = str(model_path)
        observed["boxes"] = boxes
        observed["texts"] = texts
        observed["scores"] = scores
        return "<table><tr><th>姓名</th><th>金额</th></tr><tr><td>甲</td><td>100</td></tr></table>"

    monkeypatch.setattr(slanet, "infer_table_html", _infer)

    outcome = recognize_table_markdown(image, _ocr_outcome(), model_path=model)

    assert outcome.status is TableRecognitionStatus.SUCCESS
    assert outcome.markdown == "| 姓名 | 金额 |\n| --- | --- |\n| 甲 | 100 |"
    assert observed["image_path"] == str(image)
    assert observed["model_path"] == str(model)
    assert observed["texts"] == ("姓名", "金额", "甲", "100")


def test_table_html_merged_cells_are_projected_without_dropping_values() -> None:
    from docwen_core.text.table_recognition import _html_table_to_markdown

    markdown = _html_table_to_markdown(
        "<table>"
        "<tr><th rowspan='2'>项目</th><th colspan='2'>数值</th></tr>"
        "<tr><td>本期</td><td>上期</td></tr>"
        "<tr><td>收入</td><td>10</td><td>9</td></tr>"
        "</table>",
        minimum_matches=2,
    )

    assert "| 项目 | 数值 | 数值 |" in markdown
    assert "| 项目 | 本期 | 上期 |" in markdown
    assert "| 收入 | 10 | 9 |" in markdown


def test_mixed_prose_and_table_does_not_replace_plain_ocr(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import docwen_core.text._slanet_table as slanet
    from docwen_core.text.ocr import OcrOutcome, OcrStatus, OcrTextRegion
    from docwen_core.text.table_recognition import TableRecognitionStatus, recognize_table_markdown

    image = tmp_path / "mixed.png"
    image.write_bytes(b"image")
    model = tmp_path / "slanet-plus.onnx"
    model.write_bytes(b"model")
    values = ("姓名", "金额", "甲", "100", "正文一", "正文二", "正文三", "正文四")
    regions = tuple(
        OcrTextRegion(
            points=((10.0, float(index * 25)), (50.0, float(index * 25)), (50.0, float(index * 25 + 20)), (10.0, float(index * 25 + 20))),
            text=value,
            confidence=0.99,
        )
        for index, value in enumerate(values)
    )
    outcome = OcrOutcome(OcrStatus.SUCCESS, text="\n".join(values), regions=regions)
    monkeypatch.setattr(
        slanet,
        "infer_table_html",
        lambda *_args, **_kwargs: (
            "<table><tr><th>姓名</th><th>金额</th></tr>"
            "<tr><td>甲</td><td>100</td></tr></table>"
        ),
    )

    recognized = recognize_table_markdown(image, outcome, model_path=model)

    assert recognized.status is TableRecognitionStatus.NOT_TABLE
    assert recognized.markdown == ""
