"""Lossless OCR-region accounting and offline table model contracts."""

from __future__ import annotations

import html
from pathlib import Path

import numpy as np
import pytest

from docwen_core.text import table_recognition as tables
from docwen_core.text.ocr import OcrOutcome, OcrStatus, OcrTextRegion

pytestmark = pytest.mark.unit


def _region(x: int, y: int, text: str) -> OcrTextRegion:
    return OcrTextRegion(((x, y), (x + 30, y), (x + 30, y + 15), (x, y + 15)), text, 0.99)


def _outcome() -> OcrOutcome:
    regions = (
        _region(0, 0, "Heading"),
        *(
            _region(x, y, text)
            for x, y, text in ((10, 40, "Name"), (70, 40, "Value"), (10, 80, "Alpha"), (70, 80, "10"))
        ),
        _region(0, 120, "Between"),
        *(
            _region(x, y, text)
            for x, y, text in ((10, 160, "Name"), (70, 160, "Value"), (10, 200, "Beta"), (70, 200, "20"))
        ),
        _region(0, 240, "Footnote"),
    )
    return OcrOutcome(OcrStatus.SUCCESS, "\n".join(r.text for r in regions), regions=regions)


@pytest.fixture
def inference(monkeypatch):
    from docwen_core.text import _slanet_table, _table_layout

    monkeypatch.setattr(tables, "resolve_rapidtable_model", lambda *_: Path("structure.onnx"))
    monkeypatch.setattr(tables, "resolve_table_layout_model", lambda *_: Path("layout.onnx"))
    monkeypatch.setattr(_slanet_table, "_load_image", lambda _: np.zeros((280, 130, 3), dtype=np.uint8))
    monkeypatch.setattr(_table_layout, "detect_table_regions", lambda *_: [(0, 30, 120, 110), (0, 150, 120, 230)])
    calls = []

    def infer(image, model, *, boxes, texts, scores):
        calls.append((image.shape, model, boxes.copy(), texts, scores))
        cells = [html.escape(text) for text in texts]
        return "<table><tr><td>" + "</td><td>".join(cells[:2]) + "</td></tr><tr><td>" + "</td><td>".join(
            cells[2:]
        ) + "</td></tr></table>", frozenset(range(len(texts)))

    monkeypatch.setattr(_slanet_table, "infer_table_structure", infer)
    return _slanet_table, calls


def test_multiple_tables_keep_prose_and_reuse_each_ocr_region_once(inference):
    _, calls = inference
    result = tables.recognize_table_markdown("mixed.png", _outcome())
    assert result.status is tables.TableRecognitionStatus.SUCCESS
    assert result.table_count == 2
    assert [call[3] for call in calls] == [("Name", "Value", "Alpha", "10"), ("Name", "Value", "Beta", "20")]
    assert calls[0][2][0].tolist() == [[10, 10], [40, 10], [40, 25], [10, 25]]
    assert result.markdown == (
        "Heading\n\n| Name | Value |\n| --- | --- |\n| Alpha | 10 |\n\nBetween\n\n"
        "| Name | Value |\n| --- | --- |\n| Beta | 20 |\n\nFootnote"
    )


def test_unmatched_text_inside_detected_table_is_retained(inference, monkeypatch):
    slanet, _ = inference
    monkeypatch.setattr(
        slanet,
        "infer_table_structure",
        lambda *_, **__: (
            "<table><tr><td>Name</td><td>Value</td></tr><tr><td>Alpha</td><td></td></tr></table>",
            frozenset({0, 1, 2}),
        ),
    )
    result = tables.recognize_table_markdown("mixed.png", _outcome())
    assert "\n\n10\n\nBetween" in result.markdown
    assert "\n\n20\n\nFootnote" in result.markdown


def test_failed_crop_falls_back_without_discarding_other_table(inference, monkeypatch):
    slanet, _ = inference
    original = slanet.infer_table_structure

    def partial(*args, **kwargs):
        if "Beta" in kwargs["texts"]:
            raise RuntimeError("bad structure")
        return original(*args, **kwargs)

    monkeypatch.setattr(slanet, "infer_table_structure", partial)
    result = tables.recognize_table_markdown("mixed.png", _outcome())
    assert result.table_count == 1 and result.message
    assert "Between\n\nName\n\nValue\n\nBeta\n\n20\n\nFootnote" in result.markdown


def test_cancel_during_table_inference_is_not_converted_to_fallback(inference):
    from docwen_core.errors import CancellationRequested

    calls = 0

    def check():
        nonlocal calls
        calls += 1
        if calls == 4:
            raise CancellationRequested()

    with pytest.raises(CancellationRequested):
        tables.recognize_table_markdown("mixed.png", _outcome(), check_cancelled=check)


def test_disabled_recognition_never_loads_models(monkeypatch):
    monkeypatch.setattr(tables, "resolve_rapidtable_model", lambda *_: pytest.fail("disabled model lookup"))
    enriched, result = tables.enrich_ocr_table_structure("mixed.png", _outcome(), enabled=False)
    assert result.status is tables.TableRecognitionStatus.DISABLED
    assert enriched == _outcome()


@pytest.mark.parametrize("payload", [None, b"wrong model"])
def test_missing_or_corrupt_explicit_model_is_unavailable(tmp_path, payload):
    path = tmp_path / "model.onnx"
    if payload is not None:
        path.write_bytes(payload)
    result = tables.recognize_table_markdown("mixed.png", _outcome(), model_path=path, layout_model_path=path)
    assert result.status is tables.TableRecognitionStatus.UNAVAILABLE
    assert result.markdown == ""


@pytest.mark.parametrize("strategy,covered", [("fill", "Value"), ("empty", ""), ("marker", "&lt;")])
def test_merged_cells_use_existing_export_strategy(strategy, covered):
    result = tables._html_table_to_markdown(
        "<table><tr><td rowspan='2'>Item</td><td colspan='2'>Value</td></tr>"
        "<tr><td>Current</td><td>Previous</td></tr><tr><td>Income</td><td>10</td><td>9</td></tr></table>",
        merge_strategy=strategy,
    )
    assert f"| Item | Value | {covered} |" in result
    assert "| Income | 10 | 9 |" in result


def test_html_like_ocr_is_literal_text_not_markup():
    from docwen_core.text._slanet_table import _fill_structure

    raw = _fill_structure(
        ["<table><tr>", "<td></td>", "<td></td>", "</tr><tr>", "<td></td>", "<td></td>", "</tr></table>"],
        {0: [0], 1: [1], 2: [2], 3: [3]},
        [("Value <alpha>", 1.0), ("literal &lt;tag&gt;", 1.0), ("A|B", 1.0), ("10", 1.0)],
    )
    result = tables._html_table_to_markdown(raw)
    assert "Value &lt;alpha&gt;" in result
    assert "literal &amp;lt;tag&amp;gt;" in result
    assert "A\\|B" in result


def test_borderless_candidates_require_aligned_multicolumn_rows():
    from docwen_core.text._table_layout import add_borderless_candidates

    regions = [_region(x, y, "text") for y in (20, 60, 100) for x in (10, 70)]
    boxes = np.asarray([r.points for r in regions])
    assert len(add_borderless_candidates([], boxes, (200, 200))) == 1
    assert add_borderless_candidates([], boxes[::2], (200, 200)) == []
    assert add_borderless_candidates([(0, 0, 150, 150)], boxes, (200, 200)) == [(0, 0, 150, 150)]


@pytest.mark.parametrize(
    "configured,requested,expected",
    [(True, None, True), (False, None, False), (False, True, True), (True, False, False)],
)
def test_request_option_overrides_frozen_config(configured, requested, expected):
    from types import SimpleNamespace

    options = {} if requested is None else {"recognize_tables": requested}
    context = SimpleNamespace(
        request=SimpleNamespace(options=options, config_snapshot={"ocr": {"recognize_tables": configured}})
    )
    assert tables._request_table_options(context)[0] is expected


def test_context_disable_skips_model_resolution(monkeypatch):
    from types import SimpleNamespace

    context = SimpleNamespace(request=SimpleNamespace(options={"recognize_tables": False}, config_snapshot={}))
    monkeypatch.setattr(tables, "resolve_rapidtable_model", lambda *_: pytest.fail("disabled option loaded a model"))
    original = _outcome()
    outcome, result = tables.enrich_ocr_table_structure("missing.png", original, context=context)
    assert outcome is original
    assert result.status is tables.TableRecognitionStatus.DISABLED


def test_request_merges_and_cancellation_are_forwarded():
    from types import SimpleNamespace

    from docwen_core.cancellation import CancellationToken
    from docwen_core.errors import CancellationRequested

    token = CancellationToken()
    context = SimpleNamespace(
        request=SimpleNamespace(options={"table_merge_strategy": "empty"}, config_snapshot={}),
        cancellation=token.view(),
    )
    _, strategy, check = tables._request_table_options(context)
    assert strategy == "empty"
    assert check is not None
    token.cancel()
    with pytest.raises(CancellationRequested):
        check()


def test_model_build_pins_match_runtime():
    from scripts.release.packaged_resources import EXTERNAL_MODEL_SPECS, REQUIRED_MODEL_FILES

    for filename, url, digest in (
        (tables.RAPIDTABLE_MODEL_FILENAME, tables.RAPIDTABLE_MODEL_URL, tables.RAPIDTABLE_MODEL_SHA256),
        (tables.TABLE_LAYOUT_MODEL_FILENAME, tables.TABLE_LAYOUT_MODEL_URL, tables.TABLE_LAYOUT_MODEL_SHA256),
    ):
        key = "rapidtable/" + filename
        assert EXTERNAL_MODEL_SPECS[key] == (url, digest)
        assert key in REQUIRED_MODEL_FILES


def test_ruled_grid_crop_keeps_empty_cells_and_ignores_blank_margin():
    import cv2

    from docwen_core.text._table_layout import refine_ruled_table_bounds

    image = np.full((260, 500, 3), 255, np.uint8)
    for y in (80, 120, 160, 200):
        cv2.line(image, (20, y), (470, y), (0, 0, 0), 1)
    for x in (20, 170, 320, 470):
        cv2.line(image, (x, 80), (x, 200), (0, 0, 0), 1)
    assert refine_ruled_table_bounds(image, [(15, 25, 475, 202)]) == [(16, 76, 475, 205)]
    assert refine_ruled_table_bounds(np.full_like(image, 255), [(15, 25, 475, 202)]) == [(15, 25, 475, 202)]
