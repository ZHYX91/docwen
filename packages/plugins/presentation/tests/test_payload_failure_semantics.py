"""Honest success semantics for recoverable PPTX payload loss."""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any, cast
from unittest.mock import Mock

import pytest
from lxml import etree

from docwen_core.export_semantics import MarkdownExportSemantics
from docwen_core.protocols.execution_context import ConverterContext
from docwen_plugin_presentation.pptx_md.converter import PptxToMarkdownConverter
from docwen_plugin_presentation.pptx_md.request_policy import (
    PresentationMarkdownRequestPolicy,
)

pytestmark = pytest.mark.unit


def _policy() -> PresentationMarkdownRequestPolicy:
    return PresentationMarkdownRequestPolicy(
        export=MarkdownExportSemantics(
            image_extraction_mode="file",
            ocr_placement_mode="main_md",
            image_link_style="markdown_embed",
            md_file_link_style="markdown_link",
        ),
        ocr_blockquote_title="",
    )


def _payload_stats() -> dict[str, Any]:
    return {
        "charts": 0,
        "audio": 0,
        "video": 0,
        "chart_snapshot_unavailable_locations": [],
        "notes_unavailable_locations": [],
        "payload_warning_diagnostics": [],
    }


def _slide_context(tmp_path: Path) -> ConverterContext:
    registered_artifacts: list[Any] = []

    def create_artifact_path(kind: str, suffix: str) -> str:
        return str(tmp_path / f"{kind}-{len(registered_artifacts)}{suffix}")

    return cast(
        ConverterContext,
        SimpleNamespace(
            cancellation=SimpleNamespace(check=lambda: None),
            progress=SimpleNamespace(report_progress=lambda *_args, **_kwargs: None),
            request=SimpleNamespace(options={}),
            logger=Mock(),
            workspace=SimpleNamespace(
                create_artifact_path=create_artifact_path,
                add_artifact=registered_artifacts.append,
                registered_artifacts=registered_artifacts,
            ),
        ),
    )


def _shape(**overrides: Any) -> Any:
    values = {
        "top": 0,
        "left": 0,
        "has_table": False,
        "has_chart": False,
        "shape_type": None,
        "has_text_frame": False,
        "element": None,
    }
    values.update(overrides)
    return SimpleNamespace(**values)


@pytest.mark.parametrize(
    ("payload_kind", "expected_code"),
    (
        ("table", "PPTX-TABLE-UNAVAILABLE"),
        ("chart", "PPTX-CHART-UNAVAILABLE"),
        ("media", "PPTX-MEDIA-UNAVAILABLE"),
    ),
)
def test_detected_payload_failure_is_recorded_without_aborting_slide(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    payload_kind: str,
    expected_code: str,
) -> None:
    converter = PptxToMarkdownConverter()
    shape = _shape()
    if payload_kind == "table":

        class BrokenTableShape:
            top = 0
            left = 0
            has_table = True

            @property
            def table(self) -> object:
                raise RuntimeError("table XML is corrupt")

        shape = BrokenTableShape()
    elif payload_kind == "chart":
        shape = _shape(has_chart=True)
        monkeypatch.setattr(
            converter,
            "_preserve_chart_payload",
            Mock(side_effect=RuntimeError("chart relationship is corrupt")),
        )
    else:
        shape = _shape(element=object())
        monkeypatch.setattr(converter, "_shape_media_kind", lambda _shape: "audio")
        monkeypatch.setattr(
            converter,
            "_preserve_media_payload",
            Mock(side_effect=RuntimeError("audio relationship is corrupt")),
        )

    stats = _payload_stats()
    lines, *_ = converter._process_slide(
        SimpleNamespace(shapes=[shape], part=SimpleNamespace(rels={})),
        2,
        _slide_context(tmp_path),
        request_policy=_policy(),
        payload_stats=stats,
    )

    assert lines[:2] == ["## Slide 2", ""]
    assert [warning["code"] for warning in stats["payload_warning_diagnostics"]] == [expected_code]


def test_detected_image_write_failure_is_recorded(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    converter = PptxToMarkdownConverter()
    stats = _payload_stats()
    context = _slide_context(tmp_path)
    shape = _shape(
        shape_type=13,
        image=SimpleNamespace(ext="png", blob=b"not-a-real-png"),
    )

    def fail_open(*_args: Any, **_kwargs: Any) -> Any:
        raise OSError("staging is unavailable")

    monkeypatch.setattr("builtins.open", fail_open)
    lines, image_count, *_ = converter._process_slide(
        SimpleNamespace(shapes=[shape], part=SimpleNamespace(rels={})),
        4,
        context,
        request_policy=_policy(),
        payload_stats=stats,
    )

    assert lines[:2] == ["## Slide 4", ""]
    assert image_count == 1
    assert stats["payload_warning_diagnostics"] == [
        {
            "code": "PPTX-IMAGE-UNAVAILABLE",
            "message": "An image was detected but could not be exported.",
            "location": "slide 4: image 1",
        }
    ]


def test_corrupt_smartart_relationship_is_recorded(tmp_path: Path) -> None:
    shape_xml = etree.fromstring(
        b"""
        <p:graphicFrame
          xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
              <dgm:relIds r:dm="rId7"/>
            </a:graphicData>
          </a:graphic>
        </p:graphicFrame>
        """
    )
    stats = _payload_stats()

    PptxToMarkdownConverter()._process_slide(
        SimpleNamespace(
            shapes=[_shape(element=shape_xml)],
            part=SimpleNamespace(rels={}),
        ),
        5,
        _slide_context(tmp_path),
        request_policy=_policy(),
        payload_stats=stats,
    )

    assert stats["payload_warning_diagnostics"] == [
        {
            "code": "PPTX-SMARTART-UNAVAILABLE",
            "message": "SmartArt was detected but one diagram payload could not be read.",
            "location": "slide 5: SmartArt relationship rId7",
        }
    ]


def test_chart_workbook_failure_keeps_semantics_and_records_warning(tmp_path: Path) -> None:
    converter = PptxToMarkdownConverter()
    stats = _payload_stats()
    context = _slide_context(tmp_path)
    chart = SimpleNamespace(
        plots=[SimpleNamespace(categories=[SimpleNamespace(label="Q1")])],
        series=[SimpleNamespace(name="Revenue", values=[12])],
        has_title=False,
        part=SimpleNamespace(
            rels={
                "rIdWorkbook": SimpleNamespace(
                    reltype="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
                    target_part=SimpleNamespace(blob=b"workbook"),
                )
            }
        ),
    )
    converter._register_payload_artifact = Mock(side_effect=OSError("cannot write workbook"))  # type: ignore[method-assign]

    lines = converter._preserve_chart_payload(
        SimpleNamespace(chart=chart),
        1,
        context,
        resource_stem="deck",
        chart_index=1,
        payload_stats=stats,
    )

    assert any("Revenue" in line for line in lines)
    assert [warning["code"] for warning in stats["payload_warning_diagnostics"]] == ["PPTX-CHART-WORKBOOK-UNAVAILABLE"]


def test_media_poster_failure_keeps_media_and_records_warning(tmp_path: Path) -> None:
    converter = PptxToMarkdownConverter()
    stats = _payload_stats()
    context = _slide_context(tmp_path)

    class MediaElement:
        def xpath(self, expression: str) -> list[str]:
            return ["rIdMedia"] if "audioFile" in expression else ["rIdPoster"]

    slide = SimpleNamespace(
        part=SimpleNamespace(
            rels={
                "rIdMedia": SimpleNamespace(target_part=SimpleNamespace(blob=b"audio", partname="/ppt/media/audio.mp3"))
            }
        )
    )
    lines = converter._preserve_media_payload(
        SimpleNamespace(element=MediaElement()),
        slide,
        3,
        context,
        resource_stem="deck",
        media_kind="audio",
        media_index=1,
        payload_stats=stats,
    )

    assert lines[0].startswith("[Audio payload]")
    assert [warning["code"] for warning in stats["payload_warning_diagnostics"]] == ["PPTX-MEDIA-POSTER-UNAVAILABLE"]


def test_convert_projects_payload_warnings_to_result_diagnostics(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    converter = PptxToMarkdownConverter()
    source = tmp_path / "source.pptx"
    source.write_bytes(b"pptx")
    registered_artifacts: list[Any] = []
    workspace = SimpleNamespace(
        input_path=str(source),
        registered_artifacts=registered_artifacts,
        create_artifact_path=lambda _kind, suffix: str(tmp_path / f"output{suffix}"),
        add_artifact=registered_artifacts.append,
    )
    stats = {
        "slides": 1,
        "hidden_slides": 0,
        "tables": 0,
        "images": 0,
        "smartart_texts": 0,
        "charts": 0,
        "audio": 0,
        "video": 0,
        "title": "source",
        "chart_snapshot_unavailable_locations": [],
        "notes_unavailable_locations": [],
        "payload_warning_diagnostics": [
            {
                "code": "PPTX-TABLE-UNAVAILABLE",
                "message": "A table was detected but could not be converted to Markdown.",
                "location": "slide 1: table 1",
            }
        ],
    }
    monkeypatch.setattr(converter, "_parse_pptx", lambda *_args, **_kwargs: ("# source\n", stats))
    context = cast(
        ConverterContext,
        SimpleNamespace(
            request=SimpleNamespace(request_id="payload-warning", options={}),
            workspace=workspace,
            cancellation=SimpleNamespace(check=lambda: None),
            progress=SimpleNamespace(
                report_progress=lambda *_args, **_kwargs: None,
                report_artifact_ready=lambda *_args, **_kwargs: None,
            ),
            logger=Mock(),
        ),
    )

    result = converter.convert(context)

    assert result.success is True
    assert result.artifacts[0].metadata["payload_warning_count"] == 1
    warning = next(diagnostic for diagnostic in result.diagnostics if diagnostic.code == "PPTX-TABLE-UNAVAILABLE")
    assert warning.level == "warning"
    assert warning.location == "slide 1: table 1"
