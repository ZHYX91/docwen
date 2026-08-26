"""Golden-style semantic tests for the Layout plugin routes.

Covered routes:
- pdf → png / jpg / tif     (LayoutToImageConverter — golden)
- pdf → docx / doc / odt / rtf (LayoutToDocumentConverter — mock)
- ofd/xps → md / png / docx    (preprocess chain — mock)
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import types
from pathlib import Path
from typing import TYPE_CHECKING, Any

import pytest
from tests.support.xps import (
    create_image_xps,
    create_minimal_xps,
    docx_semantic_projection,
    pdf_visual_projection,
    png_visual_projection,
    raster_visual_projection,
)

if TYPE_CHECKING:
    import fitz

pytestmark = pytest.mark.contract

_PROJECT_ROOT = Path(__file__).resolve().parents[4]

_PDF_OPS_OLD_SYSTEM_FIXTURE = (
    _PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_pdf_operations_semantics.json"
)

_LAYOUT_PDF_OLD_SYSTEM_FIXTURE = (
    _PROJECT_ROOT / "tests" / "fixtures" / "golden" / "old_system_layout_pdf_semantics.json"
)


def _ocr_success(text: str) -> Any:
    from docwen_core.text.ocr import OcrOutcome, OcrStatus

    return OcrOutcome(OcrStatus.SUCCESS, text=text)


def _page_fragments(result: Any) -> list[Any]:
    return [
        artifact
        for artifact in result.artifacts
        if artifact.kind == "auxiliary" and artifact.metadata.get("fragment_kind") == "page"
    ]


def _document_node_root(path: Path, output_dir: Path) -> Path:
    relative = path.relative_to(output_dir)
    assert len(relative.parts) >= 2
    root = output_dir / relative.parts[0]
    assert root.is_dir()
    return root


def _io_path(path: str | Path) -> Path:
    from docwen_runtime.path_io import filesystem_path

    return filesystem_path(str(path))


def _write_valid_png(path: Path) -> None:
    """Write a decodable PNG for content-first extracted-resource tests."""
    from PIL import Image

    path.parent.mkdir(parents=True, exist_ok=True)
    with Image.new("RGB", (2, 2), (1, 2, 3)) as image:
        image.save(path, format="PNG")


def _build_fake_context(
    input_path: str,
    staging_dir: str,
    target_format: str,
    options: dict[str, Any] | None = None,
    action_name: str = "",
    source_format: str = "",
    *,
    config_values: dict[str, Any] | None = None,
    config_snapshot: dict[str, Any] | None = None,
    pre_cancelled: bool = False,
) -> Any:
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    fmt = source_format or Path(input_path).suffix.lstrip(".")
    file_refs = [
        FileRef(
            path=input_path,
            format=fmt,
            category="layout",
        )
    ]
    request = ConversionRequest(
        request_id="test-layout-001",
        input_refs=file_refs,
        target_format=target_format,
        action_name=action_name,
        options=options or {},
        config_snapshot=config_snapshot or {},
        output_policy=OutputPolicy(),
    )
    config = FakeConfigView(config_values)
    token = CancellationToken()
    if pre_cancelled:
        token.cancel("test cancellation")
    return FakeExecutionContext(
        request=request,
        workspace=FakeWorkspaceHandle(input_path, staging_dir),
        config=config,
        progress=FakeProgressSink(),
        cancellation=token.view(),
        logger=FakePluginLogger(),
    )


def _build_runtime_pipeline() -> tuple[Any, Any, str]:
    """Build the real runtime pipeline with the layout plugin."""
    from docwen_plugin_layout import LayoutPlugin
    from docwen_runtime.engine.route_resolver import RouteResolver
    from docwen_runtime.engine.task_manager import TaskManager
    from docwen_runtime.output.finalizer import OutputFinalizer
    from docwen_runtime.plugin_registry.registry import PluginRegistry
    from docwen_runtime.workspace.manager import WorkspaceManager

    registry = PluginRegistry()
    registry.register(LayoutPlugin())
    resolver = RouteResolver(registry)
    ws_root = tempfile.mkdtemp(prefix="docwen_layout_runtime_")
    ws_manager = WorkspaceManager(ws_root)
    task_manager = TaskManager(registry, resolver, ws_manager, OutputFinalizer())
    return task_manager, ws_manager, ws_root


def _load_pdf_ops_old_system_fixture() -> dict[str, Any]:
    return json.loads(_PDF_OPS_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _load_layout_pdf_old_system_fixture() -> dict[str, Any]:
    return json.loads(_LAYOUT_PDF_OLD_SYSTEM_FIXTURE.read_text(encoding="utf-8"))


def _create_text_pdf(path: Path, page_texts: list[str]) -> None:
    import fitz

    doc = fitz.open()
    try:
        for text in page_texts:
            page = doc.new_page(width=240, height=120)
            page.insert_text((36, 60), text, fontsize=12)
        doc.save(str(path))
    finally:
        doc.close()


def _assert_open_pdf_document(document: Any, expected_path: Path) -> None:
    """Assert a pymupdf4llm test double received an explicit live PDF document."""
    assert bool(document.is_pdf) is True
    assert bool(document.is_closed) is False
    assert Path(str(document.name)) == expected_path


def _pdf_page_text(page: fitz.Page) -> str:
    text = page.get_text("text")
    assert isinstance(text, str)
    return text.strip()


def _pdf_page_texts(path: str | Path) -> list[str]:
    import fitz

    with fitz.open(str(path)) as doc:
        return [_pdf_page_text(page) for page in doc]


_PDF_METADATA_FIELDS = ("title", "author", "subject", "keywords", "creator", "producer")


def _create_metadata_pdf(path: Path, input_spec: dict[str, Any]) -> None:
    import fitz

    width, height = input_spec["page_size"]
    doc = fitz.open()
    try:
        for text in input_spec["page_texts"]:
            page = doc.new_page(width=width, height=height)
            page.insert_text((36, 60), text, fontsize=12)
        doc.set_metadata(input_spec["metadata"])
        doc.save(str(path))
    finally:
        doc.close()


def _create_rotated_pdf(path: Path, input_spec: dict[str, Any]) -> None:
    import fitz

    width, height = input_spec["page_size"]
    doc = fitz.open()
    try:
        for text in input_spec["page_texts"]:
            page = doc.new_page(width=width, height=height)
            page.insert_text((8, 18), text, fontsize=6)
            page.set_rotation(input_spec["rotation"])
        doc.save(str(path))
    finally:
        doc.close()


def _pdf_metadata_projection(path: str | Path) -> dict[str, Any]:
    import fitz

    pdf_path = Path(path)
    with fitz.open(str(pdf_path)) as doc:
        metadata = doc.metadata or {}
        return {
            "pdf_magic": pdf_path.read_bytes()[:5].decode("ascii"),
            "page_count": doc.page_count,
            "page_texts": [_pdf_page_text(page) for page in doc],
            "metadata": {field: metadata.get(field, "") for field in _PDF_METADATA_FIELDS},
        }


def _pdf_geometry_projection(path: str | Path) -> list[dict[str, Any]]:
    import fitz

    with fitz.open(str(path)) as doc:
        return [
            {
                "text": _pdf_page_text(page),
                "rotation": page.rotation,
                "rect": [
                    round(page.rect.x0, 2),
                    round(page.rect.y0, 2),
                    round(page.rect.x1, 2),
                    round(page.rect.y1, 2),
                ],
                "mediabox": [
                    round(page.mediabox.x0, 2),
                    round(page.mediabox.y0, 2),
                    round(page.mediabox.x1, 2),
                    round(page.mediabox.y1, 2),
                ],
            }
            for page in doc
        ]


def _create_interactive_geometry_pdf(
    path: Path,
    input_spec: dict[str, Any],
    page_profiles: dict[str, dict[str, Any]],
) -> None:
    import fitz

    doc = fitz.open()
    try:
        for page_spec in input_spec["pages"]:
            profile = page_profiles[page_spec["profile"]]
            width, height = profile["page_size"]
            page = doc.new_page(width=width, height=height)
            page.set_cropbox(fitz.Rect(profile["cropbox"]))
            page.insert_text((30, 50), page_spec["text"], fontsize=12)
            page.insert_link(
                {
                    "kind": fitz.LINK_URI,
                    "from": fitz.Rect(25, 60, 160, 82),
                    "uri": page_spec["uri"],
                }
            )
            annotation = page.add_text_annot((180, 90), page_spec["note"])
            annotation.set_info(
                title=f"{page_spec['text']}-AUTHOR",
                content=page_spec["note"],
            )
            annotation.update()
            page.set_rotation(profile["rotation"])
        doc.save(str(path))
    finally:
        doc.close()


def _pdf_interactive_geometry_projection(path: str | Path) -> list[dict[str, Any]]:
    import fitz

    def rounded_rect(rect: fitz.Rect) -> list[float]:
        return [round(value, 3) for value in (rect.x0, rect.y0, rect.x1, rect.y1)]

    projection: list[dict[str, Any]] = []
    with fitz.open(str(path)) as doc:
        for page in doc:
            annotations = []
            annotation_iter = page.annots()
            if annotation_iter is not None:
                for annotation in annotation_iter:
                    annotations.append(
                        {
                            "type": annotation.type[1],
                            "content": annotation.info.get("content") or "",
                            "title": annotation.info.get("title") or "",
                            "rect": rounded_rect(annotation.rect),
                        }
                    )
            projection.append(
                {
                    "text": _pdf_page_text(page),
                    "rotation": page.rotation,
                    "rect": rounded_rect(page.rect),
                    "mediabox": rounded_rect(page.mediabox),
                    "cropbox": rounded_rect(page.cropbox),
                    "links": [
                        {
                            "kind": link.get("kind"),
                            "uri": link.get("uri") or "",
                            "from": rounded_rect(link["from"]),
                        }
                        for link in page.get_links()
                    ],
                    "annotations": annotations,
                }
            )
    return projection


def _expected_interactive_geometry_projection(
    page_specs: list[dict[str, Any]],
    page_profiles: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    expected = []
    for page_spec in page_specs:
        profile = page_profiles[page_spec["profile"]]
        output = profile["expected_insert_pdf_geometry"]
        expected.append(
            {
                "text": page_spec["text"],
                "rotation": profile["rotation"],
                "rect": output["rect"],
                "mediabox": output["mediabox"],
                "cropbox": output["cropbox"],
                "links": [
                    {
                        "kind": 2,
                        "uri": page_spec["uri"],
                        "from": output["link_from"],
                    }
                ],
                "annotations": [
                    {
                        "type": "Text",
                        "content": page_spec["note"],
                        "title": f"{page_spec['text']}-AUTHOR",
                        "rect": output["annotation_rect"],
                    }
                ],
            }
        )
    return expected


def _create_forms_actions_pdf(path: Path, input_spec: dict[str, Any]) -> None:
    import fitz

    class _TextWidget(fitz.Widget):
        field_name: str | None
        field_label: str | None
        field_value: str | None
        rect: fitz.Rect | None

    widget_type = getattr(fitz, "PDF_WIDGET_TYPE_TEXT", None)
    assert isinstance(widget_type, int)

    doc = fitz.open()
    try:
        for page_id in input_spec["page_ids"]:
            page = doc.new_page(width=420, height=300)
            page.insert_text((36, 48), f"FORM-{page_id}", fontsize=12)

            widget = _TextWidget()
            widget.field_name = f"field_{page_id}"
            widget.field_label = f"LABEL-{page_id}"
            widget.field_type = widget_type
            widget.field_value = f"VALUE-{page_id}"
            widget.rect = fitz.Rect(36, 80, 220, 110)
            page.add_widget(widget)

            annotation = page.add_file_annot(
                (280, 90),
                f"ATTACHMENT-{page_id}\n".encode(),
                f"attachment-{page_id}.txt",
                ufilename=f"attachment-{page_id}-unicode.txt",
                desc=f"DESC-{page_id}",
            )
            annotation.set_info(
                title=f"AUTHOR-{page_id}",
                content=f"CONTENT-{page_id}",
            )
            annotation.update()

        for link in input_spec["internal_goto_links"]:
            doc[link["source_page"]].insert_link(
                {
                    "kind": fitz.LINK_GOTO,
                    "from": fitz.Rect(36, 130, 220, 155),
                    "page": link["target_page"],
                    "to": fitz.Point(36, 48),
                }
            )
        doc.save(str(path))
    finally:
        doc.close()


def _pdf_forms_actions_projection(path: str | Path) -> list[dict[str, Any]]:
    import hashlib

    import fitz

    def rounded_rect(rect: fitz.Rect) -> list[float]:
        return [round(value, 3) for value in (rect.x0, rect.y0, rect.x1, rect.y1)]

    projection = []
    with fitz.open(str(path)) as doc:
        for page in doc:
            links = []
            for link in page.get_links():
                target_index = link.get("page", -1)
                target_text = ""
                if isinstance(target_index, int) and 0 <= target_index < doc.page_count:
                    target_text = _pdf_page_text(doc[target_index])
                links.append(
                    {
                        "kind": link.get("kind"),
                        "target_page_index": target_index,
                        "target_text": target_text,
                    }
                )

            annotations = []
            annotation_iter = page.annots()
            if annotation_iter is not None:
                for annotation in annotation_iter:
                    payload = annotation.get_file()
                    annotations.append(
                        {
                            "type": annotation.type[1],
                            "content": annotation.info.get("content") or "",
                            "title": annotation.info.get("title") or "",
                            "rect": rounded_rect(annotation.rect),
                            "filename": annotation.file_info.get("filename") or "",
                            "payload_size": len(payload),
                            "payload_sha256": hashlib.sha256(payload).hexdigest(),
                        }
                    )

            widgets = []
            widget_iter = page.widgets()
            if widget_iter is not None:
                for widget in widget_iter:
                    assert isinstance(widget, fitz.Widget)
                    assert isinstance(widget.rect, fitz.Rect)
                    widgets.append(
                        {
                            "field_name": widget.field_name or "",
                            "field_label": widget.field_label or "",
                            "field_type": widget.field_type,
                            "field_type_string": widget.field_type_string or "",
                            "field_value": str(widget.field_value or ""),
                            "rect": rounded_rect(widget.rect),
                        }
                    )

            projection.append(
                {
                    "text": _pdf_page_text(page),
                    "links": links,
                    "annotations": annotations,
                    "widgets": widgets,
                }
            )
    return projection


def _assert_yaml_title(content: str, title_key: str, title_value: str) -> None:
    assert content.startswith("---\n"), content[:200]
    yaml_block = content.split("---", 2)[1]
    assert f"{title_key}: {title_value}" in yaml_block
    assert "aliases:" in yaml_block


def _create_pdf_ops_fixture_inputs(tmp_path: Path) -> dict[str, Path]:
    fixture = _load_pdf_ops_old_system_fixture()
    paths: dict[str, Path] = {}

    for item in fixture["input_pdfs"]["merge"]:
        path = tmp_path / item["name"]
        _create_text_pdf(path, item["page_texts"])
        paths[item["name"]] = path

    for key in ("split", "single_page"):
        item = fixture["input_pdfs"][key]
        path = tmp_path / item["name"]
        _create_text_pdf(path, item["page_texts"])
        paths[item["name"]] = path

    return paths


__all__ = (
    "Any",
    "Path",
    "_assert_open_pdf_document",
    "_assert_yaml_title",
    "_build_fake_context",
    "_build_runtime_pipeline",
    "_create_forms_actions_pdf",
    "_create_interactive_geometry_pdf",
    "_create_metadata_pdf",
    "_create_pdf_ops_fixture_inputs",
    "_create_rotated_pdf",
    "_create_text_pdf",
    "_document_node_root",
    "_expected_interactive_geometry_projection",
    "_io_path",
    "_load_layout_pdf_old_system_fixture",
    "_load_pdf_ops_old_system_fixture",
    "_ocr_success",
    "_page_fragments",
    "_pdf_forms_actions_projection",
    "_pdf_geometry_projection",
    "_pdf_interactive_geometry_projection",
    "_pdf_metadata_projection",
    "_pdf_page_texts",
    "_write_valid_png",
    "create_image_xps",
    "create_minimal_xps",
    "docx_semantic_projection",
    "os",
    "pdf_visual_projection",
    "png_visual_projection",
    "pytest",
    "pytestmark",
    "raster_visual_projection",
    "sys",
    "tempfile",
    "types",
)
