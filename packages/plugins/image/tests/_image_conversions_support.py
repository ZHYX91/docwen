"""Golden-style semantic tests for the image plugin routes."""

from __future__ import annotations

import hashlib
import json
import os
import struct
import tempfile
from io import BytesIO
from pathlib import Path
from typing import Any

import pytest
from PIL import Image, PngImagePlugin

pytestmark = pytest.mark.golden

PROJECT_ROOT = Path(__file__).resolve().parents[4]


def _ocr_success(text: str) -> Any:
    from docwen_core.text.ocr import OcrOutcome, OcrStatus

    return OcrOutcome(OcrStatus.SUCCESS, text=text)


def _write_exif_jpeg(path: Path, input_spec: dict[str, Any]) -> None:
    exif = Image.Exif()
    for tag, value in input_spec["selected_exif_tags"].items():
        exif[int(tag)] = value
    Image.new("RGB", tuple(input_spec["size"]), (120, 80, 160)).save(path, format="JPEG", exif=exif)


def _assert_pdf_embedded_image_projection(doc: Any, expected: dict[str, Any]) -> None:
    images = doc[0].get_images(full=True)
    assert len(images) == 1
    extracted = doc.extract_image(images[0][0])
    assert extracted["ext"] == expected["ext"]

    with Image.open(BytesIO(extracted["image"])) as embedded:
        exif = embedded.getexif()
        assert embedded.format == expected["format"]
        assert list(embedded.size) == expected["size"]
        assert len(exif) == expected["exif_tag_count"]
        for tag, value in expected["selected_exif_tags"].items():
            assert exif.get(int(tag)) == value


def _deliverable_artifacts(result: Any) -> list[Any]:
    manifests = [
        artifact for artifact in result.artifacts if artifact.media_type == "application/vnd.docwen.document-node+json"
    ]
    assert len(manifests) == 1
    return [artifact for artifact in result.artifacts if artifact not in manifests]


def _document_node_root(path: Path, output_dir: Path) -> Path:
    relative = path.relative_to(output_dir)
    assert len(relative.parts) >= 2
    root = output_dir / relative.parts[0]
    assert root.is_dir()
    return root


def _build_fake_context(
    input_path: str,
    staging_dir: str,
    target_format: str,
    options: dict[str, Any] | None = None,
    action_name: str = "",
    extra_input_paths: list[str] | None = None,
    *,
    source_format: str | None = None,
    pre_cancelled: bool = False,
    config_values: dict[str, Any] | None = None,
    config_snapshot: dict[str, Any] | None = None,
) -> Any:
    from tests.support.config import FakeConfigView
    from tests.support.execution import FakeExecutionContext
    from tests.support.logging import FakePluginLogger
    from tests.support.progress import FakeProgressSink
    from tests.support.workspace import FakeWorkspaceHandle

    from docwen_core.cancellation import CancellationToken
    from docwen_core.models.file_ref import FileRef
    from docwen_core.models.request import ConversionRequest, OutputPolicy

    paths = [input_path, *(extra_input_paths or [])]
    file_refs = [
        FileRef(
            path=p,
            format=(source_format if index == 0 and source_format is not None else Path(p).suffix.lstrip(".")),
            category="image",
        )
        for index, p in enumerate(paths)
    ]
    request = ConversionRequest(
        request_id="test-image-001",
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
        cancellation=token,
        logger=FakePluginLogger(),
    )


__all__ = (
    "PROJECT_ROOT",
    "Any",
    "BytesIO",
    "Image",
    "Path",
    "PngImagePlugin",
    "_assert_pdf_embedded_image_projection",
    "_build_fake_context",
    "_deliverable_artifacts",
    "_document_node_root",
    "_ocr_success",
    "_write_exif_jpeg",
    "hashlib",
    "json",
    "os",
    "pytest",
    "pytestmark",
    "struct",
    "tempfile",
)
