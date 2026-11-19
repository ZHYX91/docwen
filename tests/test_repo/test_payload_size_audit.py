"""Size reporting must count physical payload bytes, not resolved links."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from scripts.build.payload_size_audit import audit_payload_size

pytestmark = pytest.mark.unit


def test_payload_size_report_groups_runtime_files(tmp_path: Path) -> None:
    gui_lib = tmp_path / "_internal" / "PySide6" / "QtCore.so"
    ocr_lib = tmp_path / "_internal" / "cv2" / "cv2.so"
    model = tmp_path / "models" / "rapidocr" / "rec.onnx"
    for path, content in ((gui_lib, b"123456"), (ocr_lib, b"1234"), (model, b"12345")):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content)

    report = audit_payload_size(tmp_path)
    assert report["physical_bytes"] == 15
    assert report["regular_files"] == 3
    assert report["links"] == 0
    assert report["top_groups"] == [
        ("_internal/PySide6", 6),
        ("models", 5),
        ("_internal/cv2", 4),
    ]
    assert report["largest_files"] == [
        ("_internal/PySide6/QtCore.so", 6),
        ("models/rapidocr/rec.onnx", 5),
        ("_internal/cv2/cv2.so", 4),
    ]


@pytest.mark.skipif(os.name != "posix", reason="POSIX release symlinks only")
def test_payload_size_report_does_not_double_count_shared_models(tmp_path: Path) -> None:
    shared = tmp_path / "models" / "rapidocr" / "rec.onnx"
    shared.parent.mkdir(parents=True)
    shared.write_bytes(b"shared-model")
    embedded = tmp_path / "_internal" / "rapidocr_onnxruntime" / "models" / "rec.onnx"
    embedded.parent.mkdir(parents=True)
    embedded.symlink_to("../../../models/rapidocr/rec.onnx")

    report = audit_payload_size(tmp_path)
    assert report["physical_bytes"] == len(b"shared-model")
    assert report["regular_files"] == 1
    assert report["links"] == 1


def test_payload_size_report_requires_plain_directory(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="size_audit_requires_plain_directory"):
        audit_payload_size(tmp_path / "not-a-directory")
