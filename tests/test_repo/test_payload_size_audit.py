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
    assert report["ocr_model_bytes"] == 5
    assert report["other_model_bytes"] == 0
    assert report["non_ocr_bytes"] == 10
    assert report["non_model_bytes"] == 10
    assert report["top_non_ocr_groups"] == [("_internal/PySide6", 6), ("_internal/cv2", 4)]
    assert report["largest_non_ocr_files"] == [
        ("_internal/PySide6/QtCore.so", 6),
        ("_internal/cv2/cv2.so", 4),
    ]
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
    assert report["ocr_model_bytes"] == len(b"shared-model")
    assert report["non_ocr_bytes"] == 0
    assert report["regular_files"] == 1
    assert report["links"] == 1


@pytest.mark.skipif(os.name != "posix", reason="POSIX frozen library links")
def test_packaging_preserves_native_aliases_and_gui_only_files(tmp_path: Path) -> None:
    from scripts.build.build import copy_frozen_runtime

    cli = tmp_path / "cli" / "_internal"
    library = cli / "numpy.libs" / "blas.so"
    library.parent.mkdir(parents=True)
    library.write_bytes(b"native")
    (cli / "blas.so").symlink_to("numpy.libs/blas.so")
    (cli / "entry").write_bytes(b"cli")
    gui = tmp_path / "combined" / "_internal"
    gui.mkdir(parents=True)
    (gui / "keep").write_bytes(b"gui")
    (gui / "entry").symlink_to("keep")
    (gui / "blas.so").write_bytes(b"old library")

    copy_frozen_runtime(cli, gui)
    copy_frozen_runtime(cli, gui)

    assert (gui / "blas.so").readlink() == Path("numpy.libs/blas.so")
    assert (gui / "blas.so").read_bytes() == b"native"
    assert (gui / "entry").read_bytes() == b"cli"
    assert (gui / "keep").read_bytes() == b"gui"
    assert audit_payload_size(gui)["physical_bytes"] == len(b"nativecligui")


@pytest.mark.skipif(os.name != "posix", reason="POSIX frozen library links")
def test_packaging_rejects_external_library_link(tmp_path: Path) -> None:
    from scripts.build.build import copy_frozen_runtime

    cli = tmp_path / "cli"
    cli.mkdir()
    (tmp_path / "outside").write_bytes(b"external")
    (cli / "alias").symlink_to("../outside")

    with pytest.raises(ValueError, match="invalid frozen runtime link"):
        copy_frozen_runtime(cli, tmp_path / "combined")


def test_payload_size_report_distinguishes_other_onnx_models(tmp_path: Path) -> None:
    data = {
        "models/rapidocr/rec.onnx": b"a" * 2,
        "_internal/rapidocr_onnxruntime/models/rec.onnx": b"a" * 2,
        "models/rapidtable/layout_table.onnx": b"b" * 3,
        "_internal/pymupdf/layout/resources/onnx/layout_rf2.4.1.onnx": b"c" * 4,
        "_internal/pymupdf/layout/resources/onnx/layout_rf2.4.1.yaml": b"d" * 5,
        "_internal/cv2/cv2.so": b"e" * 6,
        "_internal/onnxruntime/datasets/mul_1.onnx": b"f" * 7,
    }
    for name, payload in data.items():
        path = tmp_path / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(payload)
    report = audit_payload_size(tmp_path)
    assert report["physical_bytes"] == 29
    assert report["ocr_model_bytes"] == 4
    assert report["other_model_bytes"] == 14
    assert report["non_ocr_bytes"] == 25
    assert report["non_model_bytes"] == 11


def test_payload_size_report_requires_plain_directory(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="size_audit_requires_plain_directory"):
        audit_payload_size(tmp_path / "not-a-directory")
