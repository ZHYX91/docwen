"""Regression tests for safe, optional POSIX RapidOCR model sharing."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from scripts.build.model_payload_dedup import compact_rapidocr_model_copies

pytestmark = pytest.mark.unit

MODEL_NAMES = (
    "ch_PP-OCRv4_det_infer.onnx",
    "ch_PP-OCRv4_rec_infer.onnx",
    "ch_ppocr_mobile_v2.0_cls_infer.onnx",
)


def _model_dirs(root: Path) -> tuple[Path, Path]:
    vendor = root / "_internal" / "rapidocr_onnxruntime" / "models"
    public = root / "models" / "rapidocr"
    vendor.mkdir(parents=True)
    public.mkdir(parents=True)
    return vendor, public


@pytest.mark.skipif(os.name != "posix", reason="POSIX release symlinks only")
def test_identical_models_keep_both_import_paths_with_one_physical_copy(tmp_path: Path) -> None:
    vendor, public = _model_dirs(tmp_path)
    for index, name in enumerate(MODEL_NAMES):
        content = (f"same-model-{index}-".encode() * (index + 2))
        (vendor / name).write_bytes(content)
        (public / name).write_bytes(content)

    old_bytes = sum((vendor / name).stat().st_size for name in MODEL_NAMES)
    result = compact_rapidocr_model_copies(tmp_path)
    assert result == {"linked": 3, "skipped": 0, "saved_bytes": old_bytes}
    for name in MODEL_NAMES:
        bundled = vendor / name
        shared = public / name
        assert bundled.is_symlink()
        assert bundled.read_bytes() == shared.read_bytes()
        assert bundled.resolve(strict=True) == shared.resolve(strict=True)
        assert not shared.is_symlink()
        assert not os.path.isabs(os.readlink(bundled))
        assert os.readlink(bundled).replace("\\", "/") == f"../../../models/rapidocr/{name}"

    # A repeated packaging pass must not rewrite or double-count existing links.
    assert compact_rapidocr_model_copies(tmp_path) == {"linked": 0, "skipped": 0, "saved_bytes": 0}


@pytest.mark.skipif(os.name != "posix", reason="POSIX release symlinks only")
def test_mismatched_vendor_models_are_kept_untouched(tmp_path: Path) -> None:
    vendor, public = _model_dirs(tmp_path)
    same_size_name, other_size_name, _ = MODEL_NAMES
    (vendor / same_size_name).write_bytes(b"ABC")
    (public / same_size_name).write_bytes(b"XYZ")
    (vendor / other_size_name).write_bytes(b"AB")
    (public / other_size_name).write_bytes(b"ABC")

    result = compact_rapidocr_model_copies(tmp_path)
    assert result == {"linked": 0, "skipped": 3, "saved_bytes": 0}
    assert (vendor / same_size_name).read_bytes() == b"ABC"
    assert (vendor / other_size_name).read_bytes() == b"AB"
    assert not (vendor / same_size_name).is_symlink()
    assert not (vendor / other_size_name).is_symlink()


@pytest.mark.skipif(os.name != "posix", reason="POSIX release symlinks only")
def test_only_verified_models_are_shared(tmp_path: Path) -> None:
    vendor, public = _model_dirs(tmp_path)
    first, second, third = MODEL_NAMES
    (vendor / first).write_bytes(b"matching")
    (public / first).write_bytes(b"matching")
    (vendor / second).write_bytes(b"vendor-only")
    (public / third).write_bytes(b"public-only")

    assert compact_rapidocr_model_copies(tmp_path) == {"linked": 1, "skipped": 2, "saved_bytes": 8}
    assert (vendor / first).is_symlink()
    assert (vendor / second).read_bytes() == b"vendor-only"
    assert (public / third).read_bytes() == b"public-only"


def test_windows_packages_are_never_rewritten(tmp_path: Path) -> None:
    vendor, public = _model_dirs(tmp_path)
    for name in MODEL_NAMES:
        (vendor / name).write_bytes(b"matching")
        (public / name).write_bytes(b"matching")

    assert compact_rapidocr_model_copies(tmp_path, platform="win32") == {
        "linked": 0,
        "skipped": 3,
        "saved_bytes": 0,
    }
    assert all(not (vendor / name).is_symlink() for name in MODEL_NAMES)
