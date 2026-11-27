"""Keep one verified on-disk copy of each RapidOCR default model."""

import base64
import csv
import hashlib
from pathlib import Path

import pytest
import yaml
from scripts.build.build import _share_rapidocr_default_models
from scripts.release.packaged_resources import (
    RAPIDOCR_DEFAULT_MODELS,
    RAPIDOCR_PACKAGE,
    verify_rapidocr_model_layout,
)

pytestmark = pytest.mark.unit


@pytest.fixture
def rapidocr_payload(tmp_path: Path) -> Path:
    vendor_dir = tmp_path / RAPIDOCR_PACKAGE / "models"
    vendor_dir.mkdir(parents=True)
    shared_dir = tmp_path / "models" / "rapidocr"
    shared_dir.mkdir(parents=True)
    config = {"Global": {"text_score": 0.5}}
    for section, filename in RAPIDOCR_DEFAULT_MODELS.items():
        (vendor_dir / filename).write_bytes(filename.encode())
        (shared_dir / filename).write_bytes(filename.encode())
        config[section] = {"model_path": f"models/{filename}"}
    (vendor_dir.parent / "config.yaml").write_text(yaml.safe_dump(config), encoding="utf-8")
    return tmp_path


def test_vendor_defaults_resolve_to_the_shared_files(rapidocr_payload: Path) -> None:
    package = rapidocr_payload / RAPIDOCR_PACKAGE
    record = package.parent / "rapidocr_onnxruntime-1.4.4.dist-info" / "RECORD"
    record.parent.mkdir()
    record.write_text(
        "rapidocr_onnxruntime/config.yaml,sha256=original,12\n"
        "rapidocr_onnxruntime/main.py,sha256=unchanged,9\n"
        + "".join(
            f"rapidocr_onnxruntime/models/{name},sha256=removed,12\n" for name in RAPIDOCR_DEFAULT_MODELS.values()
        ),
        encoding="utf-8",
    )
    _share_rapidocr_default_models(rapidocr_payload)
    verify_rapidocr_model_layout(rapidocr_payload, error_prefix="test")
    config_bytes = (package / "config.yaml").read_bytes()
    digest = base64.urlsafe_b64encode(hashlib.sha256(config_bytes).digest()).decode().rstrip("=")
    with record.open(encoding="utf-8", newline="") as stream:
        assert list(csv.reader(stream)) == [
            ["rapidocr_onnxruntime/config.yaml", f"sha256={digest}", str(len(config_bytes))],
            ["rapidocr_onnxruntime/main.py", "sha256=unchanged", "9"],
        ]
    config = yaml.safe_load(config_bytes)
    assert config["Global"]["text_score"] == 0.5
    for section, filename in RAPIDOCR_DEFAULT_MODELS.items():
        resolved = (package / config[section]["model_path"]).resolve()
        assert resolved == rapidocr_payload / "models" / "rapidocr" / filename
        assert resolved.read_bytes() == filename.encode()
        assert not (package / "models" / filename).exists()

    duplicate = package / "models" / RAPIDOCR_DEFAULT_MODELS["Det"]
    duplicate.write_bytes(b"unexpected copy")
    with pytest.raises(RuntimeError, match="rapidocr_duplicate_models"):
        verify_rapidocr_model_layout(rapidocr_payload, error_prefix="test")


def test_different_vendor_bytes_abort_before_any_changes(rapidocr_payload: Path) -> None:
    package = rapidocr_payload / RAPIDOCR_PACKAGE
    model = package / "models" / RAPIDOCR_DEFAULT_MODELS["Cls"]
    model.write_bytes(b"upgraded model")
    original_config = (package / "config.yaml").read_bytes()

    with pytest.raises(RuntimeError, match="rapidocr_default_model_bytes_differ"):
        _share_rapidocr_default_models(rapidocr_payload)

    assert (package / "config.yaml").read_bytes() == original_config
    assert {model.name for model in (package / "models").iterdir()} == set(RAPIDOCR_DEFAULT_MODELS.values())
