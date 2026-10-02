"""Permanent oracle tests for the packaged physical-page Bundle verifier."""

from __future__ import annotations

import copy
import hashlib
from pathlib import Path
from typing import Any
from urllib.parse import quote

import pytest
from scripts.release.verify_packaged_cli import _OFD_FIXTURE_SCRIPT, _verify_physical_page_bundle

pytestmark = pytest.mark.contract


def test_ofd_fixture_disables_dependency_logging_before_import() -> None:
    assert _OFD_FIXTURE_SCRIPT.index('loguru_logger.disable("easyofd")') < _OFD_FIXTURE_SCRIPT.index(
        "from easyofd import OFD"
    )


def _artifact(root: Path, artifact_id: str, kind: str, payload: bytes) -> dict[str, Any]:
    path = root / f"{artifact_id}.bin"
    path.write_bytes(payload)
    return {
        "artifact_id": artifact_id,
        "kind": kind,
        "media_type": "text/markdown" if kind != "resource" else "image/png",
        "locator": path.name,
        "logical_path": path.name,
        "suggested_name": path.name,
        "size_bytes": len(payload),
        "sha256": hashlib.sha256(payload).hexdigest(),
    }


def _canonical_terminal(root: Path) -> dict[str, Any]:
    artifacts = [_artifact(root, "document.main", "document", b"primary only\n")]
    statuses = ("success", "no_text", "recognition_failed", "success")
    fragment_payloads = (b"page one\n", b"", b"", b"page four\n")
    artifacts.extend(
        _artifact(root, f"fragment.{page}", "fragment", fragment_payloads[page - 1]) for page in range(1, 5)
    )
    artifacts.extend(_artifact(root, f"resource.{page}", "resource", bytes([page])) for page in range(1, 6))
    relations = [
        {
            "type": "fragment_of",
            "source_artifact_id": f"fragment.{page}",
            "target_artifact_id": "document.main",
            "role": "ocr_page",
            "ordinal": page - 1,
            "page_fragment": {
                "fragment_kind": "page",
                "page_index": page,
                "page_count": 4,
                "ocr_status": statuses[page - 1],
                "source_page": page,
            },
        }
        for page in range(1, 5)
    ]
    relations.extend(
        {
            "type": "resource_of",
            "source_artifact_id": f"resource.{page}",
            "target_artifact_id": f"fragment.{page}",
            "role": "image",
            "ordinal": page - 1,
            "page_resource": {"source_page": page},
        }
        for page in range(1, 5)
    )
    relations.append(
        {
            "type": "resource_of",
            "source_artifact_id": "resource.5",
            "target_artifact_id": "document.main",
            "role": "image",
            "ordinal": 4,
        }
    )
    diagnostics = [
        {
            "severity": "warning",
            "code": "OCR-BEST-EFFORT.no_text",
            "message": "best effort",
            "artifact_id": f"fragment.{page}",
        }
        for page in (2, 3)
    ]
    diagnostics.append(
        {
            "severity": "warning",
            "code": "resource_page_unresolved",
            "message": "page unresolved",
            "artifact_id": "resource.5",
        }
    )
    return {
        "method": "task/completed",
        "params": {
            "bundle": {
                "schema": "docwen.artifact_bundle.v3",
                "layout_schema": "docwen.artifact_layout.v1",
                "task_id": "task.physical",
                "producer": {"name": "DocWen", "version": "0.9.0"},
                "artifacts": artifacts,
                "entries": [{"artifact_id": "document.main", "role": "primary", "ordinal": 0, "preferred": True}],
                "relations": relations,
            },
            "diagnostics": diagnostics,
        },
    }


def _verify(terminal: dict[str, Any], root: Path) -> None:
    _verify_physical_page_bundle(
        terminal=terminal,
        task_id="task.physical",
        staging_root=root,
        page_count=4,
        resource_count=5,
        ocr_enabled=True,
        keep_images=True,
        expected_statuses=("success", "no_text", "recognition_failed", "success"),
    )


def test_packaged_physical_page_verifier_accepts_canonical_p4_k5(tmp_path: Path) -> None:
    _verify(_canonical_terminal(tmp_path), tmp_path)


def _replace_payload(terminal: dict[str, Any], root: Path, artifact_id: str, payload: bytes) -> None:
    artifact = next(item for item in terminal["params"]["bundle"]["artifacts"] if item["artifact_id"] == artifact_id)
    (root / artifact["locator"]).write_bytes(payload)
    artifact.update(size_bytes=len(payload), sha256=hashlib.sha256(payload).hexdigest())


def test_packaged_physical_page_verifier_accepts_owned_images_without_inventing_ocr(tmp_path: Path) -> None:
    terminal = _canonical_terminal(tmp_path)
    for page, body in enumerate(("page one", "", "", "page four"), 1):
        payload = f"![{page}](resource.{page}.bin)\n\n{body}\n".encode()
        _replace_payload(terminal, tmp_path, f"fragment.{page}", payload)
    _verify(terminal, tmp_path)


@pytest.mark.parametrize(
    "payload",
    [
        b"![2](missing.png)\n",
        b"![2](resource.1.bin)\n",
        b"![2](https://example.invalid/resource.2.bin)\n",
        b"![2](resource.2.bin)\n![2](resource.2.bin)\n",
        b"![2](resource.2.bin)\nInvented OCR text\n",
    ],
)
def test_packaged_physical_page_verifier_rejects_damaged_page_navigation(tmp_path: Path, payload: bytes) -> None:
    terminal = _canonical_terminal(tmp_path)
    _replace_payload(terminal, tmp_path, "fragment.2", payload)
    with pytest.raises(RuntimeError):
        _verify(terminal, tmp_path)


@pytest.mark.parametrize("damage", ["empty_success", "duplicate_ocr"])
def test_packaged_physical_page_images_do_not_hide_ocr_damage(tmp_path: Path, damage: str) -> None:
    terminal = _canonical_terminal(tmp_path)
    if damage == "empty_success":
        _replace_payload(terminal, tmp_path, "fragment.1", b"![1](resource.1.bin)\n")
    else:
        _replace_payload(terminal, tmp_path, "fragment.1", b"![1](resource.1.bin)\n\npage one\n")
        _replace_payload(terminal, tmp_path, "document.main", b"primary only\n\npage one\n")
    with pytest.raises(RuntimeError):
        _verify(terminal, tmp_path)


@pytest.mark.parametrize("encoded", [False, True])
def test_packaged_physical_page_rejects_absolute_owned_image(tmp_path: Path, encoded: bool) -> None:
    terminal = _canonical_terminal(tmp_path)
    resource = tmp_path / "resource.2.bin"
    # A rooted POSIX spelling also resolves on the current Windows drive.
    target = resource.as_posix().removeprefix(resource.drive)
    if encoded:
        target = quote(target, safe="")
    _replace_payload(terminal, tmp_path, "fragment.2", f"![2]({target})\n".encode())
    with pytest.raises(RuntimeError, match="navigation_not_portable"):
        _verify(terminal, tmp_path)


def test_packaged_physical_page_accepts_encoded_relative_parent_image(tmp_path: Path) -> None:
    terminal = _canonical_terminal(tmp_path)
    bundle = terminal["params"]["bundle"]
    image = next(item for item in bundle["artifacts"] if item["artifact_id"] == "resource.2")
    target = tmp_path / "图 # % (2).png"
    (tmp_path / image["locator"]).rename(target)
    image.update(locator=target.name, logical_path=target.name, suggested_name=target.name)
    fragment = next(item for item in bundle["artifacts"] if item["artifact_id"] == "fragment.2")
    (tmp_path / "page").mkdir()
    (tmp_path / fragment["locator"]).rename(tmp_path / "page/2.md")
    fragment.update(locator="page/2.md", logical_path="page/2.md", suggested_name="2.md")
    _replace_payload(terminal, tmp_path, "fragment.2", f"![2](../{quote(target.name, safe='')})\n".encode())
    _verify(terminal, tmp_path)


def _tiff_terminal(root: Path, *, ocr: bool, images: bool) -> dict[str, Any]:
    terminal = _canonical_terminal(root)
    bundle = terminal["params"]["bundle"]
    bundle["artifacts"] = [
        item
        for item in bundle["artifacts"]
        if item["artifact_id"] != "resource.5"
        and (item["kind"] != "fragment" or ocr)
        and (item["kind"] != "resource" or images)
    ]
    ids = {item["artifact_id"] for item in bundle["artifacts"]}
    bundle["relations"] = [item for item in bundle["relations"] if item["source_artifact_id"] in ids]
    for relation in bundle["relations"]:
        if not ocr:
            relation["target_artifact_id"] = "document.main"
    terminal["params"]["diagnostics"] = [
        item for item in terminal["params"]["diagnostics"] if item.get("artifact_id") in ids
    ]
    links = [f"[{page}](fragment.{page}.bin)" if ocr else f"![{page}](resource.{page}.bin)" for page in range(1, 5)]
    _replace_payload(
        terminal, root, "document.main", ("source context\n" + "\n".join(links if ocr or images else [])).encode()
    )
    if ocr and images:
        for page, body in enumerate(("page one", "", "", "page four"), 1):
            _replace_payload(terminal, root, f"fragment.{page}", f"![{page}](resource.{page}.bin)\n\n{body}\n".encode())
    return terminal


def _verify_tiff(terminal: dict[str, Any], root: Path, *, ocr: bool, images: bool) -> None:
    _verify_physical_page_bundle(
        terminal=terminal,
        task_id="task.physical",
        staging_root=root,
        page_count=4,
        resource_count=4,
        ocr_enabled=ocr,
        keep_images=images,
        expected_statuses=("success", "no_text", "recognition_failed", "success") if ocr else None,
        tiff_navigation=True,
    )


@pytest.mark.parametrize("ocr", [False, True])
@pytest.mark.parametrize("images", [False, True])
def test_tiff_release_verifier_requires_navigation_for_all_four_combinations(
    tmp_path: Path, ocr: bool, images: bool
) -> None:
    _verify_tiff(_tiff_terminal(tmp_path, ocr=ocr, images=images), tmp_path, ocr=ocr, images=images)


@pytest.mark.parametrize(
    ("ocr", "images", "artifact", "payload"),
    [
        (False, True, "document.main", "---\ntitle: old empty primary\n---\n"),
        (True, False, "document.main", "source context\n"),
        (True, True, "fragment.2", ""),
        (True, True, "fragment.3", ""),
        (True, False, "document.main", "[1](fragment.1.bin)\n[3](fragment.3.bin)\n[4](fragment.4.bin)"),
        (
            True,
            False,
            "document.main",
            "[1](fragment.1.bin)\n[2](fragment.2.bin)\n[2](fragment.2.bin)\n[4](fragment.4.bin)",
        ),
        (
            True,
            False,
            "document.main",
            "[2](fragment.2.bin)\n[1](fragment.1.bin)\n[3](fragment.3.bin)\n[4](fragment.4.bin)",
        ),
        (
            False,
            True,
            "document.main",
            "![2](resource.2.bin)\n![1](resource.1.bin)\n![3](resource.3.bin)\n![4](resource.4.bin)",
        ),
        (False, False, "document.main", "[unexpected](fragment.1.bin)"),
    ],
)
def test_tiff_release_verifier_rejects_missing_duplicate_and_reordered_navigation(
    tmp_path: Path, ocr: bool, images: bool, artifact: str, payload: str
) -> None:
    terminal = _tiff_terminal(tmp_path, ocr=ocr, images=images)
    _replace_payload(terminal, tmp_path, artifact, payload.encode())
    with pytest.raises(RuntimeError, match="tiff_navigation_invalid"):
        _verify_tiff(terminal, tmp_path, ocr=ocr, images=images)


@pytest.mark.parametrize(
    "damage",
    [
        "extra_relation",
        "wrong_resource_source",
        "nonempty_failure",
        "duplicate_primary",
        "missing_diagnostic",
        "dangling_diagnostic",
        "unsafe_locator",
        "unknown_page_field",
    ],
)
def test_packaged_physical_page_verifier_rejects_semantic_damage(tmp_path: Path, damage: str) -> None:
    terminal = _canonical_terminal(tmp_path)
    damaged = copy.deepcopy(terminal)
    bundle = damaged["params"]["bundle"]
    if damage == "extra_relation":
        bundle["relations"].append(copy.deepcopy(bundle["relations"][0]))
    elif damage == "wrong_resource_source":
        relation = next(item for item in bundle["relations"] if item["source_artifact_id"] == "resource.1")
        relation["source_artifact_id"] = "resource.2"
    elif damage == "nonempty_failure":
        artifact = next(item for item in bundle["artifacts"] if item["artifact_id"] == "fragment.3")
        path = tmp_path / artifact["locator"]
        path.write_bytes(b"must stay empty")
        artifact["size_bytes"] = path.stat().st_size
        artifact["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
    elif damage == "duplicate_primary":
        primary = next(item for item in bundle["artifacts"] if item["artifact_id"] == "document.main")
        path = tmp_path / primary["locator"]
        path.write_bytes(path.read_bytes() + b"page one\n")
        primary["size_bytes"] = path.stat().st_size
        primary["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
    elif damage == "missing_diagnostic":
        damaged["params"]["diagnostics"] = [
            item for item in damaged["params"]["diagnostics"] if item.get("artifact_id") != "resource.5"
        ]
    elif damage == "dangling_diagnostic":
        damaged["params"]["diagnostics"].append(
            {
                "severity": "warning",
                "code": "other",
                "message": "dangling",
                "artifact_id": "artifact.missing",
            }
        )
    elif damage == "unsafe_locator":
        next(item for item in bundle["artifacts"] if item["artifact_id"] == "resource.1")["locator"] = "C:/escape.png"
    else:
        next(item for item in bundle["relations"] if item["source_artifact_id"] == "fragment.1")["page_fragment"][
            "unknown"
        ] = True

    with pytest.raises(RuntimeError):
        _verify(damaged, tmp_path)
