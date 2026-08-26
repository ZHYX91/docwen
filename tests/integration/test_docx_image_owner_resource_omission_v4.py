"""Cross-plugin gates for resource-less DOCX image-owner recovery."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from PIL import Image

from docwen_application.bundle_mapping import build_bundle_draft
from docwen_plugin_markdown.document_semantics_v3 import analyze_markdown_semantics_v3
from tests.integration._round_trip_helper import _primary_path, _run

pytestmark = pytest.mark.integration


@pytest.fixture(scope="module")
def authenticated_owner_docx(tmp_path_factory: pytest.TempPathFactory, round_trip_runtime: Any) -> dict[str, Path]:
    root = tmp_path_factory.mktemp("docx-image-owner-v4")
    result: dict[str, Path] = {}
    for owner_kind in ("figure", "ordinary"):
        image_path = root / f"{owner_kind}.png"
        with Image.new("RGB", (2, 2), (32, 96, 160)) as image:
            image.save(image_path, format="PNG")
        source = (
            f"Figure: Recovered illustration ^figure-owner\n\n![pixel]({image_path.name})\n"
            if owner_kind == "figure"
            else f"![pixel]({image_path.name}) ^ordinary-owner\n"
        )
        source_path = root / f"{owner_kind}.md"
        source_path.write_text(source, encoding="utf-8")
        forward = _run(
            round_trip_runtime,
            f"docx-image-owner-{owner_kind}-forward",
            source_path,
            source_format="markdown",
            target_format="docx",
            output_dir=root / f"forward-{owner_kind}",
        )
        result[owner_kind] = _primary_path(forward)
    return result


@pytest.mark.parametrize("owner_kind", ["figure", "ordinary"])
@pytest.mark.parametrize(
    ("recognize_text", "preserve_resources", "ocr_placement"),
    [
        (False, False, "main_md"),
        (False, True, "main_md"),
        (True, False, "main_md"),
        (True, True, "main_md"),
        (True, False, "image_md"),
        (True, True, "image_md"),
    ],
)
def test_authenticated_owner_survives_six_bundle_combinations(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    round_trip_runtime: Any,
    authenticated_owner_docx: dict[str, Path],
    owner_kind: str,
    recognize_text: bool,
    preserve_resources: bool,
    ocr_placement: str,
) -> None:
    import docwen_core.text.ocr as ocr

    observed_ocr_inputs: list[Path] = []

    def _run_ocr(path: str, **_kwargs: Any) -> ocr.OcrOutcome:
        observed = Path(path)
        assert observed.is_file()
        observed_ocr_inputs.append(observed)
        return ocr.OcrOutcome(ocr.OcrStatus.SUCCESS, text="authenticated owner OCR")

    monkeypatch.setattr(ocr, "run_ocr_outcome", _run_ocr)
    result = _run(
        round_trip_runtime,
        f"owner-{owner_kind}-{recognize_text}-{preserve_resources}-{ocr_placement}",
        authenticated_owner_docx[owner_kind],
        source_format="docx",
        target_format="md",
        output_dir=tmp_path / "reverse",
        options={
            "to_md_enable_ocr": recognize_text,
            "to_md_keep_images": preserve_resources,
            "image_mode": "file",
            "ocr_placement": ocr_placement,
        },
    )

    assert result.success, result.error
    primary = next(artifact for artifact in result.artifacts if artifact.is_primary)
    primary_text = Path(primary.staging_path).read_text(encoding="utf-8")
    images = [artifact for artifact in result.artifacts if artifact.kind == "image"]
    fragments = [artifact for artifact in result.artifacts if artifact.metadata.get("ocr") is True]
    analysis = analyze_markdown_semantics_v3(primary_text, input_id=primary.suggested_name)

    assert not analysis.has_errors, analysis.diagnostics
    if owner_kind == "figure":
        [owner] = [target for target in analysis.projection["targets"] if target["kind"] == "figure"]
        assert (owner["id"], owner["title"]) == ("figure-owner", "Recovered illustration")
    else:
        [owner] = analysis.projection["anchors"]
        assert (owner["id"], owner["block_kind"], owner["placement"]) == (
            "ordinary-owner",
            "image",
            "inline",
        )

    assert ("![image omitted]()" in primary_text) is (not preserve_resources)
    assert len(images) == int(preserve_resources)
    assert len(observed_ocr_inputs) == int(recognize_text)
    assert len(fragments) == int(recognize_text and ocr_placement == "image_md")
    assert ("authenticated owner OCR" in primary_text) is (recognize_text and ocr_placement == "main_md")
    assert primary.metadata["image_count"] == int(preserve_resources)

    warning_payloads = [
        diagnostic.to_dict()
        for diagnostic in result.diagnostics
        if diagnostic.code == "DOCX2MD-IMAGE-OWNER-RESOURCE-OMITTED"
    ]
    assert len(warning_payloads) == int(not preserve_resources)
    if warning_payloads:
        [warning] = warning_payloads
        assert warning["artifact_id"] == primary.artifact_id
        assert not {"evidence_schema", "source", "range", "related_ranges", "fixes"}.intersection(warning)
        assert primary.metadata["image_owner_resource_omitted_count"] == 1

    if owner_kind == "ordinary" and not preserve_resources:
        assert "![image omitted]() ^ordinary-owner" in primary_text
    if recognize_text and ocr_placement == "main_md":
        assert primary_text.index("![") < primary_text.index("authenticated owner OCR")
    if fragments:
        fragment_text = Path(fragments[0].staging_path).read_text(encoding="utf-8")
        assert "authenticated owner OCR" in fragment_text
        assert "![" not in fragment_text
        assert "Figure:" not in fragment_text
        assert "^figure-owner" not in fragment_text
        assert "^ordinary-owner" not in fragment_text
        assert "__img_" not in primary_text

    bundle = build_bundle_draft(
        profile="document_with_resources",
        output_media_type="text/markdown",
        artifacts=tuple(result.artifacts),
    )
    assert len(bundle.entries) == 1
    assert (bundle.entries[0].artifact_id, bundle.entries[0].role) == (primary.artifact_id, "primary")
    assert sum(artifact.kind == "document" for artifact in bundle.artifacts) == 1
    assert sum(artifact.kind == "resource" for artifact in bundle.artifacts) == int(preserve_resources) + 1
    assert sum(artifact.kind == "fragment" for artifact in bundle.artifacts) == int(
        recognize_text and ocr_placement == "image_md"
    )
    assert sum(relation.type == "resource_of" and relation.role == "image" for relation in bundle.relations) == int(
        preserve_resources
    )
    assert sum(relation.type == "resource_of" and relation.role == "manifest" for relation in bundle.relations) == 1
    assert sum(relation.type == "fragment_of" and relation.role == "ocr_text" for relation in bundle.relations) == int(
        recognize_text and ocr_placement == "image_md"
    )


@pytest.mark.parametrize("owner_kind", ["figure", "ordinary"])
def test_authenticated_owner_omit_fails_closed_without_partial_bundle(
    tmp_path: Path,
    round_trip_runtime: Any,
    authenticated_owner_docx: dict[str, Path],
    owner_kind: str,
) -> None:
    output_dir = tmp_path / "reverse"
    result = _run(
        round_trip_runtime,
        f"owner-{owner_kind}-omit-fail-closed",
        authenticated_owner_docx[owner_kind],
        source_format="docx",
        target_format="md",
        output_dir=output_dir,
        options={
            "to_md_enable_ocr": True,
            "to_md_keep_images": True,
            "image_mode": "omit",
        },
    )

    assert not result.success
    assert result.error is not None
    assert result.error.diagnostic_code == "DOCX2MD-PARSE-ERROR"
    assert result.error.message == "image_mode=omit cannot preserve an authenticated DOCX image owner"
    assert result.artifacts == []
    assert not output_dir.exists() or not any(path.is_file() for path in output_dir.rglob("*"))


@pytest.mark.parametrize("owner_kind", ["figure", "ordinary"])
def test_authenticated_owner_base64_remains_a_semantic_image_owner(
    tmp_path: Path,
    round_trip_runtime: Any,
    authenticated_owner_docx: dict[str, Path],
    owner_kind: str,
) -> None:
    result = _run(
        round_trip_runtime,
        f"owner-{owner_kind}-base64",
        authenticated_owner_docx[owner_kind],
        source_format="docx",
        target_format="md",
        output_dir=tmp_path / "reverse",
        options={
            "to_md_enable_ocr": False,
            "to_md_keep_images": True,
            "image_mode": "base64",
        },
    )

    assert result.success, result.error
    primary = next(artifact for artifact in result.artifacts if artifact.is_primary)
    primary_text = Path(primary.staging_path).read_text(encoding="utf-8")
    analysis = analyze_markdown_semantics_v3(primary_text, input_id=primary.suggested_name)

    assert not analysis.has_errors, analysis.diagnostics
    assert "data:image/png;base64," in primary_text
    assert "<!-- image omitted:" not in primary_text
    assert sum(artifact.kind == "image" for artifact in result.artifacts) == 1
    if owner_kind == "figure":
        [owner] = [target for target in analysis.projection["targets"] if target["kind"] == "figure"]
        assert (owner["id"], owner["title"]) == ("figure-owner", "Recovered illustration")
    else:
        [owner] = analysis.projection["anchors"]
        assert (owner["id"], owner["block_kind"], owner["placement"]) == (
            "ordinary-owner",
            "image",
            "inline",
        )
