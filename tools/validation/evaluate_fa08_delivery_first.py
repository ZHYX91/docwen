"""Evaluate VIS-201 outputs against the selected FA-08 delivery-first policy.

This is deliberately separate from ``probe_image_fa08_parity.py``.  The
original probe is the immutable pre-choice three-project oracle; it correctly
records that the old projects and the then-current project behaved equally,
but its orientation predicate requires the source orientation tag to survive.
The selected O-A policy instead requires applying that orientation to pixels
and writing orientation 1 (or no orientation tag).
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any

from PIL import Image

EXPECTED_INPUTS = {
    "B1": ("fa08-b1-exif-orientation-6.jpg", "2de0c4bf5b6710ac022e1ecadcdd7144ad91f20116608a9fe50912bac7c2e78f"),
    "B2": ("fa08-b2-animated-rgb.gif", "706ad21615639cc4677128fd8083cedd8f1629238056b5d5bab8e349b7a9e3e2"),
    "N1": ("fa08-n1-iphone14-bus.jpg", "a416b2d5265506396eb54784613c6cb7791d3d05d6a2c1c770d5c217e8426a43"),
    "N2": ("fa08-n2-hainan-road-sign.jpg", "f2f852060f4980d140e10dcf4bcb10852d8cb4ba85e7961e327d1063458b3d72"),
}
EXPECTED_STAGE_CONTRACT_SHA256 = "c1299d3a4e76007ef3ad21efecf8b9ad1773cb55562ebdfee39c652bfd7f8909"
FLAT_CASES = ("B1", "N1", "N2")
FLAT_ROUTES = ("jpg", "webp")
UNCHANGED_ROUTES = ("markdown-file", "markdown-image-md")
SUPPORTED_EXIF = {271, 272, 274, 305, 306}
PRESERVED_EXIF = {271, 272, 305, 306}
MPO_WARNING = "IMG2PDF-MPO-AUXILIARY-FRAMES"
MINIMUM_DISPLAY_SIMILARITY = 0.95
N1_ACCEPTED_AUXILIARY_SCORE = 0.79880938


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"expected JSON object: {path}")
    return value


def _check(checks: dict[str, bool], name: str, value: Any) -> None:
    checks[name] = bool(value)


def _flat_evaluation(
    *,
    case: str,
    route: str,
    source: Path,
    slot: dict[str, Any],
) -> dict[str, Any]:
    checks: dict[str, bool] = {}
    projection = slot.get("projection") or {}
    artifact = projection.get("artifact") or {}
    similarity = projection.get("display_similarity") or {}
    inventory = slot.get("artifact_inventory") or []
    residue = (slot.get("processes") or {}).get("residue_added") or []

    _check(checks, "success", slot.get("success") and slot.get("returncode") == 0)
    _check(checks, "single_artifact", len(inventory) == 1)
    _check(checks, "final_placement", slot.get("final_placement"))
    _check(checks, "no_process_residue", not residue)
    _check(checks, "display_dimensions", similarity.get("same_dimensions"))
    _check(checks, "display_similarity", float(similarity.get("score", -1)) >= MINIMUM_DISPLAY_SIMILARITY)
    _check(checks, "normalized_orientation", artifact.get("orientation") in (None, 1))

    source_exif: dict[int, Any]
    source_icc: bytes
    with Image.open(source) as source_image:
        source_exif = dict(source_image.getexif())
        source_icc = source_image.info.get("icc_profile") or b""

    output_exif: dict[int, Any] = {}
    output_icc = b""
    output_path: Path | None = None
    if len(inventory) == 1:
        output_path = Path(inventory[0]["path"])
        _check(checks, "artifact_exists", output_path.is_file())
        if output_path.is_file():
            _check(checks, "artifact_sha256", _sha256(output_path) == inventory[0].get("sha256"))
            with Image.open(output_path) as output_image:
                output_exif = dict(output_image.getexif())
                output_icc = output_image.info.get("icc_profile") or b""
    else:
        _check(checks, "artifact_exists", False)
        _check(checks, "artifact_sha256", False)

    _check(checks, "supported_exif_only", set(output_exif).issubset(SUPPORTED_EXIF))
    _check(
        checks,
        "preserved_supported_exif",
        all(output_exif.get(tag) == source_exif.get(tag) for tag in PRESERVED_EXIF if tag in source_exif),
    )
    _check(checks, "orientation_value", output_exif.get(274) in (None, 1))
    _check(checks, "icc_exact", output_icc == source_icc)

    return {
        "case": case,
        "route": route,
        "checks": checks,
        "pass": all(checks.values()),
        "display_similarity": similarity.get("score"),
        "source_orientation": source_exif.get(274),
        "output_orientation": output_exif.get(274),
        "source_selected_exif": {str(tag): source_exif[tag] for tag in sorted(PRESERVED_EXIF) if tag in source_exif},
        "output_exif": {str(tag): output_exif[tag] for tag in sorted(output_exif)},
        "source_icc_sha256": hashlib.sha256(source_icc).hexdigest() if source_icc else None,
        "output_icc_sha256": hashlib.sha256(output_icc).hexdigest() if output_icc else None,
    }


def _pdf_evaluation(*, case: str, slot: dict[str, Any]) -> dict[str, Any]:
    checks: dict[str, bool] = {}
    projection = slot.get("projection") or {}
    pages = projection.get("pages") or []
    warnings = (slot.get("current_json") or {}).get("warnings") or []
    matching_warnings = [item for item in warnings if item.get("code") == MPO_WARNING]
    frame_count = projection.get("source_frame_count")
    page_count = projection.get("page_count")
    residue = (slot.get("processes") or {}).get("residue_added") or []

    _check(checks, "success", slot.get("success") and slot.get("returncode") == 0)
    _check(checks, "final_placement", slot.get("final_placement"))
    _check(checks, "no_process_residue", not residue)
    _check(checks, "all_frames_delivered", frame_count == 2 and page_count == frame_count and len(pages) == frame_count)
    _check(checks, "exactly_one_mpo_warning", len(matching_warnings) == 1)
    warning_message = matching_warnings[0].get("message", "") if matching_warnings else ""
    _check(checks, "warning_names_count", "all 2 MPO frames" in warning_message)
    _check(checks, "warning_requires_review", "review every page" in warning_message.lower())
    _check(
        checks,
        "embedded_frames_source_exact",
        bool(pages)
        and all(
            page.get("embedded_images")
            and page["embedded_images"][0].get("decoded_similarity_to_source_frame", {}).get("same_dimensions")
            and page["embedded_images"][0].get("decoded_similarity_to_source_frame", {}).get("score") == 1.0
            for page in pages
        ),
    )

    render_scores = [page.get("render", {}).get("score") for page in pages]
    if case == "N1":
        _check(
            checks,
            "accepted_auxiliary_render_boundary_exact",
            len(render_scores) == 2 and abs(float(render_scores[1]) - N1_ACCEPTED_AUXILIARY_SCORE) <= 0.0000005,
        )
    else:
        _check(
            checks,
            "page_render_similarity",
            len(render_scores) == 2 and all(float(score) >= MINIMUM_DISPLAY_SIMILARITY for score in render_scores),
        )

    return {
        "case": case,
        "route": "pdf-original",
        "checks": checks,
        "pass": all(checks.values()),
        "source_frame_count": frame_count,
        "page_count": page_count,
        "render_scores": render_scores,
        "warning_codes": [item.get("code") for item in warnings],
    }


def evaluate(evidence_root: Path) -> dict[str, Any]:
    evidence_root = evidence_root.resolve()
    result_path = evidence_root / "probe-result.json"
    result = _load_json(result_path)
    manifest = _load_json(evidence_root / "probe-input-manifest.json")
    projects = result.get("projects") or {}
    current = projects.get("docwen-current") or {}

    source_checks: dict[str, bool] = {}
    sources: dict[str, Path] = {}
    for case, (name, expected_hash) in EXPECTED_INPUTS.items():
        source = evidence_root / "inputs" / name
        sources[case] = source
        _check(source_checks, f"{case}_exists", source.is_file())
        _check(source_checks, f"{case}_sha256", source.is_file() and _sha256(source) == expected_hash)
    contract_path = evidence_root / "stage-contract.json"
    _check(source_checks, "stage_contract_sha256", _sha256(contract_path) == EXPECTED_STAGE_CONTRACT_SHA256)
    _check(
        source_checks,
        "manifest_stage_contract_sha256",
        manifest.get("stage_contract_sha256") == EXPECTED_STAGE_CONTRACT_SHA256,
    )

    flat = [
        _flat_evaluation(case=case, route=route, source=sources[case], slot=current[case][route])
        for case in FLAT_CASES
        for route in FLAT_ROUTES
    ]
    pdf = [_pdf_evaluation(case=case, slot=current[case]["pdf-original"]) for case in ("N1", "N2")]

    reused_checks: dict[str, bool] = {}
    route_acceptance = result.get("route_acceptance") or {}
    current_acceptance = route_acceptance.get("docwen-current") or {}
    for case in EXPECTED_INPUTS:
        for route in UNCHANGED_ROUTES:
            _check(reused_checks, f"{case}_{route}", (current_acceptance.get(case) or {}).get(route))
    b2_pdf = ((current.get("B2") or {}).get("pdf-original") or {}).get("projection") or {}
    _check(
        reused_checks,
        "B2_pdf_three_ordered_pages",
        b2_pdf.get("source_frame_count") == 3 and b2_pdf.get("page_count") == 3 and len(b2_pdf.get("pages") or []) == 3,
    )
    _check(
        reused_checks,
        "B2_no_process_residue",
        all(
            not (((current.get("B2") or {}).get(route) or {}).get("processes") or {}).get("residue_added")
            for route in ("jpg", "webp", "pdf-original", *UNCHANGED_ROUTES)
        ),
    )

    accepted_boundaries = {
        "M-B": {
            "accepted": True,
            "scope": "MPO auxiliary/gain-map/secondary frames remain delivered as PDF pages with a typed warning",
            "N1_page_2_render_score": pdf[0]["render_scores"][1],
            "minimum_general_display_similarity": MINIMUM_DISPLAY_SIMILARITY,
        }
    }
    pass_value = (
        all(source_checks.values())
        and all(item["pass"] for item in flat)
        and all(item["pass"] for item in pdf)
        and all(reused_checks.values())
    )
    return {
        "stage_id": "VIS-2026-07-23-201",
        "policy": "FA-08=O-A,E-A,C-A,M-B",
        "evidence_root": str(evidence_root),
        "probe_result_sha256": _sha256(result_path),
        "source_checks": source_checks,
        "flat_outputs": flat,
        "mpo_pdf_outputs": pdf,
        "reused_unchanged_evidence": reused_checks,
        "accepted_boundaries": accepted_boundaries,
        "affected_slot_count": len(flat) + len(pdf),
        "affected_slots_passed": sum(item["pass"] for item in (*flat, *pdf)),
        "pass": pass_value,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--evidence-root", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    value = evaluate(args.evidence_root)
    serialized = json.dumps(value, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(serialized, encoding="utf-8")
    print(serialized, end="")
    return 0 if value["pass"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
