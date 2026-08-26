"""Keep DocWen production source independent from product-specific semantics."""

from __future__ import annotations

from pathlib import Path

import pytest

pytestmark = pytest.mark.contract

REPO_ROOT = Path(__file__).resolve().parents[2]
PRODUCTION_ROOT = REPO_ROOT / "packages"
FORBIDDEN_MARKERS = ("wenleaf", "weftext", "pkwf")


def test_production_python_source_has_no_external_product_semantics() -> None:
    violations: list[str] = []

    for source_path in sorted(PRODUCTION_ROOT.glob("**/src/**/*.py")):
        source = source_path.read_text(encoding="utf-8").lower()
        for marker in FORBIDDEN_MARKERS:
            if marker in source:
                relative_path = source_path.relative_to(REPO_ROOT).as_posix()
                violations.append(f"{relative_path}: {marker}")

    assert not violations, "External-product production markers found:\n" + "\n".join(violations)


def test_release_authority_and_build_pipeline_are_product_neutral() -> None:
    release_files = [
        REPO_ROOT / "contracts/baselines/docwen-v4-candidate-authority.json",
        REPO_ROOT / "contracts/schemas/docwen.candidate_receipt.v4.schema.json",
        REPO_ROOT / "scripts/release/build_v4_candidate_staging.py",
        REPO_ROOT / "scripts/release/build_v4_package_input.py",
        REPO_ROOT / "scripts/release/seal_v4_candidate.py",
        REPO_ROOT / "scripts/release/v4_candidate_contract.py",
        REPO_ROOT / "scripts/release/v4_package_input_contract.py",
    ]
    violations = [
        f"{path.relative_to(REPO_ROOT).as_posix()}: {marker}"
        for path in release_files
        for marker in FORBIDDEN_MARKERS
        if marker in path.read_text(encoding="utf-8").casefold()
    ]

    assert not violations, "External-product release bindings found:\n" + "\n".join(violations)


def test_v4_cross_reference_token_is_provider_neutral() -> None:
    semantics = (
        PRODUCTION_ROOT / "plugins" / "markdown" / "src" / "docwen_plugin_markdown" / "document_semantics_v3.py"
    ).read_text(encoding="utf-8")

    assert "@[[" in semantics
    lowered = semantics.lower()
    assert not any(marker in lowered for marker in FORBIDDEN_MARKERS)
