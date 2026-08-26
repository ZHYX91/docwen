from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, cast

from scripts.release import candidate_evidence
from scripts.release import v4_evidence_contract as evidence_contract

AUTHORITY_PATH = "contracts/baselines/docwen-v4-candidate-authority.json"
RECEIPT_SCHEMA_PATH = "contracts/schemas/docwen.candidate_receipt.v4.schema.json"
EVIDENCE_SCHEMA_PATH = "contracts/schemas/docwen.candidate_evidence_index.v4.schema.json"
AUTHORITY_SCHEMA = "docwen.candidate_authority.v4"
RECEIPT_SCHEMA = "docwen.candidate_receipt.v4"
EVIDENCE_SCHEMA = "docwen.candidate_evidence_index.v4"
ACTIVE_SEMANTICS = "docwen.markdown_semantics.v3"
PRODUCT_VERSION = "0.9.0"
_HEX40 = re.compile(r"^[0-9a-f]{40}$")
_HEX64 = re.compile(r"^[0-9a-f]{64}$")
_CANDIDATE_ID = re.compile(r"^docwen-0\.9\.0-v4-[0-9]{8}T[0-9]{6}Z-[0-9a-f]{12}$")
_EXECUTABLE_SUFFIXES = frozenset(
    {
        ".bat",
        ".cmd",
        ".com",
        ".dll",
        ".dylib",
        ".exe",
        ".msi",
        ".ps1",
        ".pyc",
        ".pyd",
        ".pyo",
        ".pyz",
        ".sh",
        ".so",
    }
)
REQUIRED_INVALID_CASES = {
    "dot": "invalid-id-dot",
    "slash": "invalid-id-slash",
    "underscore": "invalid-id-underscore",
    "over128": "invalid-id-too-long",
    "duplicate": "invalid-duplicate",
    "kindMismatch": "invalid-kind-mismatch",
}
REQUIRED_LAYERS = (
    "source_oracle",
    "machine_wire",
    "source_wire_comparison",
    "packaged",
    "roundtrip",
    "headless_ooxml",
    "word_host",
    "wps_host",
    "libreoffice_host",
)
POINTER_ROLE_BY_LAYER = evidence_contract.POINTER_ROLE_BY_LAYER
WIRE_SCHEMA_PATH = "contracts/schemas/docwen.machine.diagnostic_evidence.v1.schema.json"
PACKAGE_MANIFEST_PATH = "_evidence/package-manifest.json"


class V4CandidateContractError(RuntimeError):
    """A v4 candidate authority or receipt contract failed closed."""


def json_bytes(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n").encode("utf-8")


def payload_sha256(value: object) -> str:
    canonical = json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")
    return hashlib.sha256(canonical).hexdigest()


def sha256_file(path: Path) -> str:
    try:
        return evidence_contract.sha256_file(path)
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4CandidateContractError(str(exc)) from exc


def write_json_exclusive(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as handle:
        handle.write(json_bytes(value))
        handle.flush()
        os.fsync(handle.fileno())


def read_json_object(path: Path, *, label: str) -> dict[str, Any]:
    try:
        value, _ = evidence_contract.read_json_object(path, relative_to=path.parent, label=label)
        return value
    except (OSError, evidence_contract.V4EvidenceContractError) as exc:
        raise V4CandidateContractError(str(exc)) from exc


def read_json_object_with_identity(
    path: Path,
    *,
    relative_to: Path,
    label: str,
    expected_sha256: str | None = None,
) -> tuple[dict[str, Any], dict[str, object]]:
    try:
        return evidence_contract.read_json_object(
            path,
            relative_to=relative_to,
            label=label,
            expected_sha256=expected_sha256,
        )
    except (OSError, evidence_contract.V4EvidenceContractError) as exc:
        raise V4CandidateContractError(str(exc)) from exc


def file_identity(path: Path, *, relative_to: Path) -> dict[str, str | int]:
    try:
        return cast(dict[str, str | int], evidence_contract.file_identity(path, relative_to=relative_to))
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4CandidateContractError(str(exc)) from exc


def capture_tree_stable(root: Path) -> dict[str, object]:
    try:
        return evidence_contract.capture_tree_stable(root)
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4CandidateContractError(str(exc)) from exc


def identity_from_tree(path: Path, *, relative_to: Path, stable_tree: Mapping[str, object]) -> dict[str, str | int]:
    try:
        root = candidate_evidence._safe_existing_directory(relative_to, label="identity_root")
        safe = evidence_contract.safe_regular_file(path, label="identity_file")
    except (candidate_evidence.EvidenceError, evidence_contract.V4EvidenceContractError) as exc:
        raise V4CandidateContractError(str(exc)) from exc
    try:
        relative = safe.relative_to(root).as_posix()
    except ValueError as exc:
        raise V4CandidateContractError(f"identity_outside_root:{safe}:{root}") from exc
    try:
        evidence_contract.validate_relative_path(relative, label="identity")
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4CandidateContractError(str(exc)) from exc
    raw_files = stable_tree.get("files")
    if not isinstance(raw_files, list):
        raise V4CandidateContractError("stable_tree_files_invalid")
    matches = [item for item in raw_files if isinstance(item, dict) and item.get("path") == relative]
    if len(matches) != 1:
        raise V4CandidateContractError(f"identity_missing_from_stable_tree:{relative}")
    item = matches[0]
    size = item.get("size")
    digest = item.get("sha256")
    if not isinstance(size, int) or not isinstance(digest, str) or not _HEX64.fullmatch(digest):
        raise V4CandidateContractError(f"identity_invalid_in_stable_tree:{relative}")
    return {"relativePath": relative, "bytes": size, "sha256": digest}


def _git(repo: Path, *arguments: str, check: bool = True) -> subprocess.CompletedProcess[bytes]:
    completed = subprocess.run(
        ["git", "-C", str(repo), *arguments],
        check=False,
        capture_output=True,
    )
    if check and completed.returncode:
        stderr = completed.stderr.decode("utf-8", errors="replace").strip()
        raise V4CandidateContractError(f"git_failed:{arguments}:{stderr}")
    return completed


def _git_text(repo: Path, *arguments: str) -> str:
    try:
        return _git(repo, *arguments).stdout.decode("utf-8", errors="strict").strip()
    except UnicodeDecodeError as exc:
        raise V4CandidateContractError(f"git_output_not_utf8:{arguments}") from exc


def clean_git_identity(repo: Path, *, label: str) -> dict[str, str]:
    root = Path(_git_text(repo, "rev-parse", "--show-toplevel")).resolve(strict=True)
    if root != repo.resolve(strict=True):
        raise V4CandidateContractError(f"{label}_repository_root_mismatch:{root}")
    status = _git(repo, "status", "--porcelain=v2", "-z").stdout
    if status:
        raise V4CandidateContractError(f"{label}_source_not_clean")
    commit = _git_text(repo, "rev-parse", "HEAD")
    tree = _git_text(repo, "rev-parse", "HEAD^{tree}")
    if not _HEX40.fullmatch(commit) or not _HEX40.fullmatch(tree):
        raise V4CandidateContractError(f"{label}_git_identity_invalid")
    return {"commit": commit, "tree": tree}


def verify_baseline_ancestry(
    repo: Path, baseline: Mapping[str, object], final: Mapping[str, object], *, label: str
) -> None:
    commit = str(baseline.get("commit", ""))
    tree = str(baseline.get("tree", ""))
    final_commit = str(final.get("commit", ""))
    if not _HEX40.fullmatch(commit) or not _HEX40.fullmatch(tree):
        raise V4CandidateContractError(f"{label}_baseline_identity_invalid")
    if _git_text(repo, "rev-parse", f"{commit}^{{tree}}") != tree:
        raise V4CandidateContractError(f"{label}_baseline_tree_mismatch")
    if _git(repo, "merge-base", "--is-ancestor", commit, final_commit, check=False).returncode:
        raise V4CandidateContractError(f"{label}_baseline_not_ancestor")


def ignored_source_inputs(repo: Path) -> list[str]:
    raw = _git(repo, "ls-files", "--others", "--ignored", "--exclude-standard", "-z").stdout
    try:
        paths = [item.decode("utf-8", errors="strict").replace("\\", "/") for item in raw.split(b"\0") if item]
    except UnicodeDecodeError as exc:
        raise V4CandidateContractError("ignored_input_path_not_utf8") from exc
    return sorted(paths)


def ignored_executable_inputs(repo: Path) -> list[str]:
    raw = _git(repo, "ls-files", "--others", "--ignored", "--exclude-standard", "-z").stdout
    try:
        paths = [item.decode("utf-8", errors="strict").replace("\\", "/") for item in raw.split(b"\0") if item]
    except UnicodeDecodeError as exc:
        raise V4CandidateContractError("ignored_input_path_not_utf8") from exc
    return sorted(path for path in paths if Path(path).suffix.lower() in _EXECUTABLE_SUFFIXES)


def load_authority(repo: Path) -> tuple[dict[str, Any], dict[str, str | int]]:
    authority_path = repo / AUTHORITY_PATH
    authority = read_json_object(authority_path, label="candidate_authority")
    if authority.get("schema") != AUTHORITY_SCHEMA or authority.get("productVersion") != PRODUCT_VERSION:
        raise V4CandidateContractError("candidate_authority_identity_mismatch")
    active = authority.get("activeSemantics")
    if not isinstance(active, dict) or active.get("id") != ACTIVE_SEMANTICS:
        raise V4CandidateContractError("candidate_authority_active_semantics_mismatch")
    exclusions = authority.get("exclusions")
    if not isinstance(exclusions, list) or {str(item.get("id")) for item in exclusions if isinstance(item, dict)} != {
        "docwen.markdown_semantics.v1",
        "docwen.markdown_semantics.v2",
        "docwen-v3-dirty-audit-head",
        "Attempt04",
    }:
        raise V4CandidateContractError("candidate_authority_exclusions_mismatch")
    return authority, file_identity(authority_path, relative_to=repo)


def final_spec_identity(repo: Path, authority: Mapping[str, object]) -> dict[str, object]:
    raw_paths = authority.get("finalSpecFiles")
    if not isinstance(raw_paths, list) or not raw_paths:
        raise V4CandidateContractError("final_spec_files_missing")
    files: list[dict[str, str | int]] = []
    for raw in raw_paths:
        if not isinstance(raw, str) or Path(raw).is_absolute() or ".." in Path(raw).parts:
            raise V4CandidateContractError(f"final_spec_path_invalid:{raw}")
        files.append(file_identity(repo / raw, relative_to=repo))
    files.sort(key=lambda item: str(item["relativePath"]))
    return {"files": files, "sha256": payload_sha256(files)}


def oracle_manifest_identity(repo: Path, authority: Mapping[str, object]) -> dict[str, str | int]:
    active = cast(Mapping[str, object], authority["activeSemantics"])
    raw_path = active.get("oracleManifestPath")
    if not isinstance(raw_path, str):
        raise V4CandidateContractError("oracle_manifest_path_missing")
    manifest_path = repo / raw_path
    manifest = read_json_object(manifest_path, label="oracle_manifest")
    oracle = manifest.get("oracle")
    if not isinstance(oracle, dict) or oracle.get("semantics") != ACTIVE_SEMANTICS:
        raise V4CandidateContractError("oracle_manifest_semantics_mismatch")
    if manifest.get("final_spec_baseline") != authority.get("docwenSpecBaseline"):
        raise V4CandidateContractError("oracle_manifest_spec_baseline_mismatch")
    return file_identity(manifest_path, relative_to=repo)


def validate_candidate_id(candidate_id: str, *, docwen_commit: str) -> str:
    if not _CANDIDATE_ID.fullmatch(candidate_id) or not candidate_id.endswith(f"-{docwen_commit[:12]}"):
        raise V4CandidateContractError("v4_candidate_id_invalid")
    return candidate_id


def validate_evidence_index(
    value: Mapping[str, object], *, candidate_id: str
) -> tuple[list[dict[str, object]], dict[str, str]]:
    if value.get("schema") != EVIDENCE_SCHEMA or value.get("candidateId") != candidate_id:
        raise V4CandidateContractError("evidence_index_identity_mismatch")
    statuses = value.get("layerStatus")
    records = value.get("records")
    if not isinstance(statuses, dict) or set(statuses) != set(REQUIRED_LAYERS):
        raise V4CandidateContractError("evidence_layer_status_mismatch")
    if any(statuses[layer] != "passed" for layer in REQUIRED_LAYERS[:6]):
        raise V4CandidateContractError("required_non_host_evidence_not_passed")
    if any(statuses[layer] not in {"passed", "not_run"} for layer in REQUIRED_LAYERS[6:]):
        raise V4CandidateContractError("host_evidence_status_invalid")
    if not isinstance(records, list) or not records or not all(isinstance(item, dict) for item in records):
        raise V4CandidateContractError("evidence_records_invalid")
    typed_records = cast(list[dict[str, object]], records)
    record_keys: set[tuple[str, str]] = set()
    record_paths: set[str] = set()
    for record in typed_records:
        case_id = record.get("caseId")
        layer = record.get("layer")
        relative_path = record.get("relativePath")
        pointers = record.get("identityPointers")
        if not isinstance(case_id, str) or not isinstance(layer, str) or layer not in POINTER_ROLE_BY_LAYER:
            raise V4CandidateContractError("evidence_record_identity_invalid")
        if not isinstance(relative_path, str) or not isinstance(pointers, list) or len(pointers) != 1:
            raise V4CandidateContractError(f"evidence_record_shape_invalid:{case_id}")
        pointer = pointers[0]
        if not isinstance(pointer, dict) or pointer.get("role") != POINTER_ROLE_BY_LAYER[layer]:
            raise V4CandidateContractError(f"evidence_pointer_role_mismatch:{case_id}:{layer}")
        key = (layer, case_id)
        if key in record_keys or relative_path in record_paths:
            raise V4CandidateContractError(f"duplicate_evidence_record:{layer}:{case_id}")
        record_keys.add(key)
        record_paths.add(relative_path)
    record_ids: dict[str, str] = {}
    for dimension, case_id in REQUIRED_INVALID_CASES.items():
        matches = [
            item for item in typed_records if item.get("caseId") == case_id and item.get("layer") == "source_oracle"
        ]
        if len(matches) != 1:
            raise V4CandidateContractError(f"independent_invalid_record_missing:{dimension}")
        record_ids[dimension] = case_id
    for layer in REQUIRED_LAYERS:
        matches = [item for item in typed_records if item.get("layer") == layer]
        if layer in REQUIRED_LAYERS[6:]:
            expected = 1 if statuses[layer] == "passed" else 0
            if len(matches) != expected:
                raise V4CandidateContractError(f"host_record_cardinality_invalid:{layer}:{len(matches)}")
        elif not matches:
            raise V4CandidateContractError(f"passed_layer_has_no_record:{layer}")
    return typed_records, record_ids


def _expected_pointer_path(root: Path, layer: str) -> str:
    if layer == "source_oracle":
        authority = read_json_object(root / AUTHORITY_PATH, label="staged_candidate_authority")
        active = authority.get("activeSemantics")
        if not isinstance(active, dict) or not isinstance(active.get("oracleManifestPath"), str):
            raise V4CandidateContractError("staged_oracle_manifest_path_missing")
        return cast(str, active["oracleManifestPath"])
    if layer in {"machine_wire", "packaged"}:
        return {"machine_wire": WIRE_SCHEMA_PATH, "packaged": PACKAGE_MANIFEST_PATH}[layer]
    return evidence_contract.manifest_path(layer)


def verify_index_files(
    root: Path, records: Sequence[Mapping[str, object]], *, stable_tree: Mapping[str, object] | None = None
) -> None:
    if stable_tree is not None and capture_tree_stable(root) != stable_tree:
        raise V4CandidateContractError("evidence_tree_changed_before_verification")
    tree = stable_tree
    records_by_layer = {layer: [item for item in records if item.get("layer") == layer] for layer in REQUIRED_LAYERS}
    for record in records:
        raw_path = record.get("relativePath")
        layer = str(record.get("layer"))
        if (
            not isinstance(raw_path, str)
            or not raw_path.startswith(f"evidence/records/{layer}/")
            or not raw_path.endswith(".json")
            or Path(raw_path).is_absolute()
            or ".." in Path(raw_path).parts
        ):
            raise V4CandidateContractError(f"evidence_path_invalid:{raw_path}")
        actual = (
            identity_from_tree(root / raw_path, relative_to=root, stable_tree=tree)
            if tree is not None
            else file_identity(root / raw_path, relative_to=root)
        )
        if any(actual[key] != record.get(key) for key in ("relativePath", "bytes", "sha256")):
            raise V4CandidateContractError(f"evidence_record_file_mismatch:{raw_path}")
        evidence = read_json_object(root / raw_path, label=f"evidence_record:{raw_path}")
        if not evidence_contract.record_envelope_matches(evidence, case_id=record.get("caseId"), layer=layer):
            raise V4CandidateContractError(f"evidence_record_envelope_mismatch:{raw_path}")
        pointers = record.get("identityPointers")
        if not isinstance(pointers, list) or not pointers:
            raise V4CandidateContractError(f"evidence_identity_pointer_missing:{raw_path}")
        for pointer in pointers:
            if not isinstance(pointer, dict):
                raise V4CandidateContractError(f"evidence_identity_pointer_invalid:{raw_path}")
            if pointer.get("role") != POINTER_ROLE_BY_LAYER.get(layer):
                raise V4CandidateContractError(f"evidence_identity_pointer_role_mismatch:{raw_path}")
            if pointer.get("relativePath") != _expected_pointer_path(root, layer):
                raise V4CandidateContractError(f"evidence_identity_pointer_path_mismatch:{raw_path}")
            pointed_path = root / str(pointer.get("relativePath"))
            pointed = (
                identity_from_tree(pointed_path, relative_to=root, stable_tree=tree)
                if tree is not None
                else file_identity(pointed_path, relative_to=root)
            )
            if any(pointed[key] != pointer.get(key) for key in ("relativePath", "bytes", "sha256")):
                raise V4CandidateContractError(f"evidence_identity_pointer_mismatch:{raw_path}")
    authority, _ = load_authority(root)
    source_identity = oracle_manifest_identity(root, authority)
    source_manifest = read_json_object(root / str(source_identity["relativePath"]), label="source_manifest")
    if source_manifest.get("evidence_layer") != "source_oracle":
        raise V4CandidateContractError("source_manifest_layer_mismatch")
    wire_identity = file_identity(root / WIRE_SCHEMA_PATH, relative_to=root)
    wire_schema = read_json_object(root / WIRE_SCHEMA_PATH, label="wire_schema")
    if wire_schema.get("$id") != "urn:docwen:schema:machine-diagnostic-evidence:v1":
        raise V4CandidateContractError("wire_schema_identity_mismatch")
    package_path = root / PACKAGE_MANIFEST_PATH
    package_roots = {f"DocWen_v{PRODUCT_VERSION}_win-x64", f"DocWenCLI_v{PRODUCT_VERSION}_win-x64"}
    allowed = tuple(sorted(item.name for item in root.iterdir() if item.name not in package_roots and item.is_dir()))
    package_manifest = read_json_object(package_path, label="package_manifest")
    candidate_evidence.verify_package_manifest(root, package_path, allowed_root_entries=allowed)
    package_identity = file_identity(package_path, relative_to=root)
    source_pointer = evidence_contract.pointer_identity("source_manifest", source_identity)
    wire_pointer = evidence_contract.pointer_identity("wire_schema", wire_identity)
    package_pointer = evidence_contract.pointer_identity("package_manifest", package_identity)
    try:
        evidence_contract.verify_observations(
            root,
            records,
            source_manifest_identity=source_identity,
            source_manifest=source_manifest,
            package_identity=package_identity,
            package_manifest=package_manifest,
        )
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4CandidateContractError(str(exc)) from exc
    for layer in ("source_wire_comparison", "roundtrip", "headless_ooxml", *REQUIRED_LAYERS[6:]):
        layer_records = records_by_layer[layer]
        if layer_records:
            value = read_json_object(root / evidence_contract.manifest_path(layer), label=f"{layer}_manifest")
            expected = evidence_contract.manifest_expected(
                layer=layer,
                records=layer_records,
                source_pointer=source_pointer,
                wire_pointer=wire_pointer,
                package_pointer=package_pointer,
            )
            if value != expected:
                raise V4CandidateContractError(f"evidence_manifest_not_closed_or_mismatched:{layer}")
        elif (root / evidence_contract.manifest_path(layer)).exists():
            raise V4CandidateContractError(f"not_run_layer_manifest_present:{layer}")
    if stable_tree is not None and capture_tree_stable(root) != stable_tree:
        raise V4CandidateContractError("evidence_tree_changed_during_verification")


def machine_consumer_contract() -> dict[str, object]:
    common_properties = {
        "recognize_text": {"type": "boolean", "default": False},
        "preserve_resources": {"type": "boolean", "default": True},
        "ocr_language": {
            "type": "string",
            "enum": ["auto", "chinese", "chinese_cht", "english", "japanese", "korean", "latin", "cyrillic"],
            "default": "auto",
        },
    }
    docx_properties = {
        **common_properties,
        "image_mode": {"type": "string", "enum": ["file", "base64", "embed", "omit"], "default": "file"},
        "ocr_placement": {"type": "string", "enum": ["image_md", "main_md"], "default": "main_md"},
        "image_link_style": {
            "type": "string",
            "enum": ["wiki_embed", "wiki_link", "markdown_embed", "markdown_link"],
            "default": "wiki_embed",
        },
        "table_merge_strategy": {"type": "string", "enum": ["fill", "empty", "marker"], "default": "fill"},
        "remove_numbering": {"type": "boolean", "default": True},
        "add_numbering": {"type": "boolean", "default": False},
        "numbering_scheme": {
            "type": "string",
            "default": "gongwen_standard",
            "x-docwen-resource-kind": "numbering-schemes",
        },
    }
    fixed_layout_properties = {
        **common_properties,
        "image_mode": {"type": "string", "enum": ["file"], "default": "file"},
        "render_dpi": {"type": "integer", "minimum": 72, "maximum": 600, "default": 200},
    }

    def closed_schema(properties: Mapping[str, object]) -> dict[str, object]:
        return {
            "$schema": "https://json-schema.org/draft/2020-12/schema",
            "type": "object",
            "properties": dict(properties),
            "required": [],
            "additionalProperties": False,
        }

    return {
        "optionsSchemas": {
            "docx": {
                "profile": "document_with_resources",
                "schema": closed_schema(docx_properties),
            },
            "pdfOfdXps": {
                "profile": "physical_page_ocr",
                "schema": closed_schema(fixed_layout_properties),
            },
            "tiff": {
                "profile": "physical_page_ocr",
                "schema": closed_schema(common_properties),
            },
        },
        "bundleMatrices": {
            "physicalPagePK": [
                {
                    "recognizeText": False,
                    "preserveResources": False,
                    "artifacts": "1 document",
                    "entries": 1,
                    "relations": 0,
                },
                {
                    "recognizeText": False,
                    "preserveResources": True,
                    "artifacts": "1 document + K resources",
                    "entries": 1,
                    "relations": "K resource_of",
                },
                {
                    "recognizeText": True,
                    "preserveResources": False,
                    "artifacts": "1 document + P fragments",
                    "entries": 1,
                    "relations": "P fragment_of/ocr_page",
                },
                {
                    "recognizeText": True,
                    "preserveResources": True,
                    "artifacts": "1 document + P fragments + K resources",
                    "entries": 1,
                    "relations": "P fragment_of/ocr_page + K resource_of",
                },
            ],
            "docx": {
                "entries": "exactly one primary preferred ordinal-0 document",
                "resources": "exactly K image resources and K resource_of/image relations iff preserve_resources",
                "mainMd": "recognition text stays in primary; zero OCR fragments",
                "imageMd": "exactly R non-empty OCR fragments and R fragment_of/ocr_text relations iff recognize_text",
                "forbidden": ["ocr_page", "page_fragment", "page_resource", "consumer_node_instruction"],
            },
            "imageSidecar": "resource only; never an entry, document, fragment, or implicit consumer Node",
        },
        "syntaxObservability": {
            "ordinaryWikiLink": ["[[Page#^id]]", "![[Page#^id]]"],
            "stableCrossReference": ["@[[#^id]]", "@[[Page#^id]]"],
            "softHeadingCrossReference": ["@[[#Heading]]", "@[[Page#Parent#Heading]]"],
            "citation": ["@citation-key", "[@first; @second]"],
            "typedPrefixDoesNotEncodeKind": True,
        },
        "diagnostics": {
            "schema": "docwen.machine.diagnostic_evidence.v1",
            "wire": "exact raw task terminal JSON object evidence",
            "comparison": "separate source_wire_comparison layer; source oracle is never wire evidence",
            "range": {"coordinateSystem": "unicode_code_point", "offsetBase": 0, "end": "exclusive"},
        },
    }


def validate_json_schema(instance: Any, schema_path: Path) -> None:
    try:
        from jsonschema import Draft202012Validator
    except ImportError as exc:  # pragma: no cover - packaging gate owns dependency presence
        raise V4CandidateContractError("jsonschema_dependency_missing") from exc
    schema = read_json_object(schema_path, label="json_schema")
    Draft202012Validator.check_schema(schema)
    errors = sorted(Draft202012Validator(schema).iter_errors(instance), key=lambda error: list(error.absolute_path))
    if errors:
        first = errors[0]
        path = "/".join(str(item) for item in first.absolute_path)
        raise V4CandidateContractError(f"json_schema_validation_failed:{schema_path.name}:{path}:{first.message}")


def require_hex64(value: str, *, label: str) -> str:
    if not _HEX64.fullmatch(value):
        raise V4CandidateContractError(f"{label}_sha256_invalid")
    return value


def capture_source_checkpoint(
    *,
    docwen_repo: Path,
) -> dict[str, object]:
    authority, authority_identity = load_authority(docwen_repo)
    raw_docwen_implementation = authority.get("docwenImplementationBaseline")
    if (
        not isinstance(raw_docwen_implementation, dict)
        or raw_docwen_implementation.get("status") != "frozen"
        or not isinstance(raw_docwen_implementation.get("commit"), str)
        or not isinstance(raw_docwen_implementation.get("tree"), str)
    ):
        raise V4CandidateContractError("docwen_implementation_baseline_pending")
    docwen_final = clean_git_identity(docwen_repo, label="docwen")
    docwen_spec = cast(Mapping[str, object], authority["docwenSpecBaseline"])
    verify_baseline_ancestry(docwen_repo, docwen_spec, docwen_final, label="docwen_spec")
    docwen_implementation = {
        "commit": raw_docwen_implementation["commit"],
        "tree": raw_docwen_implementation["tree"],
    }
    verify_baseline_ancestry(docwen_repo, docwen_implementation, docwen_final, label="docwen_implementation")
    docwen_authority = {
        "specBaseline": dict(docwen_spec),
        "implementationBaseline": docwen_implementation,
        "final": docwen_final,
    }
    ignored_executable = ignored_executable_inputs(docwen_repo)
    if ignored_executable:
        raise V4CandidateContractError(f"docwen_ignored_executable_build_input_present:{ignored_executable}")
    ignored_source = ignored_source_inputs(docwen_repo)
    if ignored_source:
        raise V4CandidateContractError(f"docwen_ignored_source_input_present:{ignored_source}")
    return {
        "schema": "docwen.v4_candidate_source_checkpoint.v1",
        "authorityBaseline": authority_identity,
        "activeSemantics": dict(cast(Mapping[str, object], authority["activeSemantics"])),
        "finalSpecIdentity": final_spec_identity(docwen_repo, authority),
        "oracleManifest": oracle_manifest_identity(docwen_repo, authority),
        "docwen": docwen_authority,
        "exclusions": authority["exclusions"],
        "ignoredExecutableInputs": [],
        "ignoredSourceInputs": [],
    }


def verify_source_checkpoint(
    checkpoint: Mapping[str, object],
    *,
    docwen_repo: Path,
) -> None:
    if checkpoint.get("schema") != "docwen.v4_candidate_source_checkpoint.v1":
        raise V4CandidateContractError("source_checkpoint_schema_mismatch")
    current = capture_source_checkpoint(docwen_repo=docwen_repo)
    if dict(checkpoint) != current:
        raise V4CandidateContractError("source_checkpoint_payload_mismatch")


def consumer_eligible(layer_status: Mapping[str, object]) -> bool:
    return all(layer_status.get(layer) == "passed" for layer in REQUIRED_LAYERS[6:])
