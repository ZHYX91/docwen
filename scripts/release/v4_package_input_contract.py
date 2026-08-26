from __future__ import annotations

import hashlib
import json
import re
import subprocess
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, cast

from scripts.release import v4_candidate_contract as candidate_contract
from scripts.release import v4_evidence_contract as evidence_contract
from scripts.release import v4_evidence_io as evidence_io
from scripts.release import v4_package_input_machine as machine_proof
from scripts.release import v4_package_input_ooxml as ooxml_proof

HARNESS_ID = "docwen-v4-exact-two-numbering"
HARNESS_VERSION = 1
HARNESS_SCHEMA = "docwen.v4_package_input_harness.v1"
HARNESS_MANIFEST_RELATIVE = (
    "packages/plugins/markdown/tests/fixtures/resolved_v4/docwen-v4-package-input-harness.v1.json"
)
HARNESS_CASE_PREFIX = "docwen-x2-v1"
TRANSCRIPT_SCHEMA = machine_proof.TRANSCRIPT_SCHEMA
SOURCE_ORACLE_ROOT = Path("contracts/oracles/docwen.markdown_semantics.v3")
VALIDATION_CASE_ID = "invalid-id-dot"
EXACT_CAPABILITY_ID = "convert.markdown.to_docx"
NEUTRAL_MEDIA_TYPE = "application/vnd.docwen.resolved-document+json"
PLAN_MEDIA_TYPE = "application/vnd.docwen.numbering-export-plan+json"
DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PACKAGE_NAMES = (
    f"DocWen_v{candidate_contract.PRODUCT_VERSION}_win-x64",
    f"DocWenCLI_v{candidate_contract.PRODUCT_VERSION}_win-x64",
)

# The DocWen-owned release fixture is closed and provider-neutral.
REQUIRED_HARNESS_CASE_IDS = ("rich-semantics-composite",)

_HEX64 = re.compile(r"^[0-9a-f]{64}$")
_HEX40 = re.compile(r"^[0-9a-f]{40}$")
_CASE_ID = re.compile(r"^[a-z0-9][a-z0-9.-]{0,127}$")
_LEGACY_HARNESS_MARKERS = (
    "attempt04",
    "c6-v",
    "c6v",
    "docwen-v3-dirty-audit-head",
    "docwen.local_candidate.v3",
    "legacy-v3-harness",
    "machine.typed-markdown-png",
    "source+bibliography",
    "source-bibliography",
)
_EXPECTED_OOXML_KEYS = {
    "abstractNumCount",
    "numCount",
    "bookmarkCount",
    "seqFieldCount",
    "styleRefFieldCount",
    "refFieldCount",
    "citationFieldCount",
}
_RICH_COVERAGE = {
    "citation",
    "embedded-raster",
    "five-kind-enabled",
    "ordinary-wikilink-preserved",
    "semantic-bibliography",
    "semantic-cross-reference",
}


class V4PackageInputError(RuntimeError):
    """A package input cannot be proved against the frozen v4 contract."""


@dataclass(frozen=True)
class BuildOutput:
    gui: Path
    cli: Path
    metadata: dict[str, object]


@dataclass(frozen=True)
class HarnessCase:
    case_id: str
    source_oracle_case_id: str
    coverage: tuple[str, ...]
    source_bytes: bytes
    source_identity: dict[str, object]
    neutral_path: Path
    neutral_bytes: bytes
    neutral_identity: dict[str, object]
    neutral_envelope: dict[str, Any]
    plan_path: Path
    plan_bytes: bytes
    plan_identity: dict[str, object]
    plan_envelope: dict[str, Any]
    expected_ooxml: dict[str, int | bool]
    expected_roundtrip: bytes
    roundtrip_path: Path
    roundtrip_identity: dict[str, object]


@dataclass(frozen=True)
class HarnessInput:
    cases: tuple[HarnessCase, ...]
    manifest_path: Path
    manifest_identity: dict[str, object]


@dataclass(frozen=True)
class HarnessCaseOutput:
    case_id: str
    docx: bytes
    roundtrip: bytes
    inspection: dict[str, object]


@dataclass(frozen=True)
class HarnessOutput:
    validation_terminal: dict[str, object]
    stdout: bytes
    stderr: bytes
    transcript: bytes
    cases: tuple[HarnessCaseOutput, ...]
    request_digest: str


def json_bytes(value: object) -> bytes:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True).encode("utf-8")


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def transcript_event(
    *,
    direction: str,
    operation: str,
    case_id: str | None,
    raw_frame: bytes,
    message: Mapping[str, object],
) -> dict[str, object]:
    try:
        return machine_proof.transcript_event(
            direction=direction,
            operation=operation,
            case_id=case_id,
            raw_frame=raw_frame,
            message=message,
        )
    except machine_proof.V4MachineProofError as exc:
        raise V4PackageInputError(str(exc)) from exc


def transcript_request_digest(events: Sequence[Mapping[str, object]]) -> str:
    try:
        return machine_proof.transcript_request_digest(events)
    except machine_proof.V4MachineProofError as exc:
        raise V4PackageInputError(str(exc)) from exc


def build_session_transcript(
    *,
    harness: HarnessInput,
    events: Sequence[Mapping[str, object]],
    request_digest: str,
    stdout: bytes,
    stderr: bytes,
    exit_code: int,
) -> bytes:
    try:
        return machine_proof.build_session_transcript(
            harness=harness,
            events=events,
            request_digest=request_digest,
            stdout=stdout,
            stderr=stderr,
            exit_code=exit_code,
        )
    except machine_proof.V4MachineProofError as exc:
        raise V4PackageInputError(str(exc)) from exc


def reject_legacy_harness(value: object) -> None:
    serialized = json.dumps(value, ensure_ascii=False, sort_keys=True).casefold()
    marker = next((item for item in _LEGACY_HARNESS_MARKERS if item in serialized), None)
    if marker is not None:
        raise V4PackageInputError(f"legacy_harness_marker_rejected:{marker}")


def _safe_relative(value: object, *, label: str) -> str:
    try:
        return evidence_contract.validate_relative_path(value, label=label)
    except evidence_contract.V4EvidenceContractError as exc:
        raise V4PackageInputError(str(exc)) from exc


def _identity(path: Path, *, root: Path) -> dict[str, object]:
    try:
        return evidence_io.file_identity(path, relative_to=root)
    except evidence_io.V4EvidenceContractError as exc:
        raise V4PackageInputError(str(exc)) from exc


def _read_json(path: Path, *, root: Path, label: str) -> tuple[dict[str, Any], dict[str, object]]:
    try:
        return evidence_io.read_json_object(path, relative_to=root, label=label)
    except evidence_io.V4EvidenceContractError as exc:
        raise V4PackageInputError(str(exc)) from exc


def _stable_bytes(path: Path, *, root: Path, label: str) -> tuple[bytes, dict[str, object]]:
    try:
        safe, raw, size, digest = evidence_io._read_stable(path, label=label, collect=True)
        assert raw is not None
        relative = safe.relative_to(evidence_io._root(root)).as_posix()
        evidence_io.validate_relative_path(relative, label=label)
    except (ValueError, evidence_io.V4EvidenceContractError) as exc:
        raise V4PackageInputError(str(exc)) from exc
    return raw, {"relativePath": relative, "bytes": size, "sha256": digest}


def _json_from_bytes(raw: bytes, *, label: str) -> dict[str, Any]:
    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        value: dict[str, Any] = {}
        for key, item in pairs:
            if key in value:
                raise V4PackageInputError(f"{label}_duplicate_json_key:{key}")
            value[key] = item
        return value

    try:
        value = json.loads(raw.decode("utf-8", errors="strict"), object_pairs_hook=reject_duplicates)
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise V4PackageInputError(f"{label}_invalid_json") from exc
    if not isinstance(value, dict):
        raise V4PackageInputError(f"{label}_not_object")
    return value


def _bound_pointer_bytes(
    root: Path,
    value: object,
    *,
    label: str,
) -> tuple[Path, bytes, dict[str, object]]:
    if not isinstance(value, dict) or set(value) != {"relativePath", "bytes", "sha256"}:
        raise V4PackageInputError(f"{label}_pointer_not_closed")
    relative = _safe_relative(value.get("relativePath"), label=label)
    path = root / relative
    raw, actual = _stable_bytes(path, root=root, label=label)
    if actual != value:
        raise V4PackageInputError(f"{label}_pointer_identity_mismatch")
    return path, raw, actual


def _bound_pointer_json(
    root: Path,
    value: object,
    *,
    label: str,
) -> tuple[Path, bytes, dict[str, Any], dict[str, object]]:
    path, raw, identity = _bound_pointer_bytes(root, value, label=label)
    return path, raw, _json_from_bytes(raw, label=label), identity


def run_bytes(
    command: Sequence[str],
    *,
    cwd: Path,
    env: Mapping[str, str] | None = None,
    timeout: int = 900,
) -> subprocess.CompletedProcess[bytes]:
    return subprocess.run(
        list(command),
        cwd=cwd,
        env=dict(env) if env is not None else None,
        stdin=subprocess.DEVNULL,
        capture_output=True,
        timeout=timeout,
        check=False,
    )


def require_command(
    command: Sequence[str],
    *,
    cwd: Path,
    env: Mapping[str, str] | None = None,
    label: str,
    timeout: int = 900,
) -> subprocess.CompletedProcess[bytes]:
    completed = run_bytes(command, cwd=cwd, env=env, timeout=timeout)
    if completed.returncode != 0:
        tail = completed.stderr.decode("utf-8", errors="replace")[-4000:]
        raise V4PackageInputError(f"{label}_failed:{completed.returncode}:{tail}")
    return completed


def git(repo: Path, *arguments: str, label: str = "git") -> str:
    completed = require_command(["git", *arguments], cwd=repo, label=label)
    try:
        return completed.stdout.decode("utf-8", errors="strict").strip()
    except UnicodeDecodeError as exc:
        raise V4PackageInputError(f"{label}_output_not_utf8") from exc


def load_checkpoint(
    path: Path,
    expected_sha256: str,
    docwen_repo: Path,
) -> tuple[dict[str, Any], dict[str, object]]:
    candidate_contract.require_hex64(expected_sha256, label="source_checkpoint")
    try:
        value, identity = candidate_contract.read_json_object_with_identity(
            path,
            relative_to=path.parent,
            label="source_checkpoint",
            expected_sha256=expected_sha256,
        )
        candidate_contract.verify_source_checkpoint(value, docwen_repo=docwen_repo)
    except candidate_contract.V4CandidateContractError as exc:
        raise V4PackageInputError(str(exc)) from exc
    return value, identity


def clone_exact(source: Path, destination: Path, commit: str, tree: str, label: str) -> None:
    if not _HEX40.fullmatch(commit) or not _HEX40.fullmatch(tree):
        raise V4PackageInputError(f"{label}_clone_identity_invalid")
    destination.parent.mkdir(parents=True, exist_ok=True)
    require_command(
        ["git", "clone", "--local", "--no-hardlinks", "--no-checkout", str(source), str(destination)],
        cwd=destination.parent,
        label=f"{label}_clone",
    )
    require_command(["git", "checkout", "--detach", commit], cwd=destination, label=f"{label}_checkout")
    actual_commit = git(destination, "rev-parse", "HEAD", label=f"{label}_head")
    actual_tree = git(destination, "rev-parse", "HEAD^{tree}", label=f"{label}_tree")
    status = git(destination, "status", "--porcelain=v2", label=f"{label}_status")
    if (actual_commit, actual_tree, status) != (commit, tree, ""):
        raise V4PackageInputError(f"{label}_clone_not_exact_or_clean")


def _expected_ooxml(value: object, *, case_id: str) -> dict[str, int | bool]:
    if (
        not isinstance(value, dict)
        or set(value) != _EXPECTED_OOXML_KEYS
        or any(
            not isinstance(value.get(key), int) or isinstance(value.get(key), bool) or cast(int, value[key]) < 0
            for key in _EXPECTED_OOXML_KEYS
        )
    ):
        raise V4PackageInputError(f"exact_two_expected_ooxml_invalid:{case_id}")
    return cast(dict[str, int | bool], value)


def _validate_case_semantics(case_id: str, expected: Mapping[str, int | bool]) -> None:
    if case_id.endswith("-off") and (
        expected["abstractNumCount"] != 0
        or expected["numCount"] != 0
        or expected["seqFieldCount"] != 0
        or expected["refFieldCount"] != 0
    ):
        raise V4PackageInputError(f"exact_two_disabled_case_not_zero:{case_id}")
    if case_id in {"heading-level-template-empty", "ordinary-wikilink-navigation"} and (
        expected["abstractNumCount"] != 0
        or expected["numCount"] != 0
        or expected["seqFieldCount"] != 0
        or expected["refFieldCount"] != 0
    ):
        raise V4PackageInputError(f"exact_two_nonmaterialized_case_not_zero:{case_id}")
    if case_id == "numbering-heading-on" and (
        expected["abstractNumCount"] < 1 or expected["numCount"] < 1 or expected["refFieldCount"] < 1
    ):
        raise V4PackageInputError("exact_two_heading_on_expectation_invalid")
    if (
        case_id.startswith("numbering-")
        and case_id.endswith("-on")
        and case_id != "numbering-heading-on"
        and (expected["seqFieldCount"] < 1 or expected["refFieldCount"] < 1)
    ):
        raise V4PackageInputError(f"exact_two_caption_on_expectation_invalid:{case_id}")
    if case_id == "heading-authored-number-preserved" and (
        expected["abstractNumCount"] < 1 or expected["numCount"] < 1
    ):
        raise V4PackageInputError("exact_two_authored_heading_expectation_invalid")


def _load_case(raw: object, *, docwen: Path) -> HarnessCase:
    if not isinstance(raw, dict) or set(raw) != {
        "caseId",
        "sourceOracleCaseId",
        "coverage",
        "source",
        "neutralDocument",
        "numberingExportPlan",
        "expectedOoxml",
        "expectedRoundtrip",
    }:
        raise V4PackageInputError("exact_two_harness_case_not_closed")
    case_id = raw.get("caseId")
    oracle_case_id = raw.get("sourceOracleCaseId")
    if not isinstance(case_id, str) or _CASE_ID.fullmatch(case_id) is None:
        raise V4PackageInputError("exact_two_case_id_invalid")
    if not isinstance(oracle_case_id, str) or _CASE_ID.fullmatch(oracle_case_id) is None:
        raise V4PackageInputError(f"exact_two_source_case_id_invalid:{case_id}")
    coverage = raw.get("coverage")
    if (
        not isinstance(coverage, list)
        or not coverage
        or any(not isinstance(item, str) or re.fullmatch(r"[a-z0-9][a-z0-9-]{0,63}", item) is None for item in coverage)
        or coverage != sorted(coverage)
        or len(set(cast(list[str], coverage))) != len(coverage)
    ):
        raise V4PackageInputError(f"exact_two_coverage_invalid:{case_id}")
    if case_id == "rich-semantics-composite" and not _RICH_COVERAGE.issubset(cast(set[str], set(coverage))):
        raise V4PackageInputError("exact_two_rich_coverage_incomplete")
    _source_path, source_bytes, source_identity = _bound_pointer_bytes(
        docwen,
        raw.get("source"),
        label=f"harness_source:{case_id}",
    )
    neutral_path, neutral_bytes, neutral, neutral_identity = _bound_pointer_json(
        docwen,
        raw.get("neutralDocument"),
        label=f"harness_neutral_document:{case_id}",
    )
    plan_path, plan_bytes, plan, plan_identity = _bound_pointer_json(
        docwen,
        raw.get("numberingExportPlan"),
        label=f"harness_numbering_export_plan:{case_id}",
    )
    roundtrip_path, expected_roundtrip, roundtrip_identity = _bound_pointer_bytes(
        docwen,
        raw.get("expectedRoundtrip"),
        label=f"harness_expected_roundtrip:{case_id}",
    )
    if (
        source_identity["bytes"] != roundtrip_identity["bytes"]
        or source_identity["sha256"] != roundtrip_identity["sha256"]
        or source_bytes != expected_roundtrip
    ):
        raise V4PackageInputError(f"exact_two_expected_roundtrip_not_source_exact:{case_id}")
    try:
        authored_source = source_bytes.decode("utf-8", errors="strict")
    except UnicodeDecodeError as exc:
        raise V4PackageInputError(f"exact_two_source_not_utf8:{case_id}") from exc
    reject_legacy_harness(authored_source)
    reject_legacy_harness({"neutralDocument": neutral, "numberingExportPlan": plan})
    if (
        set(neutral) != {"$schema", "schema", "input_id", "source_sha256", "plan_sha256", "document"}
        or neutral.get("$schema") != "urn:docwen:schema:resolved-document:v1"
        or neutral.get("schema") != "docwen.resolved_document.v1"
        or set(plan) != {"$schema", "schema", "input_id", "source_sha256", "plan_sha256", "plan"}
        or plan.get("$schema") != "urn:docwen:schema:numbering-export-plan:v1"
        or plan.get("schema") != "docwen.numbering_export_plan.v1"
    ):
        raise V4PackageInputError(f"exact_two_envelope_identity_invalid:{case_id}")
    shared = (neutral.get("input_id"), neutral.get("source_sha256"), neutral.get("plan_sha256"))
    if shared != (plan.get("input_id"), plan.get("source_sha256"), plan.get("plan_sha256")):
        raise V4PackageInputError(f"exact_two_envelope_pointer_mismatch:{case_id}")
    document = neutral.get("document")
    if not isinstance(document, dict) or document.get("authored_markdown") != authored_source:
        raise V4PackageInputError(f"exact_two_authored_source_mismatch:{case_id}")
    canonical_plan_sha = sha256_bytes(json_bytes(plan.get("plan")))
    if shared[1] != sha256_bytes(source_bytes) or shared[2] != canonical_plan_sha:
        raise V4PackageInputError(f"exact_two_source_or_plan_hash_invalid:{case_id}")
    oracle_case_path = docwen / SOURCE_ORACLE_ROOT / "corpus" / f"{oracle_case_id}.case.json"
    oracle_case, _ = _read_json(oracle_case_path, root=docwen, label=f"exact_two_source_oracle_case:{case_id}")
    if oracle_case.get("case_id") != oracle_case_id or oracle_case.get("source") != authored_source:
        raise V4PackageInputError(f"exact_two_source_not_bound_to_docwen_oracle:{case_id}")
    expected = _expected_ooxml(raw.get("expectedOoxml"), case_id=case_id)
    _validate_case_semantics(case_id, expected)
    return HarnessCase(
        case_id=case_id,
        source_oracle_case_id=oracle_case_id,
        coverage=tuple(cast(list[str], coverage)),
        source_bytes=source_bytes,
        source_identity=source_identity,
        neutral_path=neutral_path,
        neutral_bytes=neutral_bytes,
        neutral_identity=neutral_identity,
        neutral_envelope=neutral,
        plan_path=plan_path,
        plan_bytes=plan_bytes,
        plan_identity=plan_identity,
        plan_envelope=plan,
        expected_ooxml=expected,
        expected_roundtrip=expected_roundtrip,
        roundtrip_path=roundtrip_path,
        roundtrip_identity=roundtrip_identity,
    )


def load_harness(docwen: Path) -> HarnessInput:
    manifest, manifest_identity = _read_json(
        docwen / HARNESS_MANIFEST_RELATIVE,
        root=docwen,
        label="exact_two_harness_manifest",
    )
    if set(manifest) != {"schema", "harness", "requiredCaseIds", "pendingCaseIds", "cases"}:
        raise V4PackageInputError("exact_two_harness_manifest_not_closed")
    if manifest.get("schema") != HARNESS_SCHEMA or manifest.get("harness") != {
        "id": HARNESS_ID,
        "version": HARNESS_VERSION,
    }:
        raise V4PackageInputError("exact_two_harness_identity_mismatch")
    reject_legacy_harness(manifest)
    required = manifest.get("requiredCaseIds")
    pending = manifest.get("pendingCaseIds")
    if required != list(REQUIRED_HARNESS_CASE_IDS):
        raise V4PackageInputError("exact_two_required_case_ids_invalid")
    if (
        not isinstance(pending, list)
        or any(not isinstance(item, str) or item not in REQUIRED_HARNESS_CASE_IDS for item in pending)
        or pending != sorted(pending)
        or len(set(cast(list[str], pending))) != len(pending)
    ):
        raise V4PackageInputError("exact_two_pending_case_ids_invalid")
    raw_cases = manifest.get("cases")
    if not isinstance(raw_cases, list):
        raise V4PackageInputError("exact_two_harness_cases_invalid")
    case_ids = [item.get("caseId") if isinstance(item, dict) else None for item in raw_cases]
    expected_present = [case_id for case_id in REQUIRED_HARNESS_CASE_IDS if case_id not in pending]
    if case_ids != expected_present or len(set(case_ids)) != len(case_ids):
        raise V4PackageInputError("exact_two_harness_case_set_or_order_invalid")
    if pending:
        raise V4PackageInputError(f"exact_two_harness_cases_pending:{','.join(cast(list[str], pending))}")
    cases = tuple(_load_case(item, docwen=docwen) for item in raw_cases)
    return HarnessInput(
        cases=cases,
        manifest_path=docwen / HARNESS_MANIFEST_RELATIVE,
        manifest_identity=manifest_identity,
    )


def revalidate_harness(harness: HarnessInput, *, docwen: Path) -> None:
    if _identity(harness.manifest_path, root=docwen) != harness.manifest_identity:
        raise V4PackageInputError("exact_two_harness_manifest_changed")
    for case in harness.cases:
        for label, path, expected_identity, expected_bytes in (
            ("source", docwen / str(case.source_identity["relativePath"]), case.source_identity, case.source_bytes),
            ("neutral", case.neutral_path, case.neutral_identity, case.neutral_bytes),
            ("plan", case.plan_path, case.plan_identity, case.plan_bytes),
            ("roundtrip", case.roundtrip_path, case.roundtrip_identity, case.expected_roundtrip),
        ):
            raw, identity = _stable_bytes(path, root=docwen, label=f"harness_recheck:{case.case_id}:{label}")
            if identity != expected_identity:
                raise V4PackageInputError(f"exact_two_harness_pointer_changed:{case.case_id}:{label}")
            if raw != expected_bytes:
                raise V4PackageInputError(f"exact_two_harness_bytes_changed:{case.case_id}:{label}")
            if label in {"neutral", "plan"}:
                expected_value = case.neutral_envelope if label == "neutral" else case.plan_envelope
                if _json_from_bytes(raw, label=f"harness_recheck:{case.case_id}:{label}") != expected_value:
                    raise V4PackageInputError(f"exact_two_harness_json_changed:{case.case_id}:{label}")


def verify_clone_identity(repo: Path, commit: str, tree: str, label: str) -> None:
    actual_commit = git(repo, "rev-parse", "HEAD", label=f"{label}_recheck_head")
    actual_tree = git(repo, "rev-parse", "HEAD^{tree}", label=f"{label}_recheck_tree")
    status = git(repo, "status", "--porcelain=v2", label=f"{label}_recheck_status")
    if (actual_commit, actual_tree, status) != (commit, tree, ""):
        raise V4PackageInputError(f"{label}_clone_identity_drifted")


def exact_two_inputs(neutral: Path, plan: Path) -> list[dict[str, object]]:
    try:
        return machine_proof.exact_two_inputs(neutral, plan)
    except machine_proof.V4MachineProofError as exc:
        raise V4PackageInputError(str(exc)) from exc


def validate_exact_two_request(inputs: object, options: object) -> None:
    try:
        machine_proof.validate_exact_two_request(inputs, options)
    except machine_proof.V4MachineProofError as exc:
        raise V4PackageInputError(str(exc)) from exc


def validate_session_transcript(
    raw: bytes,
    *,
    harness: HarnessInput,
    outputs: Sequence[HarnessCaseOutput] | None = None,
) -> dict[str, Any]:
    try:
        return machine_proof.validate_session_transcript(raw, harness=harness, outputs=outputs)
    except machine_proof.V4MachineProofError as exc:
        raise V4PackageInputError(str(exc)) from exc


def inspect_docx(
    payload: bytes,
    expected: Mapping[str, int | bool],
    neutral_envelope: Mapping[str, object],
    plan_envelope: Mapping[str, object],
) -> dict[str, object]:
    try:
        return ooxml_proof.inspect_resolved_docx(payload, expected, neutral_envelope, plan_envelope)
    except ooxml_proof.V4OoxmlProofError as exc:
        raise V4PackageInputError(f"headless_ooxml_inspection_failed:{exc}") from exc
