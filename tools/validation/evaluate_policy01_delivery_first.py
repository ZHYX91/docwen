"""Evaluate the frozen POLICY-01=B delivery-first oracle.

The evaluator reuses the VIS-116 Apache POI sources, creates deterministic
content-tampered siblings whose signature graph remains byte-identical, runs
the 15 current conversion slots, exercises CLI inspect JSON/text, and projects
the same results through GUI admission/history consumers.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any

VALIDATION_CODE = "OOXML_SIGNATURE_VALIDATION_UNAVAILABLE"
DERIVED_CODE = "OOXML_SIGNATURE_DERIVED_OUTPUT_UNSIGNED"
MUTATIONS = {
    "docx": "word/document.xml",
    "xlsx": "xl/sharedStrings.xml",
    "pptx": "ppt/slides/slide1.xml",
}
TESTED_REPO_PATHS = (
    "packages/application/src/docwen_application/controller.py",
    "packages/apps/cli/src/docwen_cli/commands/inspect.py",
    "packages/apps/gui/src/docwen_gui/view_models/batch_list_vm.py",
    "packages/apps/gui/src/docwen_gui/view_models/main_window_vm.py",
    "packages/core/src/docwen_core/detection/__init__.py",
    "packages/core/src/docwen_core/detection/_validation.py",
    "packages/core/src/docwen_core/detection/ooxml_signature.py",
    "packages/runtime/src/docwen_runtime/engine/task_manager.py",
    "tools/validation/evaluate_policy01_delivery_first.py",
)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2),
        encoding="utf-8",
        newline="\n",
    )


def _make_invalid_sibling(source: Path, destination: Path) -> dict[str, Any]:
    target_part = MUTATIONS[source.suffix.lstrip(".").lower()]
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as original:
        original_infos = original.infolist()
        original_payloads = {info.filename: original.read(info.filename) for info in original_infos}
        target_payload = original_payloads[target_part]
        if target_payload.count(b"Hello") < 1:
            raise AssertionError(f"{source.name}: no Hello token in {target_part}")
        changed_payload = target_payload.replace(b"Hello", b"Jello", 1)
        with zipfile.ZipFile(destination, "w") as invalid:
            for source_info in original_infos:
                info = copy.copy(source_info)
                payload = (
                    changed_payload if source_info.filename == target_part else original_payloads[source_info.filename]
                )
                invalid.writestr(info, payload)

    with zipfile.ZipFile(destination) as invalid:
        invalid_infos = invalid.infolist()
        invalid_payloads = {info.filename: invalid.read(info.filename) for info in invalid_infos}
    assert [info.filename for info in invalid_infos] == [info.filename for info in original_infos]
    assert invalid_payloads[target_part] == changed_payload
    for name, payload in original_payloads.items():
        if name != target_part:
            assert invalid_payloads[name] == payload
    signature_names = [name for name in original_payloads if name.startswith("_xmlsignatures/")]
    assert signature_names
    assert all(invalid_payloads[name] == original_payloads[name] for name in signature_names)
    return {
        "path": str(destination),
        "bytes": destination.stat().st_size,
        "sha256": _sha256(destination),
        "mutated_part": target_part,
        "source_entry_order_preserved": True,
        "all_other_entry_payloads_preserved": True,
        "signature_graph_payloads_preserved": True,
    }


def _run_cli(
    executable: Path,
    arguments: list[str],
    *,
    timeout: int = 180,
) -> subprocess.CompletedProcess[str]:
    completed = subprocess.run(
        [str(executable), *arguments],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )
    if completed.returncode != 0:
        raise AssertionError(
            f"CLI failed ({completed.returncode}): {arguments}\nstdout={completed.stdout}\nstderr={completed.stderr}"
        )
    return completed


def _warning_codes(payload: dict[str, Any]) -> list[str]:
    return [str(item.get("code", "")) for item in payload.get("warnings", [])]


def _evaluate_conversion(
    executable: Path,
    source: Path,
    target: str,
    output_dir: Path,
    *,
    signature_material: bool,
) -> dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    before = _sha256(source)
    completed = _run_cli(
        executable,
        ["--json", "--yes", "run", str(source), "--to", target, "--output", str(output_dir)],
    )
    payload = json.loads(completed.stdout)
    assert payload["success"] is True
    output_path = Path(payload["data"]["output_file"])
    assert output_path.is_file() and output_path.stat().st_size > 0
    codes = _warning_codes(payload)
    expected = [VALIDATION_CODE, DERIVED_CODE] if signature_material else []
    assert codes == expected
    if signature_material:
        messages = {warning["code"]: warning["message"] for warning in payload["warnings"]}
        assert "intact and tampered inputs cannot be distinguished" in messages[VALIDATION_CODE]
        assert "derived and unsigned" in messages[DERIVED_CODE]
        assert "did not preserve or transfer" in messages[DERIVED_CODE]
    assert _sha256(source) == before
    return {
        "source": str(source),
        "target": target,
        "success": True,
        "artifact": str(output_path),
        "artifact_bytes": output_path.stat().st_size,
        "artifact_sha256": _sha256(output_path),
        "warning_codes": codes,
        "source_sha256_before": before,
        "source_sha256_after": _sha256(source),
        "stdout": completed.stdout,
        "stderr": completed.stderr,
    }


def _evaluate_inspect(
    executable: Path,
    source: Path,
    *,
    signature_material: bool,
) -> dict[str, Any]:
    before = _sha256(source)
    json_run = _run_cli(executable, ["--json", "inspect", str(source)])
    envelope = json.loads(json_run.stdout)
    data = envelope["data"]
    text_run = _run_cli(executable, ["inspect", str(source)])
    codes = [str(item.get("code", "")) for item in data["warnings"]]
    expected = [VALIDATION_CODE] if signature_material else []
    assert codes == expected
    assert data["ooxml_signature"]["state"] == ("complete" if signature_material else "unsigned")
    if signature_material:
        assert VALIDATION_CODE in text_run.stdout
        assert "intact and tampered inputs cannot be distinguished" in text_run.stdout
    else:
        assert VALIDATION_CODE not in text_run.stdout
    assert DERIVED_CODE not in text_run.stdout
    assert _sha256(source) == before
    return {
        "source": str(source),
        "json_warning_codes": codes,
        "signature_state": data["ooxml_signature"]["state"],
        "text_has_validation_code": VALIDATION_CODE in text_run.stdout,
        "text_has_derived_code": DERIVED_CODE in text_run.stdout,
        "source_sha256_before": before,
        "source_sha256_after": _sha256(source),
        "json_stdout": json_run.stdout,
        "text_stdout": text_run.stdout,
    }


def _evaluate_gui(
    cases: list[dict[str, Any]],
    markdown_results: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from docwen_core.models.artifact import ArtifactManifest
    from docwen_core.models.result import ConversionDiagnostic, ConversionResult
    from docwen_gui.app import create_main_window, create_qapplication

    app = create_qapplication(["policy01-gui-oracle"])
    window = create_main_window(initial_files=[str(case["path"]) for case in cases])
    app.processEvents()
    refs = {Path(ref.path).name: ref for ref in window.view_model.files}
    results: list[dict[str, Any]] = []
    try:
        for index, case in enumerate(cases):
            source = Path(case["path"])
            signature_material = bool(case["signature_material"])
            ref = refs[source.name]
            expected_admission = signature_material
            assert (VALIDATION_CODE in ref.warning_message) is expected_admission
            batch_entry = window._batch_list_vm.get_file_entry(str(source))
            assert batch_entry is not None
            assert (VALIDATION_CODE in str(batch_entry.warning_message or "")) is expected_admission

            conversion = markdown_results[case["id"]]
            response = json.loads(conversion["stdout"])
            diagnostics = [ConversionDiagnostic.from_dict(warning) for warning in response.get("warnings", [])]
            artifact_path = Path(conversion["artifact"])
            operation_id = f"policy01-gui-{index}"
            result = ConversionResult(
                task_id=operation_id,
                success=True,
                artifacts=[
                    ArtifactManifest(
                        artifact_id=operation_id,
                        kind="primary",
                        staging_path=str(artifact_path),
                        suggested_name=artifact_path.name,
                        is_primary=True,
                    )
                ],
                diagnostics=diagnostics,
            )
            window._on_execution_finished(
                result,
                {
                    "request_id": operation_id,
                    "file_path": str(source),
                    "file_paths": [str(source)],
                    "total_count": 1,
                    "batch": False,
                    "display_name": source.name,
                },
            )
            app.processEvents()
            rows = [row for row in window._info_area_vm.history_rows if row.operation_id == operation_id]
            warning_rows = [row for row in rows if row.message_type == "warning"]
            expected_warning_count = 3 if signature_material else 0
            # The completion row itself uses warning tone when diagnostics exist.
            assert len(warning_rows) == expected_warning_count
            warning_messages = [row.message for row in warning_rows]
            if signature_material:
                assert any(
                    "intact and tampered inputs cannot be distinguished" in message for message in warning_messages
                )
                assert any("derived and unsigned" in message for message in warning_messages)
            results.append(
                {
                    "case_id": case["id"],
                    "admission_warning": ref.warning_message,
                    "batch_admission_warning": batch_entry.warning_message,
                    "history_row_types": [row.message_type for row in rows],
                    "history_warning_messages": warning_messages,
                    "pass": True,
                }
            )
    finally:
        window.close()
        app.processEvents()
    return results


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--evidence-root", type=Path, required=True)
    args = parser.parse_args()

    repo_root = args.repo_root.resolve()
    source_root = args.source_root.resolve()
    evidence_root = args.evidence_root.resolve()
    executable = repo_root / ".venv" / "Scripts" / "docwen.exe"
    if not executable.is_file():
        raise FileNotFoundError(executable)
    evidence_root.mkdir(parents=True, exist_ok=True)
    if any(evidence_root.iterdir()):
        raise RuntimeError(f"Evidence root must be empty: {evidence_root}")

    owners = ("docx", "xlsx", "pptx")
    source_hashes_before: dict[str, str] = {}
    invalid: dict[str, dict[str, Any]] = {}
    cases: list[dict[str, Any]] = []
    for owner in owners:
        signed = source_root / f"hello-world-signed.{owner}"
        unsigned = source_root / f"hello-world-unsigned.{owner}"
        source_hashes_before[signed.name] = _sha256(signed)
        source_hashes_before[unsigned.name] = _sha256(unsigned)
        invalid_path = evidence_root / "invalid" / f"hello-world-invalid.{owner}"
        invalid[owner] = _make_invalid_sibling(signed, invalid_path)
        cases.extend(
            [
                {
                    "id": f"{owner}-signed",
                    "owner": owner,
                    "kind": "signed",
                    "path": signed,
                    "signature_material": True,
                },
                {
                    "id": f"{owner}-invalid",
                    "owner": owner,
                    "kind": "invalid",
                    "path": invalid_path,
                    "signature_material": True,
                },
                {
                    "id": f"{owner}-unsigned",
                    "owner": owner,
                    "kind": "unsigned",
                    "path": unsigned,
                    "signature_material": False,
                },
            ]
        )

    conversions: list[dict[str, Any]] = []
    markdown_results: dict[str, dict[str, Any]] = {}
    for case in cases:
        result = _evaluate_conversion(
            executable,
            Path(case["path"]),
            "md",
            evidence_root / "conversions" / case["id"] / "md",
            signature_material=bool(case["signature_material"]),
        )
        result["case_id"] = case["id"]
        conversions.append(result)
        markdown_results[case["id"]] = result
        if case["owner"] in {"docx", "xlsx"}:
            pdf_result = _evaluate_conversion(
                executable,
                Path(case["path"]),
                "pdf",
                evidence_root / "conversions" / case["id"] / "pdf",
                signature_material=bool(case["signature_material"]),
            )
            pdf_result["case_id"] = case["id"]
            conversions.append(pdf_result)

    inspect_results = [
        {
            "case_id": case["id"],
            **_evaluate_inspect(
                executable,
                Path(case["path"]),
                signature_material=bool(case["signature_material"]),
            ),
        }
        for case in cases
    ]
    gui_results = _evaluate_gui(cases, markdown_results)

    source_hashes_after = {name: _sha256(source_root / name) for name in source_hashes_before}
    assert source_hashes_after == source_hashes_before
    result = {
        "stage": "VIS-2026-07-23-202",
        "policy": "POLICY-01=B",
        "repo_head_at_start": subprocess.check_output(
            ["git", "rev-parse", "HEAD"],
            cwd=repo_root,
            text=True,
        ).strip(),
        "repo_index_tree_at_start": subprocess.check_output(
            ["git", "write-tree"],
            cwd=repo_root,
            text=True,
        ).strip(),
        "tested_repo_path_hashes": {relative: _sha256(repo_root / relative) for relative in TESTED_REPO_PATHS},
        "source_root": str(source_root),
        "evidence_root": str(evidence_root),
        "source_hashes_before": source_hashes_before,
        "source_hashes_after": source_hashes_after,
        "invalid_siblings": invalid,
        "conversion_slots": conversions,
        "inspect_slots": inspect_results,
        "gui_slots": gui_results,
        "counts": {
            "conversion_expected": 15,
            "conversion_passed": sum(item["success"] for item in conversions),
            "inspect_expected": 9,
            "inspect_passed": len(inspect_results),
            "gui_expected": 9,
            "gui_passed": sum(item["pass"] for item in gui_results),
        },
        "pass": (len(conversions) == 15 and len(inspect_results) == 9 and len(gui_results) == 9),
        "accepted_boundary": (
            "Presence-only structure detection cannot distinguish an intact "
            "signature from tampering and provides no signer, trust, timestamp, "
            "or revocation assurance."
        ),
    }
    assert result["pass"] is True
    _write_json(evidence_root / "result.json", result)
    print(json.dumps(result["counts"], ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
