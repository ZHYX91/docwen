from __future__ import annotations

import argparse
import json
import os
import re
import sys
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

if __package__ in {None, ""}:
    _BOOTSTRAP_ROOT = Path(__file__).resolve().parents[1]
    if str(_BOOTSTRAP_ROOT) not in sys.path:
        sys.path.insert(0, str(_BOOTSTRAP_ROOT))

from tools import workspace_cleanup
from tools.workspace_root import resolve_workspace_root

RECEIPT_SCHEMA = "docwen.acceptance-receipt.v1"
_TOKEN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$")
_SHA256 = re.compile(r"^[0-9a-f]{64}$")


class AcceptanceCloseoutError(ValueError):
    """A raw acceptance run could not be closed without weakening identity."""


def _atomic_json(path: Path, payload: dict[str, Any]) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    temporary.replace(path)


def _validate_token(value: str, *, field: str) -> str:
    if not _TOKEN.fullmatch(value):
        raise AcceptanceCloseoutError(f"invalid_{field}:{value}")
    return value


def _validate_subject(*, name: str, sha256: str, size: int) -> dict[str, object]:
    if not name or Path(name).name != name:
        raise AcceptanceCloseoutError(f"invalid_subject_name:{name}")
    if not _SHA256.fullmatch(sha256):
        raise AcceptanceCloseoutError(f"invalid_subject_sha256:{sha256}")
    if size < 0:
        raise AcceptanceCloseoutError(f"invalid_subject_bytes:{size}")
    return {"name": name, "bytes": size, "sha256": sha256}


def _managed_boundary(run_root: Path, *, workspace_root: Path) -> Path:
    for name in workspace_cleanup.MANAGED_ROOT_NAMES:
        boundary = (workspace_root / name).resolve(strict=True)
        if run_root == boundary or boundary in run_root.parents:
            if run_root == boundary:
                raise AcceptanceCloseoutError(f"managed_root_itself_not_a_run:{run_root}")
            return boundary
    raise AcceptanceCloseoutError(f"run_outside_managed_scratch:{run_root}")


def close_run(
    *,
    workspace_root: Path,
    run_root: Path,
    candidate_id: str,
    gate: str,
    subject_name: str,
    subject_sha256: str,
    subject_bytes: int,
    limitations: tuple[str, ...] = (),
) -> Path:
    workspace = resolve_workspace_root(Path(__file__).resolve().parents[1], explicit=workspace_root)
    candidate = _validate_token(candidate_id, field="candidate_id")
    gate_name = _validate_token(gate, field="gate")
    subject = _validate_subject(name=subject_name, sha256=subject_sha256, size=subject_bytes)
    run = Path(os.path.abspath(run_root)).resolve(strict=True)
    _managed_boundary(run, workspace_root=workspace)

    marker = run / workspace_cleanup.LEASE_NAME
    try:
        lease = json.loads(marker.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise AcceptanceCloseoutError(f"run_lease_unreadable:{run}") from error
    if not isinstance(lease, dict):
        raise AcceptanceCloseoutError(f"run_lease_not_object:{run}")
    if lease.get("root") != str(run) or not str(lease.get("owner", "")).startswith("docwen."):
        raise AcceptanceCloseoutError(f"run_lease_identity_mismatch:{run}")
    if str(lease.get("state", "")).casefold() not in workspace_cleanup.SUCCESS_STATES:
        raise AcceptanceCloseoutError(f"run_not_success_terminal:{run}:{lease.get('state')}")
    if workspace_cleanup._process_alive(lease.get("pid")):
        raise AcceptanceCloseoutError(f"run_owner_still_alive:{run}:{lease.get('pid')}")

    identity = workspace_cleanup._snapshot_tree(run)
    receipt = workspace / "acceptance" / f"{candidate}--{gate_name}.json"
    if receipt.exists():
        raise AcceptanceCloseoutError(f"receipt_exists:{receipt}")
    generated_at = datetime.now(UTC).isoformat().replace("+00:00", "Z")
    payload: dict[str, Any] = {
        "schema": RECEIPT_SCHEMA,
        "candidateId": candidate,
        "gate": gate_name,
        "result": "passed",
        "generatedAt": generated_at,
        "subject": subject,
        "runSummary": {
            "bytes": identity["bytes"],
            "files": identity["files"],
            "directories": identity["directories"],
            "contentSha256": identity["contentSha256"],
            "metadataSha256": identity["metadataSha256"],
        },
        "limitations": list(limitations),
        "rawRunRemoved": False,
    }

    stamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%S%fZ")
    diagnostics = workspace / "diagnostics"
    provisional = diagnostics / f"acceptance-closeout-{candidate}-{gate_name}-{stamp}.json"
    plan_path = diagnostics / f"acceptance-closeout-plan-{candidate}-{gate_name}-{stamp}.json"
    _atomic_json(provisional, payload)
    plan = workspace_cleanup.create_plan(
        workspace_root=workspace,
        explicit_targets=(run,),
        reason=f"acceptance closeout {candidate}/{gate_name}",
    )
    workspace_cleanup.save_plan(plan, plan_path)
    workspace_cleanup.apply_saved_plan(plan_path, workspace_root=workspace)
    payload["rawRunRemoved"] = True
    _atomic_json(receipt, payload)
    plan_path.unlink()
    provisional.unlink()
    return receipt


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--workspace-root", type=Path, required=True)
    result.add_argument("--run-root", type=Path, required=True)
    result.add_argument("--candidate-id", required=True)
    result.add_argument("--gate", required=True)
    result.add_argument("--subject-name", required=True)
    result.add_argument("--subject-sha256", required=True)
    result.add_argument("--subject-bytes", type=int, required=True)
    result.add_argument("--limitation", action="append", default=[])
    return result


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        receipt = close_run(
            workspace_root=args.workspace_root,
            run_root=args.run_root,
            candidate_id=args.candidate_id,
            gate=args.gate,
            subject_name=args.subject_name,
            subject_sha256=args.subject_sha256,
            subject_bytes=args.subject_bytes,
            limitations=tuple(args.limitation),
        )
    except (AcceptanceCloseoutError, workspace_cleanup.HousekeepingError, OSError, ValueError) as error:
        print(f"acceptance_closeout_error:{error}", file=sys.stderr)
        return 2
    print(f"acceptance_receipt:{receipt}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
